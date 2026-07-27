[CmdletBinding()]
param(
    [Parameter()]
    [string] $ManifestPath,

    [Parameter()]
    [string] $TargetSwp,

    [Parameter()]
    [switch] $PreflightOnly,

    [Parameter()]
    [ValidateRange(15, 300)]
    [int] $TimeoutSeconds = 120
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
    $ManifestPath = Join-Path $PSScriptRoot 'deployment-manifest.json'
}

function Resolve-ConfiguredPath {
    param(
        [Parameter(Mandatory)] [string] $BaseDirectory,
        [Parameter(Mandatory)] [string] $ConfiguredPath
    )

    if ([IO.Path]::IsPathRooted($ConfiguredPath)) {
        return [IO.Path]::GetFullPath($ConfiguredPath)
    }

    return [IO.Path]::GetFullPath((Join-Path $BaseDirectory $ConfiguredPath))
}

function Test-ExclusiveFileAccess {
    param([Parameter(Mandatory)] [string] $Path)

    $stream = $null
    try {
        $stream = [IO.File]::Open(
            $Path,
            [IO.FileMode]::Open,
            [IO.FileAccess]::ReadWrite,
            [IO.FileShare]::None)
        return $true
    }
    catch {
        return $false
    }
    finally {
        if ($null -ne $stream) {
            $stream.Dispose()
        }
    }
}

function Wait-ForFile {
    param(
        [Parameter(Mandatory)] [string] $Path,
        [Parameter(Mandatory)] [int] $Timeout
    )

    $deadline = [DateTime]::UtcNow.AddSeconds($Timeout)
    while ([DateTime]::UtcNow -lt $deadline) {
        if (Test-Path -LiteralPath $Path -PathType Leaf) {
            return
        }
        Start-Sleep -Milliseconds 250
    }

    throw "Timed out waiting for deployment evidence: $Path"
}

function Invoke-SolidWorksMacro {
    param(
        [Parameter(Mandatory)] [object] $SolidWorks,
        [Parameter(Mandatory)] [string] $MacroPath,
        [Parameter(Mandatory)] [string] $ModuleName,
        [Parameter(Mandatory)] [string] $ProcedureName
    )

    $programFiles = [Environment]::GetFolderPath(
        [Environment+SpecialFolder]::ProgramFiles)
    $interopPath = Join-Path $programFiles `
        'SOLIDWORKS Corp\SOLIDWORKS\api\redist\SolidWorks.Interop.sldworks.dll'

    $interopType = 'SolidWorks.Interop.sldworks.ISldWorks' -as [type]
    if ($null -eq $interopType) {
        if (-not (Test-Path -LiteralPath $interopPath -PathType Leaf)) {
            throw "SOLIDWORKS interop assembly was not found: $interopPath"
        }

        Add-Type -Path $interopPath
        $interopType = 'SolidWorks.Interop.sldworks.ISldWorks' -as [type]
    }

    if ($null -eq $interopType) {
        throw 'SOLIDWORKS ISldWorks interop type could not be loaded.'
    }

    $invokerType = 'TargetSpecHybrid.SwpDeployment.SolidWorksMacroInvoker' -as [type]
    if ($null -eq $invokerType) {
        $invokerSource = Join-Path $PSScriptRoot 'SolidWorksMacroInvoker.cs'
        if (-not (Test-Path -LiteralPath $invokerSource -PathType Leaf)) {
            throw "SOLIDWORKS macro invoker source was not found: $invokerSource"
        }

        Add-Type `
            -Path $invokerSource `
            -ReferencedAssemblies $interopPath
        $invokerType = `
            'TargetSpecHybrid.SwpDeployment.SolidWorksMacroInvoker' -as [type]
    }

    if ($null -eq $invokerType) {
        throw 'The strongly typed SOLIDWORKS macro invoker could not be loaded.'
    }

    $unloadAfterRun = 1
    $runResult = $invokerType::Run(
        $SolidWorks,
        $MacroPath,
        $ModuleName,
        $ProcedureName,
        $unloadAfterRun)

    if (-not $runResult.Succeeded) {
        throw "RunMacro2 failed for $ModuleName.$ProcedureName with swRunMacroError_e value $($runResult.Error)."
    }
}

$manifestFile = (Resolve-Path -LiteralPath $ManifestPath).Path
$manifestDirectory = Split-Path -Parent $manifestFile
$manifest = Get-Content -LiteralPath $manifestFile -Raw | ConvertFrom-Json

$sourceDirectory = Resolve-ConfiguredPath `
    -BaseDirectory $manifestDirectory `
    -ConfiguredPath ([string] $manifest.sourceDirectory)

if ([string]::IsNullOrWhiteSpace($TargetSwp)) {
    $resolvedTarget = Resolve-ConfiguredPath `
        -BaseDirectory $manifestDirectory `
        -ConfiguredPath ([string] $manifest.targetSwp)
}
else {
    $resolvedTarget = [IO.Path]::GetFullPath($TargetSwp)
}

if (-not (Test-Path -LiteralPath $resolvedTarget -PathType Leaf)) {
    throw "Target SWP does not exist: $resolvedTarget"
}

$sourceFiles = @()
foreach ($component in $manifest.components) {
    $sourcePath = Join-Path $sourceDirectory ([string] $component.file)
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Managed source file is missing: $sourcePath"
    }

    $sourceText = Get-Content -LiteralPath $sourcePath -Raw
    if ($sourceText -match '(?im)^\s*(?:ï»¿)?Attribute\s+') {
        throw "VBA Attribute metadata is not allowed in deployable source: $sourcePath"
    }
    if ($sourceText -notmatch '(?im)^Option\s+Explicit\s*$') {
        throw "Option Explicit is missing from $sourcePath"
    }

    $sourceFiles += [pscustomobject]@{
        Name = [string] $component.name
        Kind = [string] $component.kind
        Path = [IO.Path]::GetFullPath($sourcePath)
    }
}

$revisionSource = Join-Path $sourceDirectory ([string] $manifest.revisionSourceFile)
$revisionText = Get-Content -LiteralPath $revisionSource -Raw
$revisionMatch = [Regex]::Match($revisionText, [string] $manifest.revisionPattern)
if (-not $revisionMatch.Success) {
    throw "Could not determine the expected macro revision from $revisionSource"
}
$expectedRevision = $revisionMatch.Value

$pythonPath = 'C:\Users\V.T\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe'
$verifyScript = Join-Path $PSScriptRoot 'verify_swp_vba.py'
if (-not (Test-Path -LiteralPath $pythonPath -PathType Leaf)) {
    throw "Bundled Codex Python is unavailable: $pythonPath"
}

$preflightRoot = Join-Path $env:TEMP 'target-spec-swp-deploy-preflight'
New-Item -ItemType Directory -Path $preflightRoot -Force | Out-Null
$inventoryJson = Join-Path $preflightRoot 'inventory.json'
& $pythonPath $verifyScript `
    --manifest $manifestFile `
    --swp $resolvedTarget `
    --inventory-only `
    --json-output $inventoryJson
if ($LASTEXITCODE -ne 0) {
    throw "Unable to inventory the target SWP."
}

$inventory = Get-Content -LiteralPath $inventoryJson -Raw | ConvertFrom-Json
$bootstrapModule = [string] $manifest.bootstrapModule
$hasBootstrap = @($inventory.embeddedComponents) -contains $bootstrapModule

Write-Host "Target SWP: $resolvedTarget"
Write-Host "Source directory: $sourceDirectory"
Write-Host "Expected revision: $expectedRevision"
Write-Host "Managed components: $($sourceFiles.Count)"
Write-Host "Bootstrap present: $hasBootstrap"

if ($PreflightOnly) {
    if (-not $hasBootstrap) {
        Write-Warning "One-time bootstrap import is still required: $PSScriptRoot\Module0_SourceDeployment.bas"
    }
    return
}

if (-not $hasBootstrap) {
    throw "The target SWP does not contain $bootstrapModule. Perform the one-time bootstrap import described in $PSScriptRoot\README.md."
}

if (-not (Test-ExclusiveFileAccess -Path $resolvedTarget)) {
    throw "The target SWP is open or locked. Close Fable.swp in the VBA editor, but leave SOLIDWORKS running, and retry."
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$workspaceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$evidenceDirectory = Join-Path $workspaceRoot "test_assets\iteration_evidence\swp_deployment\$timestamp"
New-Item -ItemType Directory -Path $evidenceDirectory -Force | Out-Null

$originalBackup = Join-Path $evidenceDirectory 'original.swp'
$candidateInput = Join-Path $evidenceDirectory 'candidate-input.swp'
$candidateOutput = Join-Path $evidenceDirectory 'candidate-output.swp'
$deploymentResult = Join-Path $evidenceDirectory 'deployment-result.txt'
$compileResult = Join-Path $evidenceDirectory 'compile-result.txt'
$verificationJson = Join-Path $evidenceDirectory 'verification.json'
$postPromotionVerificationJson = Join-Path $evidenceDirectory 'post-promotion-verification.json'
$promotionBackup = Join-Path $evidenceDirectory 'pre-promotion.swp'

Copy-Item -LiteralPath $resolvedTarget -Destination $originalBackup
Copy-Item -LiteralPath $resolvedTarget -Destination $candidateInput

$requestPath = Join-Path $PSScriptRoot 'deployment-request.txt'
$requestLines = @(
    'SCHEMA=1'
    "EXPECTED_PROJECT=$($manifest.projectName)"
    "EXPECTED_REVISION=$expectedRevision"
    "INPUT_SWP=$candidateInput"
    "OUTPUT_SWP=$candidateOutput"
    "RESULT_FILE=$deploymentResult"
    "COMPILE_RESULT_FILE=$compileResult"
)
foreach ($component in $sourceFiles) {
    $requestLines += "COMPONENT=$($component.Name)|$($component.Kind)|$($component.Path)"
}
[IO.File]::WriteAllLines($requestPath, $requestLines, [Text.Encoding]::Default)

$solidWorks = $null
try {
    $solidWorks = [Runtime.InteropServices.Marshal]::GetActiveObject('SldWorks.Application')
}
catch {
    throw "No running SOLIDWORKS instance is available. Start SOLIDWORKS and retry."
}

try {
    Invoke-SolidWorksMacro `
        -SolidWorks $solidWorks `
        -MacroPath $candidateInput `
        -ModuleName $bootstrapModule `
        -ProcedureName ([string] $manifest.bootstrapProcedure)

    Wait-ForFile -Path $deploymentResult -Timeout $TimeoutSeconds
    $deploymentText = Get-Content -LiteralPath $deploymentResult -Raw
    if ($deploymentText -notmatch 'DEPLOYMENT\|status=SUCCESS') {
        throw "Candidate deployment failed. See $deploymentResult"
    }

    if (Test-Path -LiteralPath $candidateOutput) {
        throw "Candidate output unexpectedly exists before promotion copy: $candidateOutput"
    }
    Copy-Item -LiteralPath $candidateInput -Destination $candidateOutput

    if (-not (Test-Path -LiteralPath $candidateOutput -PathType Leaf)) {
        throw "The saved candidate could not be copied to candidate-output.swp."
    }

    Invoke-SolidWorksMacro `
        -SolidWorks $solidWorks `
        -MacroPath $candidateOutput `
        -ModuleName $bootstrapModule `
        -ProcedureName ([string] $manifest.compileProbeProcedure)

    Wait-ForFile -Path $compileResult -Timeout $TimeoutSeconds
    $compileText = Get-Content -LiteralPath $compileResult -Raw
    if ($compileText -notmatch 'COMPILE_PROBE\|status=SUCCESS') {
        throw "The deployed project did not pass the compile probe. See $compileResult"
    }
}
finally {
    if ($null -ne $solidWorks) {
        [void] [Runtime.InteropServices.Marshal]::ReleaseComObject($solidWorks)
    }
}

& $pythonPath $verifyScript `
    --manifest $manifestFile `
    --swp $candidateOutput `
    --json-output $verificationJson
if ($LASTEXITCODE -ne 0) {
    throw "Candidate read-back verification failed. See $verificationJson"
}

if (-not (Test-ExclusiveFileAccess -Path $resolvedTarget)) {
    throw "The target SWP became locked before promotion. The verified candidate remains at $candidateOutput"
}

$promotionTemp = "$resolvedTarget.deploying"
if (Test-Path -LiteralPath $promotionTemp) {
    Remove-Item -LiteralPath $promotionTemp -Force
}
Copy-Item -LiteralPath $candidateOutput -Destination $promotionTemp
[IO.File]::Replace($promotionTemp, $resolvedTarget, $promotionBackup, $true)

& $pythonPath $verifyScript `
    --manifest $manifestFile `
    --swp $resolvedTarget `
    --json-output $postPromotionVerificationJson
if ($LASTEXITCODE -ne 0) {
    $failedPromotion = Join-Path $evidenceDirectory 'failed-promoted-target.swp'
    Copy-Item -LiteralPath $resolvedTarget -Destination $failedPromotion

    $restoreTemp = "$resolvedTarget.restoring"
    if (Test-Path -LiteralPath $restoreTemp) {
        Remove-Item -LiteralPath $restoreTemp -Force
    }
    Copy-Item -LiteralPath $promotionBackup -Destination $restoreTemp
    [IO.File]::Replace($restoreTemp, $resolvedTarget, $null, $true)
    throw "Post-promotion verification failed; the original SWP was restored. See $postPromotionVerificationJson"
}

Write-Host "Deployment complete: $resolvedTarget"
Write-Host "Embedded revision: $expectedRevision"
Write-Host "Evidence: $evidenceDirectory"
