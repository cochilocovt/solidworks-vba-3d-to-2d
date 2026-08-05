<#
.SYNOPSIS
    Runs one authorized mutating production invocation of
    Module1_Main.main, replacing the manual deploy / VBE-compile /
    open-part / run-macro loop.

.DESCRIPTION
    1. Refuses outright unless -AllowMutation is present. That switch is
       the per-run authorization; it has no default and is never implied.
    2. Optionally redeploys src/baseline-model-dims into the target SWP.
    3. Attaches to the already-running SOLIDWORKS instance and opens the
       authorized fixture part. The part is opened READ-ONLY, so no macro
       behaviour can save a fixture.
    4. Calls Module20_ProbeRunner.R23_PrepareProductionRun, which compiles
       the whole VBA project through the VBE Compile control and activates
       the part. This is the programmatic form of the manual
       "Debug > Compile Project" gate: VBA compiles lazily, so a module
       that only fails when first called would otherwise abort a run after
       several views already exist.
    5. Reads the pre-flight verdict out of the evidence log. If it is not
       ready=True, the run STOPS here and main is never invoked.
    6. Invokes Module1_Main.main. main shows UserForm1 modally, so this
       call blocks until the operator dismisses the form and the pipeline
       finishes. That wait is intended; the timeout exists only so a
       forgotten dialog reports instead of hanging forever.
    7. Locates the QA report the run wrote and prints its stage table.

    WHAT THIS SCRIPT DOES NOT DO. It never saves the model, never saves
    the generated drawing, never closes a document, and never widens the
    authorized fixture list. The drawing main creates is unsaved and
    disposable; closing it without saving leaves nothing behind.

.PARAMETER AllowMutation
    Mandatory. Without it the script refuses. Present to make an
    accidental mutating run impossible: no other switch, default or
    environment setting can stand in for it.

.PARAMETER Deploy
    Redeploys src/baseline-model-dims into the target SWP first.

.PARAMETER PartPath
    Full path to the authorized fixture part. Defaults to
    P-0251-14A-001.SLDPRT. Refused if it is not one of the three paths
    Module1_Main.IsAuthorizedFixture accepts.

.PARAMETER TimeoutSeconds
    Seconds to wait for main to return. The default is generous because a
    human drives UserForm1. A timeout means the form is still open or a
    modal error dialog needs reading; it is not a failure verdict.

.PARAMETER ReportTailLines
    Number of trailing QA-report lines to print when the stage table
    cannot be located.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [switch] $AllowMutation,

    [Parameter()]
    [switch] $Deploy,

    [Parameter()]
    [string] $ManifestPath,

    [Parameter()]
    [string] $TargetSwp,

    [Parameter()]
    [string] $PartPath,

    [Parameter()]
    [ValidateRange(60, 3600)]
    [int] $TimeoutSeconds = 900,

    [Parameter()]
    [ValidateRange(1, 500)]
    [int] $ReportTailLines = 40
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if (-not $AllowMutation) {
    throw "Refusing to run: -AllowMutation was not supplied. This script " + `
        "invokes Module1_Main.main, which creates and edits a drawing."
}

# Mirrors Module1_Main.bas FIXTURE_1/FIXTURE_2/FIXTURE_3 exactly, the same
# way tools/probe-runner/Run-R23Probes.ps1 does. This script never widens
# the allowlist; if the VBA constants change, update this array to match.
$AuthorizedFixtureParts = @(
    'C:\Users\V.T\Documents\VBA 3D TO 2D\test_assets\models\P-0251-14A-001.SLDPRT'
    'C:\Users\V.T\Documents\VBA 3D TO 2D\test_assets\models\P-0252-01-001.SLDPRT'
    'C:\Users\V.T\Documents\VBA 3D TO 2D\test_assets\models\P-0252-01-013.SLDPRT'
)

function Test-AuthorizedFixturePart {
    param([Parameter(Mandatory)] [string] $Path)

    $normalized = $Path.Trim().Replace('/', '\').ToLowerInvariant()
    foreach ($fixture in $AuthorizedFixtureParts) {
        if ($normalized -eq $fixture.ToLowerInvariant()) {
            return $true
        }
    }
    return $false
}

if ([string]::IsNullOrWhiteSpace($PartPath)) {
    $PartPath = $AuthorizedFixtureParts[0]
}
$PartPath = [IO.Path]::GetFullPath($PartPath)

if (-not (Test-AuthorizedFixturePart -Path $PartPath)) {
    throw "Refusing an unauthorized fixture part: $PartPath. " + `
        "Only the three paths in Module1_Main.bas (FIXTURE_1..3) are " + `
        "permitted; this script does not widen that list."
}

if (-not (Test-Path -LiteralPath $PartPath -PathType Leaf)) {
    throw "Authorized fixture part does not exist on disk: $PartPath"
}

$workspaceRoot = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$evidenceRoot = Join-Path $workspaceRoot 'test_assets\iteration_evidence'
$probeRunRoot = Join-Path $evidenceRoot 'probe_runs'
$macroQaRoot = Join-Path $evidenceRoot 'macro_qa'

function Get-NewestChildDirectory {
    param([Parameter(Mandatory)] [string] $Root)

    if (-not (Test-Path -LiteralPath $Root -PathType Container)) { return $null }
    return Get-ChildItem -LiteralPath $Root -Directory |
        Sort-Object -Property Name -Descending |
        Select-Object -First 1
}

# Recorded before the run so a stale folder from an earlier session can
# never be mistaken for this run's evidence.
$qaBefore = Get-NewestChildDirectory -Root $macroQaRoot
$qaBeforeName = if ($null -eq $qaBefore) { '' } else { $qaBefore.Name }

$deployScript = [IO.Path]::GetFullPath(
    (Join-Path $PSScriptRoot '..\swp-deploy\Deploy-TargetSpecHybrid.ps1'))

if ($Deploy) {
    Write-Host 'Deploying src/baseline-model-dims before the production run...'
    $deployArgs = @{}
    if (-not [string]::IsNullOrWhiteSpace($ManifestPath)) {
        $deployArgs['ManifestPath'] = $ManifestPath
    }
    if (-not [string]::IsNullOrWhiteSpace($TargetSwp)) {
        $deployArgs['TargetSwp'] = $TargetSwp
    }
    & $deployScript @deployArgs
}

if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
    $ManifestPath = Join-Path $PSScriptRoot '..\swp-deploy\deployment-manifest.json'
}
$manifestFile = (Resolve-Path -LiteralPath $ManifestPath).Path
$manifestDirectory = Split-Path -Parent $manifestFile
$manifest = Get-Content -LiteralPath $manifestFile -Raw | ConvertFrom-Json

if ([string]::IsNullOrWhiteSpace($TargetSwp)) {
    $configuredTarget = [string] $manifest.targetSwp
    if ([IO.Path]::IsPathRooted($configuredTarget)) {
        $resolvedTarget = [IO.Path]::GetFullPath($configuredTarget)
    }
    else {
        $resolvedTarget = [IO.Path]::GetFullPath(
            (Join-Path $manifestDirectory $configuredTarget))
    }
}
else {
    $resolvedTarget = [IO.Path]::GetFullPath($TargetSwp)
}

if (-not (Test-Path -LiteralPath $resolvedTarget -PathType Leaf)) {
    throw "Target SWP does not exist: $resolvedTarget"
}

$invokerScript = [IO.Path]::GetFullPath(
    (Join-Path $PSScriptRoot '..\swp-deploy\Invoke-SolidWorksMacro.ps1'))
. $invokerScript

# Invokes one VBA procedure from a background STA runspace so a modal
# dialog yields a timeout report instead of an indefinite hang. The
# runspace re-acquires SOLIDWORKS from the Running Object Table itself: an
# STA COM proxy created on one thread is not safe to hand to another.
function Invoke-MacroWithTimeout {
    param(
        [Parameter(Mandatory)] [string] $MacroPath,
        [Parameter(Mandatory)] [string] $ModuleName,
        [Parameter(Mandatory)] [string] $ProcedureName,
        [Parameter(Mandatory)] [string] $InvokerScriptPath,
        [Parameter(Mandatory)] [int] $Seconds
    )

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = [Threading.ApartmentState]::STA
    $runspace.ThreadOptions =
        [Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $runspace.Open()

    $ps = [powershell]::Create()
    $ps.Runspace = $runspace

    [void] $ps.AddScript({
        param($InvokerScriptPath, $MacroPath, $ModuleName, $ProcedureName)

        . $InvokerScriptPath

        $sw = [Runtime.InteropServices.Marshal]::GetActiveObject(
            'SldWorks.Application')
        try {
            Invoke-SolidWorksMacro `
                -SolidWorks $sw `
                -MacroPath $MacroPath `
                -ModuleName $ModuleName `
                -ProcedureName $ProcedureName
        }
        finally {
            [void] [Runtime.InteropServices.Marshal]::ReleaseComObject($sw)
        }
    })
    [void] $ps.AddArgument($InvokerScriptPath)
    [void] $ps.AddArgument($MacroPath)
    [void] $ps.AddArgument($ModuleName)
    [void] $ps.AddArgument($ProcedureName)

    $asyncResult = $ps.BeginInvoke()
    $signalled = $asyncResult.AsyncWaitHandle.WaitOne(
        [TimeSpan]::FromSeconds($Seconds))

    if (-not $signalled) {
        return [pscustomobject]@{ Returned = $false; Error = 'Timeout' }
    }

    try {
        $ps.EndInvoke($asyncResult)
        $result = [pscustomobject]@{ Returned = $true; Error = '' }
    }
    catch {
        $result = [pscustomobject]@{
            Returned = $true; Error = $_.Exception.Message
        }
    }
    $ps.Dispose()
    $runspace.Close()
    return $result
}

$solidWorks = $null
try {
    $solidWorks = [Runtime.InteropServices.Marshal]::GetActiveObject(
        'SldWorks.Application')
}
catch {
    throw 'No running SOLIDWORKS instance is available. Start SOLIDWORKS and retry.'
}

# swDocumentTypes_e.swDocPART = 1 and
# swOpenDocOptions_e.swOpenDocOptions_ReadOnly = 2, both already relied on
# by tools/probe-runner/Run-R23Probes.ps1 and src/baseline-model-dims.
# The read-only handle is an extra safety net on top of the fixture guard:
# Agents.md forbids saving a fixture, and this makes it impossible
# regardless of macro behaviour.
$swDocPART = 1
$swOpenDocOptions_ReadOnly = 2

$preflightRun = $null
try {
    Write-Host "Opening authorized part read-only: $PartPath"
    $partOpenResult = Open-SolidWorksDocument `
        -SolidWorks $solidWorks `
        -FileName $PartPath `
        -DocumentType $swDocPART `
        -Options $swOpenDocOptions_ReadOnly
    Write-Host ("Part opened: $($partOpenResult.Title) " +
        "(warnings=$($partOpenResult.Warnings))")

    Write-Host 'Compiling the VBA project and activating the part...'
    $preflight = Invoke-MacroWithTimeout `
        -MacroPath $resolvedTarget `
        -ModuleName 'Module20_ProbeRunner' `
        -ProcedureName 'R23_PrepareProductionRun' `
        -InvokerScriptPath $invokerScript `
        -Seconds 300

    if (-not $preflight.Returned) {
        throw "Pre-flight did not return within 300 seconds. A modal VBE " + `
            "compile-error dialog most likely needs reading and dismissing " + `
            "in the SOLIDWORKS/VBE window. main was NOT invoked."
    }
    if (-not [string]::IsNullOrWhiteSpace($preflight.Error)) {
        throw "Pre-flight invocation failed: $($preflight.Error). " + `
            "main was NOT invoked."
    }

    $preflightRun = Get-NewestChildDirectory -Root $probeRunRoot
    if ($null -eq $preflightRun) {
        throw "Pre-flight wrote no evidence folder under $probeRunRoot. " + `
            "main was NOT invoked."
    }

    $preflightLog = Join-Path $preflightRun.FullName 'probe_log.txt'
    $preflightEnd = Select-String -LiteralPath $preflightLog `
        -Pattern '^R23_PREFLIGHT_END\|' | Select-Object -Last 1

    if ($null -eq $preflightEnd) {
        throw "Pre-flight log has no R23_PREFLIGHT_END line: $preflightLog. " + `
            "main was NOT invoked."
    }

    Write-Host $preflightEnd.Line

    if ($preflightEnd.Line -notmatch 'ready=True') {
        Write-Host ''
        Write-Host "Pre-flight is not ready. main was NOT invoked."
        Write-Host "Evidence: $preflightLog"
        Get-Content -LiteralPath $preflightLog -Tail 20
        return
    }

    Write-Host ''
    Write-Host 'Invoking Module1_Main.main (MUTATING).'
    Write-Host ('UserForm1 is modal: choose the settings in SOLIDWORKS and ' +
        'confirm. This waits until the pipeline finishes.')

    $mainResult = Invoke-MacroWithTimeout `
        -MacroPath $resolvedTarget `
        -ModuleName 'Module1_Main' `
        -ProcedureName 'main' `
        -InvokerScriptPath $invokerScript `
        -Seconds $TimeoutSeconds

    if (-not $mainResult.Returned) {
        Write-Warning ("main did not return within $TimeoutSeconds seconds. " +
            "UserForm1 or a message box is most likely still open. This is " +
            "not a failure verdict: finish the dialog, then read the QA " +
            "report folder printed below.")
    }
    elseif (-not [string]::IsNullOrWhiteSpace($mainResult.Error)) {
        Write-Warning "main invocation reported: $($mainResult.Error)"
    }
    else {
        Write-Host 'main returned.'
    }
}
finally {
    if ($null -ne $solidWorks) {
        [void] [Runtime.InteropServices.Marshal]::ReleaseComObject($solidWorks)
    }
}

$qaAfter = Get-NewestChildDirectory -Root $macroQaRoot

if ($null -eq $qaAfter) {
    Write-Warning "No QA report folder exists under $macroQaRoot."
    return
}

if ($qaAfter.Name -eq $qaBeforeName) {
    Write-Warning ("No new QA report was written. The newest folder is " +
        "still $($qaAfter.Name), which predates this run. Either the form " +
        "was cancelled, or the run has not reached evidence emission yet.")
    return
}

$reportPath = Join-Path $qaAfter.FullName 'QA_REPORT.txt'
if (-not (Test-Path -LiteralPath $reportPath -PathType Leaf)) {
    Write-Warning "Expected QA report was not found: $reportPath"
    return
}

Write-Host ''
Write-Host "QA report: $reportPath"

# The QA report is written as UTF-16. Reading it with the default encoding
# yields NUL-separated characters, which is why the stage table is selected
# explicitly rather than tailed.
$reportLines = Get-Content -LiteralPath $reportPath
$stageLines = $reportLines | Select-String -Pattern '^- STAGE\|'

if ($null -eq $stageLines -or $stageLines.Count -eq 0) {
    Write-Host "Last $ReportTailLines line(s):"
    $reportLines | Select-Object -Last $ReportTailLines
    return
}

Write-Host ''
Write-Host 'REQUIRED STAGES'
foreach ($line in $stageLines) { Write-Host $line.Line }

$counts = $reportLines | Select-String `
    -Pattern '^(Views created|Annotations imported|Visible display dimensions|Horizontal ordinate groups|Vertical ordinate groups|Layout moves):'

if ($null -ne $counts) {
    Write-Host ''
    Write-Host 'COUNTS'
    foreach ($line in $counts) { Write-Host $line.Line }
}

Write-Host ''
Write-Host ('The drawing this run created is UNSAVED and disposable. ' +
    'Inspect it visually, then close it without saving.')
