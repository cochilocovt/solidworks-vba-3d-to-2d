<#
.SYNOPSIS
    Runs the nine read-only R23_Probe* entry points unattended and prints
    the evidence log path, replacing the manual deploy/compile/open/run/
    paste loop described in docs/R23_PROBE_AUTOMATION_IMPLEMENTATION_PLAN.md.

.DESCRIPTION
    1. Optionally redeploys src/baseline-model-dims into the target SWP.
    2. Attaches to the already-running SOLIDWORKS instance.
    3. Opens the authorized fixture part, then the authorized drawing, via
       ISldWorks.OpenDoc6 -- part first, per the 2026-08-02 finding that a
       drawing bound to the V: network sibling rebinds to the local
       test_assets copy only if that copy is already open when the drawing
       loads.
    4. Invokes Module20_ProbeRunner.R23_RunAllProbes through the strongly
       typed SolidWorksMacroInvoker inside a background runspace, so a
       modal compile-error dialog yields a timeout report instead of an
       indefinite hang.
    5. Prints the evidence log path and its last lines.

    Contains no allowMutation switch and calls no mutating procedure. This
    script only automates the read-only probe iteration loop; production
    acceptance keeps its manual full-project VBA compile gate.

.PARAMETER Deploy
    Redeploys src/baseline-model-dims into the target SWP first by
    calling tools/swp-deploy/Deploy-TargetSpecHybrid.ps1 before probing.

.PARAMETER PartPath
    Full path to the authorized fixture part to open first. Defaults to
    P-0251-14A-001.SLDPRT. Refused if it is not one of the three paths
    Module1_Main.IsAuthorizedFixture accepts -- this script never widens
    that allowlist.

.PARAMETER DrawingPath
    Full path to the drawing to open after the part. Mandatory: no
    authoritative default exists in this repository (the fixture drawings
    live outside it, per finding 5.6 in the implementation plan), and
    guessing one would violate the "never guess a path/contract" rule in
    Agents.md. The true authorization gate is still enforced live, inside
    each drawing-scoped probe, against the part each view actually
    references.

.PARAMETER TimeoutSeconds
    Seconds to wait for R23_RunAllProbes to return before reporting a
    suspected modal dialog instead of hanging. RunMacro2 is synchronous
    (finding 5.4), so this timeout is enforced from a background runspace,
    not from RunMacro2 itself.

.PARAMETER LogTailLines
    Number of trailing lines of the evidence log to print after the run.
#>
[CmdletBinding()]
param(
    [Parameter()]
    [switch] $Deploy,

    [Parameter()]
    [string] $ManifestPath,

    [Parameter()]
    [string] $TargetSwp,

    [Parameter()]
    [string] $PartPath,

    [Parameter(Mandatory)]
    [string] $DrawingPath,

    [Parameter()]
    [ValidateRange(15, 300)]
    [int] $TimeoutSeconds = 120,

    [Parameter()]
    [ValidateRange(1, 500)]
    [int] $LogTailLines = 40
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Mirrors Module1_Main.bas FIXTURE_1/FIXTURE_2/FIXTURE_3 exactly. This
# script never widens the allowlist; if the VBA constants ever change,
# update this array to match, do not add a fourth entry independently.
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

$DrawingPath = [IO.Path]::GetFullPath($DrawingPath)
if (-not (Test-Path -LiteralPath $DrawingPath -PathType Leaf)) {
    throw "Drawing does not exist on disk: $DrawingPath"
}

$deployScript = Join-Path $PSScriptRoot '..\swp-deploy\Deploy-TargetSpecHybrid.ps1'
$deployScript = [IO.Path]::GetFullPath($deployScript)

if ($Deploy) {
    Write-Host 'Deploying src/baseline-model-dims before probing...'
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

# Same strongly typed RunMacro2 invoker the guarded deployer uses --
# see tools/swp-deploy/Invoke-SolidWorksMacro.ps1.
. (Join-Path $PSScriptRoot '..\swp-deploy\Invoke-SolidWorksMacro.ps1')

$solidWorks = $null
try {
    $solidWorks = [Runtime.InteropServices.Marshal]::GetActiveObject('SldWorks.Application')
}
catch {
    throw 'No running SOLIDWORKS instance is available. Start SOLIDWORKS and retry.'
}

# swDocumentTypes_e: swDocPART = 1, swDocDRAWING = 3. Both verified
# against the installed SOLIDWORKS 2025 SP1.2 type library and already in
# use throughout src/baseline-model-dims (e.g. Module1_Main.bas).
$swDocPART = 1
$swDocDRAWING = 3

# swOpenDocOptions_e.swOpenDocOptions_ReadOnly = 2. Opening both documents
# read-only is an extra safety net on top of the fixture guard and the
# probes' own creations=0/mutations=0 evidence: Agents.md forbids saving a
# fixture, and a read-only handle makes that impossible regardless of
# macro behaviour.
$swOpenDocOptions_ReadOnly = 2

try {
    Write-Host "Opening authorized part: $PartPath"
    $partOpenResult = Open-SolidWorksDocument `
        -SolidWorks $solidWorks `
        -FileName $PartPath `
        -DocumentType $swDocPART `
        -Options $swOpenDocOptions_ReadOnly
    Write-Host "Part opened: $($partOpenResult.Title) (warnings=$($partOpenResult.Warnings))"

    Write-Host "Opening drawing: $DrawingPath"
    $drawingOpenResult = Open-SolidWorksDocument `
        -SolidWorks $solidWorks `
        -FileName $DrawingPath `
        -DocumentType $swDocDRAWING `
        -Options $swOpenDocOptions_ReadOnly
    Write-Host "Drawing opened: $($drawingOpenResult.Title) (warnings=$($drawingOpenResult.Warnings))"

    Write-Host 'Invoking Module20_ProbeRunner.R23_RunAllProbes...'

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = [Threading.ApartmentState]::STA
    $runspace.ThreadOptions = [Management.Automation.Runspaces.PSThreadOptions]::ReuseThread
    $runspace.Open()

    $ps = [powershell]::Create()
    $ps.Runspace = $runspace

    # The runspace re-acquires the SOLIDWORKS COM object from the Running
    # Object Table itself rather than receiving $solidWorks from this
    # thread: an STA COM proxy created on one thread is not safe to hand
    # directly to another without marshaling, and GetActiveObject is the
    # already-proven way this repo's tooling attaches to SOLIDWORKS.
    [void] $ps.AddScript({
        param($InvokerScriptPath, $TargetSwp)

        . $InvokerScriptPath

        $sw = [Runtime.InteropServices.Marshal]::GetActiveObject('SldWorks.Application')
        try {
            Invoke-SolidWorksMacro `
                -SolidWorks $sw `
                -MacroPath $TargetSwp `
                -ModuleName 'Module20_ProbeRunner' `
                -ProcedureName 'R23_RunAllProbes'
        }
        finally {
            [void] [Runtime.InteropServices.Marshal]::ReleaseComObject($sw)
        }
    })
    $invokerScript = Join-Path $PSScriptRoot '..\swp-deploy\Invoke-SolidWorksMacro.ps1'
    [void] $ps.AddArgument($invokerScript)
    [void] $ps.AddArgument($resolvedTarget)

    $asyncResult = $ps.BeginInvoke()
    $signalled = $asyncResult.AsyncWaitHandle.WaitOne(
        [TimeSpan]::FromSeconds($TimeoutSeconds))

    if (-not $signalled) {
        Write-Warning ("R23_RunAllProbes did not return within " + `
            "$TimeoutSeconds seconds. compile=DialogSuspected -- a " + `
            "modal compile-error dialog most likely needs a human to " + `
            "read it and dismiss it in the SOLIDWORKS/VBE window " + `
            "(finding 5.5). The background runspace is left running; " + `
            "re-run this script once the dialog is dismissed.")
    }
    else {
        try {
            $ps.EndInvoke($asyncResult)
            Write-Host 'R23_RunAllProbes returned.'
        }
        catch {
            Write-Warning "RunMacro2 invocation failed: $($_.Exception.Message)"
        }
        $ps.Dispose()
        $runspace.Close()
    }
}
finally {
    if ($null -ne $solidWorks) {
        [void] [Runtime.InteropServices.Marshal]::ReleaseComObject($solidWorks)
    }
}

$evidenceRoot = Join-Path (Split-Path -Parent (Split-Path -Parent $PartPath)) `
    'iteration_evidence\probe_runs'
if (-not (Test-Path -LiteralPath $evidenceRoot -PathType Container)) {
    Write-Warning "No probe_runs evidence folder found under $evidenceRoot"
    return
}

$latestRun = Get-ChildItem -LiteralPath $evidenceRoot -Directory |
    Sort-Object -Property Name -Descending |
    Select-Object -First 1

if ($null -eq $latestRun) {
    Write-Warning "No probe run folders found under $evidenceRoot"
    return
}

$logPath = Join-Path $latestRun.FullName 'probe_log.txt'
if (-not (Test-Path -LiteralPath $logPath -PathType Leaf)) {
    Write-Warning "Expected log file was not found: $logPath"
    return
}

Write-Host ''
Write-Host "Evidence log: $logPath"
Write-Host "Last $LogTailLines line(s):"
Get-Content -LiteralPath $logPath -Tail $LogTailLines
