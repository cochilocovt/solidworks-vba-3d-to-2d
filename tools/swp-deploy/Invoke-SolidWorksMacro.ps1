function Get-SolidWorksInteropPath {
    $programFiles = [Environment]::GetFolderPath(
        [Environment+SpecialFolder]::ProgramFiles)
    return Join-Path $programFiles `
        'SOLIDWORKS Corp\SOLIDWORKS\api\redist\SolidWorks.Interop.sldworks.dll'
}

function Get-SolidWorksInteropType {
    $interopType = 'SolidWorks.Interop.sldworks.ISldWorks' -as [type]
    if ($null -ne $interopType) {
        return $interopType
    }

    $interopPath = Get-SolidWorksInteropPath
    if (-not (Test-Path -LiteralPath $interopPath -PathType Leaf)) {
        throw "SOLIDWORKS interop assembly was not found: $interopPath"
    }

    Add-Type -Path $interopPath
    $interopType = 'SolidWorks.Interop.sldworks.ISldWorks' -as [type]
    if ($null -eq $interopType) {
        throw 'SOLIDWORKS ISldWorks interop type could not be loaded.'
    }

    return $interopType
}

function Invoke-SolidWorksMacro {
    param(
        [Parameter(Mandatory)] [object] $SolidWorks,
        [Parameter(Mandatory)] [string] $MacroPath,
        [Parameter(Mandatory)] [string] $ModuleName,
        [Parameter(Mandatory)] [string] $ProcedureName
    )

    [void] (Get-SolidWorksInteropType)
    $interopPath = Get-SolidWorksInteropPath

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

# PowerShell's native COM late binding fails ISldWorks.OpenDoc6 with
# TYPE_E_ELEMENTNOTFOUND (it cannot resolve the out-parameter overload),
# and a direct bracket-cast of the raw GetActiveObject() proxy to
# ISldWorks fails too, even after Add-Type has loaded the interop
# assembly -- both confirmed live on 2026-08-04, see
# docs/SOLIDWORKS_API_VALIDATION.md. Routing through compiled C#
# (SolidWorksDocumentOpener.cs), the same pattern already proved for
# RunMacro2 below, is what actually works.
function Open-SolidWorksDocument {
    param(
        [Parameter(Mandatory)] [object] $SolidWorks,
        [Parameter(Mandatory)] [string] $FileName,
        [Parameter(Mandatory)] [int] $DocumentType,
        [Parameter(Mandatory)] [int] $Options
    )

    [void] (Get-SolidWorksInteropType)
    $interopPath = Get-SolidWorksInteropPath

    $openerType = 'TargetSpecHybrid.SwpDeployment.SolidWorksDocumentOpener' -as [type]
    if ($null -eq $openerType) {
        $openerSource = Join-Path $PSScriptRoot 'SolidWorksDocumentOpener.cs'
        if (-not (Test-Path -LiteralPath $openerSource -PathType Leaf)) {
            throw "SOLIDWORKS document opener source was not found: $openerSource"
        }

        Add-Type `
            -Path $openerSource `
            -ReferencedAssemblies $interopPath
        $openerType = `
            'TargetSpecHybrid.SwpDeployment.SolidWorksDocumentOpener' -as [type]
    }

    if ($null -eq $openerType) {
        throw 'The strongly typed SOLIDWORKS document opener could not be loaded.'
    }

    $openResult = $openerType::Open(
        $SolidWorks,
        $FileName,
        $DocumentType,
        $Options)

    if (-not $openResult.Succeeded) {
        # swFileLoadError_e.swFileWithSameTitleAlreadyOpen = 65536.
        # Confirmed against the installed SOLIDWORKS 2025 SP1.2 type
        # library. Do not close the existing document: it may be the
        # user's manual drawing or contain unsaved work.
        if (($openResult.Errors -band 65536) -ne 0) {
            $openDocuments = Get-SolidWorksOpenDocumentInfo `
                -SolidWorks $SolidWorks
            $requestedTitle = [IO.Path]::GetFileName($FileName)
            $conflicts = @(
                $openDocuments | Where-Object {
                    [IO.Path]::GetFileName($_.PathName) -ieq $requestedTitle
                } | ForEach-Object { $_.PathName })

            $conflictText = [string]::Join(';', $conflicts)
            throw "OpenDoc6 title conflict|requested=$FileName|" + `
                "openSameTitle=$conflictText|action=" + `
                'CloseConflictingDocumentsWithoutSavingThenRetry.'
        }

        throw "OpenDoc6 failed to open $FileName (errors=$($openResult.Errors), warnings=$($openResult.Warnings))."
    }

    return $openResult
}

function Get-SolidWorksOpenDocumentInfo {
    param([Parameter(Mandatory)] [object] $SolidWorks)

    [void] (Get-SolidWorksInteropType)
    $interopPath = Get-SolidWorksInteropPath

    $openerType = 'TargetSpecHybrid.SwpDeployment.SolidWorksDocumentOpener' -as [type]
    if ($null -eq $openerType) {
        $openerSource = Join-Path $PSScriptRoot 'SolidWorksDocumentOpener.cs'
        if (-not (Test-Path -LiteralPath $openerSource -PathType Leaf)) {
            throw "SOLIDWORKS document opener source was not found: $openerSource"
        }

        Add-Type `
            -Path $openerSource `
            -ReferencedAssemblies $interopPath
        $openerType = `
            'TargetSpecHybrid.SwpDeployment.SolidWorksDocumentOpener' -as [type]
    }

    if ($null -eq $openerType) {
        throw 'The SOLIDWORKS document opener could not be loaded.'
    }

    return @($openerType::ListOpenDocuments($SolidWorks))
}
