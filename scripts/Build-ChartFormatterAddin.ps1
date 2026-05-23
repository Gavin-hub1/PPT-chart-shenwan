param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot "..\dist"),
    [switch]$KeepPowerPointOpen
)

$ErrorActionPreference = "Stop"

$root = Resolve-Path (Join-Path $PSScriptRoot "..")
$modulePath = Join-Path $root "src\ChartFormatter.bas"
$ribbonPath = Join-Path $root "ribbon\customUI.xml"
$outDir = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputDirectory)
$pptmPath = Join-Path $outDir "ChartFormatter.pptm"
$ppamPath = Join-Path $outDir "ChartFormatter.ppam"

New-Item -ItemType Directory -Force -Path $outDir | Out-Null

if (-not (Test-Path $modulePath)) {
    throw "VBA module not found: $modulePath"
}

if (-not (Test-Path $ribbonPath)) {
    throw "Ribbon XML not found: $ribbonPath"
}

Write-Host "Creating PowerPoint macro workbook..."
$powerPoint = $null
$presentation = $null

try {
    try {
        $powerPoint = [Runtime.InteropServices.Marshal]::GetActiveObject("PowerPoint.Application")
        Write-Host "Using existing PowerPoint instance..."
    }
    catch {
        try {
            $powerPoint = New-Object -ComObject PowerPoint.Application
        }
        catch {
            throw "Could not start PowerPoint automation. Save your work, close all PowerPoint windows, end any POWERPNT.EXE processes in Task Manager, then run the build again. Original error: $($_.Exception.Message)"
        }
    }

    $powerPoint.Visible = -1
    $presentation = $powerPoint.Presentations.Add(-1)

    try {
        $vbProject = $presentation.VBProject
        if ($null -eq $vbProject) {
            throw "PowerPoint returned a null VBProject. Enable 'Trust access to the VBA project object model' and make sure Office VBA support is installed."
        }

        $vbComponents = $vbProject.VBComponents
        if ($null -eq $vbComponents) {
            throw "PowerPoint returned a null VBComponents collection. Enable 'Trust access to the VBA project object model' and restart PowerPoint."
        }

        $vbComponents.Import($modulePath) | Out-Null
    }
    catch {
        throw "PowerPoint blocked VBA import. Enable: File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model. Original error: $($_.Exception.Message)"
    }

    # ppSaveAsOpenXMLPresentationMacroEnabled = 25; ppSaveAsOpenXMLAddin = 30.
    $presentation.SaveAs($pptmPath, 25)
    $presentation.SaveAs($ppamPath, 30)
}
finally {
    if ($presentation -ne $null -and -not $KeepPowerPointOpen) {
        $presentation.Close()
    }
    if ($powerPoint -ne $null -and -not $KeepPowerPointOpen) {
        $powerPoint.Quit()
    }
}

Write-Host "Injecting Ribbon XML..."
& (Join-Path $PSScriptRoot "Inject-RibbonXml.ps1") -OfficeFilePath $pptmPath -RibbonXmlPath $ribbonPath
& (Join-Path $PSScriptRoot "Inject-RibbonXml.ps1") -OfficeFilePath $ppamPath -RibbonXmlPath $ribbonPath

Write-Host ""
Write-Host "Build complete:"
Write-Host "  $pptmPath"
Write-Host "  $ppamPath"
