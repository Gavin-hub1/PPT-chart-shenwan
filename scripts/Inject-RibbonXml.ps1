param(
    [Parameter(Mandatory = $true)]
    [string]$OfficeFilePath,

    [Parameter(Mandatory = $true)]
    [string]$RibbonXmlPath
)

$ErrorActionPreference = "Stop"

$officeFile = [System.IO.Path]::GetFullPath($OfficeFilePath)
$ribbonXml = [System.IO.Path]::GetFullPath($RibbonXmlPath)

if (-not (Test-Path $officeFile)) {
    throw "Office file not found: $officeFile"
}

if (-not (Test-Path $ribbonXml)) {
    throw "Ribbon XML not found: $ribbonXml"
}

Add-Type -AssemblyName System.IO.Compression.FileSystem

$tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("chartformatter_" + [guid]::NewGuid().ToString("N"))
$backupPath = "$officeFile.bak"

try {
    New-Item -ItemType Directory -Force -Path $tempDir | Out-Null
    [System.IO.Compression.ZipFile]::ExtractToDirectory($officeFile, $tempDir)

    $customUiDir = Join-Path $tempDir "customUI"
    New-Item -ItemType Directory -Force -Path $customUiDir | Out-Null
    $targetRibbonXml = Join-Path $customUiDir "customUI.xml"
    $ribbonContent = Get-Content -Raw -Path $ribbonXml
    $utf8Bom = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($targetRibbonXml, $ribbonContent, $utf8Bom)

    $relsDir = Join-Path $tempDir "_rels"
    $relsPath = Join-Path $relsDir ".rels"

    if (-not (Test-Path $relsPath)) {
        throw "Package relationship file not found: $relsPath"
    }

    [xml]$rels = Get-Content -Raw -Path $relsPath
    $ns = $rels.DocumentElement.NamespaceURI
    $existing = $rels.Relationships.Relationship | Where-Object {
        $_.Type -eq "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
    }

    if ($null -eq $existing) {
        $maxId = 0
        foreach ($rel in $rels.Relationships.Relationship) {
            if ($rel.Id -match "^rId(\d+)$") {
                $maxId = [Math]::Max($maxId, [int]$Matches[1])
            }
        }

        $newRel = $rels.CreateElement("Relationship", $ns)
        $newRel.SetAttribute("Id", "rId$($maxId + 1)")
        $newRel.SetAttribute("Type", "http://schemas.microsoft.com/office/2007/relationships/ui/extensibility")
        $newRel.SetAttribute("Target", "customUI/customUI.xml")
        $rels.Relationships.AppendChild($newRel) | Out-Null
        $rels.Save($relsPath)
    }
    else {
        $existing.Target = "customUI/customUI.xml"
        $rels.Save($relsPath)
    }

    Copy-Item -Force $officeFile $backupPath
    Remove-Item -Force $officeFile
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $officeFile)
    Remove-Item -Force $backupPath
}
catch {
    if ((Test-Path $backupPath) -and -not (Test-Path $officeFile)) {
        Copy-Item -Force $backupPath $officeFile
    }
    throw
}
finally {
    if (Test-Path $tempDir) {
        Remove-Item -Recurse -Force $tempDir
    }
}
