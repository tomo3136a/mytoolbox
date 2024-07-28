param([string]$path = "")

$addins_path = "${env:APPDATA}/Microsoft/Addins"
$name = ""
$root = ""
if ($path -ne "") {
  if ((Test-Path -PathType Container $path) -eq $True) {
    $name = (Split-Path -Leaf $path)
    $root = $path
  }
}
if ($name -eq "") { $name = Read-Host "アドイン名？" }
if ($name -eq "") {
  Write-Host "no addin name." -ForegroundColor Yellow
  $host.UI.RawUI.ReadKey() | Out-Null
  exit
}
if ($root -eq "") { $root = Join-Path (Get-Location).Path $name }
if (-not (Test-Path $root)) { [void](mkdir $root) }
Write-Host "name: ${name}" -ForegroundColor Yellow
Write-Host "root: ${root}" -ForegroundColor Yellow

$xlam_file = $name + ".xlam"
$xlam = Join-Path $addins_path $xlam_file

Push-Location $root

#
$zip_file = $name + ".zip"
$zip = Join-Path $root $zip_file
if (-not (Test-Path $zip)) {
  if (-not (Test-Path $xlam)) {
    Write-Host "# Create ${name}" -ForegroundColor Yellow
    $app = New-Object -ComObject Excel.Application
    $wb = $null
    try {
      $app.Visible = $false
      $app.DisplayAlerts = $false
      $wb = $app.Workbooks.Add()
      [void]$wb.SaveAs($xlam, 55)
    } finally {
      # release workbook
      if ($null -ne $wb) {
          [void]$wb.Close($false)
          [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)
      }
      [void]$app.Quit()
      # release application
      [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)
    }
  }
  Write-Host "# Load ${xlam}" -ForegroundColor Yellow
  Copy-Item $xlam $zip
}

$path = Join-Path $root "_rels\.rels"
if (-not (Test-Path $path)) {
  Write-Host "# Create _rels/.rels" -ForegroundColor Yellow
  7za.exe x -y $zip "_rels/.rels" | Out-Null
  $xml=[xml](Get-Content $path)
  $tags=$xml.GetElementsByTagName("Relationship")
  $id='customUI'
  $tags.Where({$_.Id -eq $id})|%{$xml.Relationships.RemoveChild($_)}|Out-Null
  $elm=$xml.Relationships.Relationship[0].Clone()
  $elm.Id=$id
  $elm.Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
  $elm.Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
  $elm.Target="customUI/customUI.xml"
  $xml.Relationships.AppendChild($elm) | Out-Null
  $xml.Save($path)
}
if (Test-Path $path) {
    Write-Host "# Add _rels/.rels" -ForegroundColor Yellow
    7za.exe u -ux2 -y $zip "_rels/.rels" | Out-Null
}

$path = Join-Path $root "customUI\customUI.xml"
if (-not (Test-Path $path)) {
    Write-Host "# Creeate customUI/customUI.xml" -ForegroundColor Yellow
    $xml = [xml]@"
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="Tab${name}" label="tab">
        <group id="${name}.G1" label="group" autoScale="true">
          <button id="${name}.B1" label="button" imageMso="ListMacros"
            onAction="${name}.M1" size="large" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
"@
    mkdir -Force ($root + "\customUI") | Out-Null
    $xml.Save($path)
}

if (Test-Path $path) {
  Write-Host "# Add customUI/customUI.xml" -ForegroundColor Yellow
  7za.exe u -ux2 -y $zip "customUI/customUI.xml" | Out-Null
}

if (Test-Path $zip) {
  Write-Host "# Save ${name}" -ForegroundColor Yellow
  Copy-Item $zip $xlam
}

Pop-Location

Write-Host "completed." -ForegroundColor Yellow
#$host.UI.RawUI.ReadKey() | Out-Null
