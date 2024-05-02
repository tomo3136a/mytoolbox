param([string]$name)

if ($name -eq "") { $name = Read-Host "アドイン名？" }
if ($name -eq "") { Write-Host "no name"; exit }

$root=Join-Path (get-location).path $name
if (-not (Test-Path $root)) { [void](mkdir $root) }
Push-Location $root
$out="${env:APPDATA}/Microsoft/Addins"

$xlam=Join-Path $out ($name+".xlam")
if (-not (Test-Path $xlam)) {
    "# Create ${xlam}"|Out-Host
    $app=New-Object -ComObject Excel.Application
    $wb=$null
    try {
        $app.Visible = $false
        $app.DisplayAlerts = $false
        $wb=$app.Workbooks.Add()
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

$path=$root+"\_rels\.rels"
if (-not (Test-Path $path)) {
    "# Add _rels/.rels"|Out-Host
    7za.exe x -y $xlam "_rels/.rels"|Out-Null
    $xml=[xml](Get-Content $path)
    $tags=$xml.GetElementsByTagName("Relationship")
    $id='customUI'
    $tags.Where({$_.Id -eq $id})|%{$xml.Relationships.RemoveChild($_)}|Out-Null
    $elm=$xml.Relationships.Relationship[0].Clone()
    $elm.Id=$id
    $elm.Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
    $elm.Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
    $elm.Target="customUI/customUI.xml"
    $xml.Relationships.AppendChild($elm)|Out-Null
    $xml.Save($path)
    7za.exe u -y $xlam "_rels/.rels"|Out-Null
}

$path=$root+"\customUI\customUI.xml"
if (-not (Test-Path $path)) {
    "# Add customUI/customUI.xml"|Out-Host
    $xml=[xml]@"
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
    mkdir -Force ($root+"\customUI")|Out-Null
    $xml.Save($path)
}
7za.exe u -y $xlam "customUI/customUI.xml"|Out-Null

Pop-Location
