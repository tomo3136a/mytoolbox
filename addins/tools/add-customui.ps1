param([string]$path = "")

Set-Item env:Path "${env:ProgramFiles}\7-Zip\;${env:Path}"
$arc = "7z.exe"

##############################################################################
#
$name = ""
$root = ""
if ($path -ne "") {
  $path = Resolve-Path $path
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

$addins_path = "${env:APPDATA}/Microsoft/Addins"
$xlam_file = $name + ".xlam"
$xlam = Join-Path $addins_path $xlam_file
Write-Host "xlam: ${xlam}" -ForegroundColor Yellow
Push-Location $root

##############################################################################
#
$zip_file = $name + ".zip"
$zip = Join-Path $root $zip_file
Write-Host "zip: ${zip}" -ForegroundColor Yellow
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

##############################################################################
#
$path = Join-Path $root "_rels\.rels"
Write-Host "path: ${path}" -ForegroundColor Yellow
if (-not (Test-Path $path)) {
  Write-Host "# Create _rels/.rels" -ForegroundColor Yellow
  . $arc x -y $zip "_rels/.rels" | Out-Null
  $xml=[xml](Get-Content $path)
  $tags=$xml.GetElementsByTagName("Relationship")
  $id='customUI'
  $tags.Where({$_.Id -eq $id})|%{$xml.Relationships.RemoveChild($_)}|Out-Null
  $elm=$xml.Relationships.Relationship[0].Clone()
  $elm.Id=$id
  $elm.Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
  $elm.Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
  $elm.Target="customUI/${id}.xml"
  $xml.Relationships.AppendChild($elm) | Out-Null
  $xml.Save($path)
}
if (Test-Path $path) {
    Write-Host "# Add _rels/.rels" -ForegroundColor Yellow
    . $arc u -ux2 -y $zip "_rels/.rels" | Out-Null
}

##############################################################################
#
$id="customUI"
$path = Join-Path $root "customUI\${id}.xml"
if (-not (Test-Path $path)) {
    Write-Host "# Creeate customUI/${id}.xml" -ForegroundColor Yellow

    $s1 = @"
  <ribbon>
    <tabs>
      <tab id="Tab${name}" label="${name}">
        <group id="${name}.g1" label="${name}" autoScale="true">
          <button id="${name}.b1" label="${name}" imageMso="ListMacros"
            onAction="${name}_onAction" size="large" />
        </group>
      </tab>
    </tabs>
  </ribbon>
"@
    $s2 = @"
  <contextMenus>
    <contextMenu idMso="ContextMenuWorkbookPly">
      <button id="${name}.sheet.b1" label="${name}" imageMso="ListMacros" onAction="${name}_onAction" />
    </contextMenu>
    <contextMenu idMso="ContextMenuCell">
      <button id="${name}.cell.b1" label="${name}" insertBeforeMso="Cut" imageMso="ListMacros" onAction="${name}_onAction" />
      <menuSeparator id="${name}.cell.s1" insertBeforeMso="Cut" />
    </contextMenu>
    <contextMenu idMso="ContextMenuCellLayout">
      <button id="${name}.celllayout.b1" label="${name} layout" insertBeforeMso="Cut" imageMso="ListMacros" onAction="${name}_onAction" />
      <menuSeparator id="${name}.celllayout.s1" insertBeforeMso="Cut" />
    </contextMenu>
    <contextMenu idMso="ContextMenuRow">
      <button id="${name}.row.b1" label="${name}" insertBeforeMso="Cut" imageMso="ListMacros" onAction="${name}_onAction" />
      <menuSeparator id="${name}.row.s1" insertBeforeMso="Cut" />
    </contextMenu>
    <contextMenu idMso="ContextMenuRowLayout">
      <button id="${name}.rowlayout.b1" label="${name} layout" insertBeforeMso="Cut" imageMso="ListMacros" onAction="${name}_onAction" />
      <menuSeparator id="${name}.rowlayout.s1" insertBeforeMso="Cut" />
    </contextMenu>
    <contextMenu idMso="ContextMenuColumn">
      <button id="${name}.col.b1" label="${name}" insertBeforeMso="Cut" imageMso="ListMacros" onAction="${name}_onAction" />
      <menuSeparator id="${name}.col.s1" insertBeforeMso="Cut" />
    </contextMenu>
    <contextMenu idMso="ContextMenuColumnLayout">
      <button id="${name}.collayout.b1" label="${name} layout" insertBeforeMso="Cut" imageMso="ListMacros" onAction="${name}_onAction" />
      <menuSeparator id="${name}.collayout.s1" insertBeforeMso="Cut" />
    </contextMenu>
    <contextMenu idMso="ContextMenuShape">
      <button id="${name}.shape.b1" label="${name}" insertBeforeMso="Cut" imageMso="ListMacros" onAction="${name}_onAction" />
      <menuSeparator id="${name}.shape.s1" insertBeforeMso="Cut" />
    </contextMenu>
    <contextMenu idMso="ContextMenuPicture">
      <button id="${name}.pict.b1" label="${name}" insertBeforeMso="Cut" imageMso="ListMacros" onAction="${name}_onAction" />
      <menuSeparator id="${name}.pict.s1" insertBeforeMso="Cut" />
    </contextMenu>
  </contextMenus>
"@

    $xml = [xml]@"
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="${name}_onLoad">
  ${s1}
  ${s2}
</customUI>
"@
    mkdir -Force ($root + "\customUI") | Out-Null
    $xml.Save($path)
}
if (Test-Path $path) {
  Write-Host "# Add customUI/${id}.xml" -ForegroundColor Yellow
  . $arc u -ux2 -y $zip "customUI/${id}.xml" | Out-Null
}

##############################################################################
#
if (Test-Path $zip) {
  Write-Host "# Save ${name}" -ForegroundColor Yellow
  Copy-Item $zip $xlam
  Remove-Item $zip
}

Pop-Location

Write-Host "completed." -ForegroundColor Yellow
