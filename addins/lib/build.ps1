param([string]$path = "", [bool]$pass)

$addins="addindev","addindev2","myworks","mydesigner"

##############################################################################
#
function New-Xlam($xlam, $root) {
  if (-not (Test-Path $xlam)) {
    Write-Host "# Create ${xlam}" -ForegroundColor Yellow
    $app = New-Object -ComObject Excel.Application
    $wb = $null
    try {
      $app.Visible = $false
      $app.DisplayAlerts = $false
      $wb = $app.Workbooks.Add()

      #sheet data
      $dts = @()
      for ($i=0; $i -lt 255; $i++) { $dts += 2 }
      Get-ChildItem $root -Filter *.csv | %{
        $ws = $wb.Worksheets.Item(1)
        $ws.name = $_.Name -replace ".csv",""
        $name =$ws.name
        $csv = $_.FullName
        Write-Host "load ${name}.csv"
        $qt = $ws.QueryTables.Add("TEXT;$csv", $ws.cells(1, 1))
        $qt.TextFileCommaDelimiter = $True
        $qt.TextFileTabDelimiter = $True
        $qt.TextFilePlatform = 932
        $qt.TextFileStartRow = 1
        $qt.TextFileColumnDataTypes = $dts[0..255]
        $qt.name = "tmp_tbl"
        $qt.Refresh($false) | Out-Null
        $qt.Delete() | Out-Null
        $qt = $null
        foreach ($n in $wb.Names) {
          $sname = $n.Name
          Write-Host "search name ${sname}"
          if ($sname -Like ($name + "!" + "tmp_tbl*")) {
            Write-Host "delete name ${name}"
            $n.Delete()
          }
        }
      }

      # references
      try {
        $col = $wb.VBProject.References
        $s = $col.AddFromGuid("{420B2830-E718-11CF-893D-00A0C9054228}", 1, 0)
        $s = $s.Name
        Write-Host "reference ${s}"
      } catch {}
      try {
        $col = $wb.VBProject.References
        $s = $col.AddFromGuid("{3F4DACA7-160D-11D2-A8E9-00104B365C9F}", 5, 5)
        $s = $s.Name
        Write-Host "reference ${s}"
      } catch {}

      # sources
      try {
        $col = $wb.VBProject.VBComponents
        Get-ChildItem $root -Filter *.bas | %{
          $s = $_.Name
          $col.Import($_.FullName) | Out-Null
          Write-Host "load ${s}"
        }
        Get-ChildItem $root -Filter *.frm | %{
          $s = $_.Name
          $col.Import($_.FullName) | Out-Null
          Write-Host "load ${s}"
        }
        Get-ChildItem $root -Filter *.cls | %{
          $s = $_.Name -replace ".cls",""
          $col.Import($_.FullName) | Out-Null
          Write-Host "load ${s}"
        }
      } catch {}

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
}

##############################################################################
#wait for file close
function Wait-FileClosed($xlam)
{
  Start-Sleep 1
  if (Test-Path $xlam) {
    Try {
        $st = [System.IO.File]::Open($xlam,'Open','Write')
        $st.Close()
        $st.Dispose()
    } Catch {
        Write-Host "# wait... 10sec" -ForegroundColor Yellow
        Start-Sleep 10
    }
  }
}

##############################################################################
#
Set-Item env:Path "${env:ProgramFiles}\7-Zip\;${env:Path}"
$arc = "7z.exe"

function Update-ArchiveXml($xml, $dst, $target) {
  $stm = New-Object System.IO.MemoryStream
  $wrt = New-Object System.Xml.XmlTextWriter($stm, [System.Text.Encoding]::Unicode)
  $xml.WriteContentTo($wrt)
  $wrt.Flush()
  $stm.Flush()
  $stm.Position = 0
  $rdr = New-Object System.Io.StreamReader($stm)
  $txt = $rdr.ReadToEnd()
  $txt | Out-Host

  Write-Host "# Add ${target}" -ForegroundColor Yellow
  $txt | . $arc u -ux2 -y -tzip $dst -si"${target}" | Out-Null
}

##############################################################################
# add relationships
function Update-Relationships($dst) {
  $target="_rels\.rels"
  Write-Host "target: ${target}" -ForegroundColor Yellow
  $da=. $arc x -so -tzip $dst $target
  if ($da.Count -ne 0) {
    $xml=[xml]$da
    $tags=$xml.GetElementsByTagName("Relationship")
    $id='customUI'
    $tags.Where({$_.Id -eq $id})|%{$xml.Relationships.RemoveChild($_)}|Out-Null
    $elm=$xml.Relationships.Relationship[0].Clone()
    $elm.Id=$id
    $elm.Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
    $elm.Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"
    $elm.Target="customUI/${id}.xml"
    $xml.Relationships.AppendChild($elm) | Out-Null
    Update-ArchiveXml $xml $dst $target
  }
}

##############################################################################
#
function Add-CustomUI($dst, $name, $noribbon, $ctmenu) {
  $id="customUI"

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

#  $path = Join-Path $root "customUI\${id}.xml"
#  if (-not (Test-Path $path)) {
  Write-Host "# Creeate customUI/${id}.xml" -ForegroundColor Yellow
  $src = @"
<?xml version="1.0" encoding="UTF-8"?>
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="${name}_onLoad">
"@
  if (-not $noribbon) { $src = $src + $s1 }
  if ($ctmenu) { $src = $src + $s2 }
  $src = $src + @"
</customUI>
"@
  $target="customUI/${id}.xml"
  $xml = [xml]$src
  Write-Host "# Update ${target}" -ForegroundColor Yellow
  Update-ArchiveXml $xml $dst $target
}

##############################################################################
#
$out_path = Join-Path (Resolve-Path ".").Path "addins"
if (-Not (Test-Path $out_path)) {
  New-Item $out_path -ItemType Directory | Out-Null
}

foreach ($name in $addins) {
  $xlam_file = $name + ".xlam"
  $xlam = Join-Path $out_path $xlam_file
  Write-Host "xlam: ${xlam_file}" -ForegroundColor Yellow

  New-Xlam $xlam $name
  Wait-FileClosed $xlam
  Update-Relationships $xlam

  Push-Location $name
  . $arc a -tzip $xlam "customUI\customUI.xml" | Out-Null
  Pop-Location
  Write-Host "# Update customUI\customUI.xml" -ForegroundColor Yellow
}

#if (Test-Path $zip) {
#  Write-Host "# Save ${name}" -ForegroundColor Yellow
#  Copy-Item $zip $xlam
#  Remove-Item $zip
#}

##############################################################################
#
Write-Host "build completed." -ForegroundColor Yellow
if (-not $pass) { $host.UI.RawUI.ReadKey() | Out-Null }
