param([string]$path = "")

##############################################################################
#
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

$addins_path = "${env:APPDATA}/Microsoft/Addins"
$xlam_file = $name + ".xlam"
$xlam = Join-Path $addins_path $xlam_file
Write-Host "xlam: ${xlam}" -ForegroundColor Yellow

##############################################################################
#
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

#Write-Host "open ${xlam_file}." -ForegroundColor Yellow
#. $xlam

Start-Sleep 10
