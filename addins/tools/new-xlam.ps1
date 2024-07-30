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

if (-not (Test-Path $xlam)) {
  Write-Host "# Create ${xlam}" -ForegroundColor Yellow
  $app = New-Object -ComObject Excel.Application
  $wb = $null
  try {
    $app.Visible = $false
    $app.DisplayAlerts = $false
    $wb = $app.Workbooks.Add()

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

Write-Host "open ${xlam_file}." -ForegroundColor Yellow
. $xlam

Start-Sleep 10
