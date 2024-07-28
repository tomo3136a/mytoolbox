
$name = "PERSONAL"
$root = "${env:APPDATA}/Microsoft/Excel/XLSTART"

if (-not (Test-Path $root)) { exit }

$xls_file = $name + ".XLSB"
$xls = Join-Path $root $xls_file
Write-Host "# ${xls}" -ForegroundColor Yellow

if (Test-Path $xls) {
  Write-Host "# File already exists, so it stopped." -ForegroundColor Yellow
  exit
}

Write-Host "# Create file..." -ForegroundColor Yellow
$app = New-Object -ComObject Excel.Application
$wb = $null
try {
  $app.Visible = $false
  $app.DisplayAlerts = $false
  $wb = $app.Workbooks.Add()
  $wb.Windows(1).Visible = $false
  [void]$wb.SaveAs($xls, 50)
} finally {
  # release workbook
  if ($null -ne $wb) {
    [void]$wb.Close($false)
    [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)
  }
  [void]$app.Quit()
  # release application
  [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($app)
  Write-Host "# Created." -ForegroundColor Yellow
}
