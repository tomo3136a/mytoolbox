#builder
#
param($OutputPath = "./bin", [switch]$pass)

$AppName = Split-Path "." -Leaf
$AppExts = ".exe"

$Path = "src/*.cs"
$ReferencedAssemblies = "System.Drawing", "System.Windows.Forms", `
  "System.Xml", "System.Xml.Linq"

Write-Host "build start." -ForegroundColor Yellow

if (-not (Test-Path $OutputPath)) {
  New-Item -Path $OutputPath -ItemType Directory | Out-Null
  Write-Output "make directory."
}
$OutputAssembly = Join-Path (Resolve-Path -Path $OutputPath) $AppName$AppExts

Write-Output "build program:  $AppName$AppExts"
Write-Output "    Path:       $Path"
Write-Output "    Output:     $OutputAssembly"
Write-Output "    References: $ReferencedAssemblies"
Add-Type -Path $Path -OutputType WindowsApplication `
  -OutputAssembly $OutputAssembly `
  -ReferencedAssemblies $ReferencedAssemblies

Copy-Item -Force lib/uninstall*.cmd $OutputPath\..\lib
Copy-Item -Force lib/install*.cmd $OutputPath\..\lib
Copy-Item -Force lib/install*.ps1 $OutputPath\..\lib

Write-Host "build completed." -ForegroundColor Yellow
if (-not $pass) { $host.UI.RawUI.ReadKey() | Out-Null }
