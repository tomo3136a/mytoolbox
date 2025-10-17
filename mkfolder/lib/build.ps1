#builder
#
param($OutputPath = "./bin", [switch]$pass)

$AppName = Split-Path "." -Leaf
$AppExts = ".exe"

$Path = "src/*.cs"
$ReferencedAssemblies = "System.Drawing", "System.Windows.Forms", `
  "System.Xml", "System.Xml.Linq", "System.Windows", "System.Configuration"

Write-Host "build start." -ForegroundColor Yellow

if (-not (Test-Path $OutputPath)) {
  New-Item -Path $OutputPath -ItemType Directory | Out-Null
  Write-Output "make directory."
}
$OutputAssembly = Join-Path (Resolve-Path -Path $OutputPath) $AppName$AppExts

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

Write-Output "build program:  $AppName$AppExts"
Write-Output "    Path:       $Path"
Write-Output "    Output:     $OutputAssembly"
Write-Output "    References: $ReferencedAssemblies"
Add-Type -Path $Path -OutputType WindowsApplication `
  -OutputAssembly $OutputAssembly `
  -ReferencedAssemblies $ReferencedAssemblies

Write-Host "build completed." -ForegroundColor Yellow
if (-not $pass) { $host.UI.RawUI.ReadKey() | Out-Null }
