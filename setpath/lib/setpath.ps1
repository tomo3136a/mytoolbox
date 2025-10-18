param($p="c:\opt\bin", [switch]$pass)

if (test-path $p) {
    Write-Host "path: ${p}" -ForegroundColor Yellow
    $path = [Environment]::GetEnvironmentVariable('PATH', 'User')
    $paths=($path).Split(";")
    if(-not ($paths|?{$_ -like $p})){
        $path=($paths + @(,$p)) -join ";"
        [Environment]::SetEnvironmentVariable('PATH', $path, 'User')
    }
}

Write-Host "setpath completed." -ForegroundColor Yellow
if (-not $pass) { $host.UI.RawUI.ReadKey() | Out-Null }
