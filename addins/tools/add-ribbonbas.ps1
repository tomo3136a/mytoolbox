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

##############################################################################
#ribbon.bas
$path = Join-Path $root "ribbon.bas"
if (-not (Test-Path $path)) {
    Write-Host "# Creeate ribbon.bas" -ForegroundColor Yellow
    $txt = @"
Option Explicit
Option Private Module

'========================================
'rebbon interface
'========================================

Private g_ribbon As IRibbonUI

'----------------------------------------
'ribbon helper
'----------------------------------------

'リボンを更新
Private Sub RefreshRibbon(Optional id As String)
    If Not g_ribbon Is Nothing Then
        If id = "" Then
            g_ribbon.Invalidate
        Else
            g_ribbon.InvalidateControl id
        End If
    End If
    DoEvents
End Sub

'リボンID取得
Private Function RibbonID(control As IRibbonControl, Optional n As Long) As String
    Dim vs As Variant
    vs = Split(re_replace(control.id, "[^0-9.]", ""), ".")
    If UBound(vs) >= n Then RibbonID = Val("0" & vs(n))
End Function

'----------------------------------------
'イベント
'----------------------------------------

'起動時実行
Private Sub ${name}_onLoad(ByVal Ribbon As IRibbonUI)
    Set g_ribbon = Ribbon
End Sub

'----------------------------------------
'機能
'----------------------------------------
Private Sub ${name}_onAction(ByVal control As IRibbonControl)
    MsgBox control.id
End Sub
"@
  $txt | Set-Content $path
}
