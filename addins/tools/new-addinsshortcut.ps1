$name="AddIn"
$wsh=New-Object -ComObject Wscript.Shell
$lnk=$wsh.CreateShortcut($name+".lnk")
$lnk.TargetPath="${env:APPDATA}\Microsoft\AddIns"
$lnk.WorkingDirectory="."
$lnk.Save()
