# 6/1/23 JS
# This script is for use with Command Center. It pulls the desired computer name from a text file and uses that to enter the keystrokes to connect to that computer's remote registry
$Computer = Get-Content -Path C:\Users\Public\Documents\RemoteReg.txt

function Start-Reg {

    $wshell = New-Object -ComObject wscript.shell

    Start-Process C:\Windows\Regedit.exe

    Sleep 1

    $wshell.AppActivate('Registry Editor')

    Sleep 1

    $wshell.SendKeys("%FC")

    Sleep 1

    $wshell.SendKeys("$Computer")

    Sleep 1

    $wshell.SendKeys("{ENTER}")
}

Start-Reg