Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\Users\avera\Desktop\Ucar_Hukuk"
WshShell.Run chr(34) & "C:\Users\avera\Desktop\Ucar_Hukuk\Ucar_Hukuk.bat" & chr(34), 0
Set WshShell = Nothing