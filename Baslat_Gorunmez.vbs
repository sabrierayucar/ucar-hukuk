Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = "C:\Users\avera\Desktop\Ucar_Hukuk"
WshShell.Run """C:\Users\avera\AppData\Local\Programs\Python\Python314\pythonw.exe"" ""C:\Users\avera\Desktop\Ucar_Hukuk\main.py""", 0, False
