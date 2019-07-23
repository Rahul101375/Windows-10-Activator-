Set oShell = CreateObject ("Wscript.Shell") 
Dim strArgs
strArgs = "cmd /c file.bat"
oShell.Run strArgs, 0, false