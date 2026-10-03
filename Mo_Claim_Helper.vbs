Set WshShell = CreateObject("WScript.Shell")
scriptDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
WshShell.Run "pythonw.exe """ & scriptDir & "\claim_helper.pyw""", 0, False
