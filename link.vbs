Set WshShell = Wscript.CreateObject("Wscript.Shell")
desktop = WshShell.SpecialFolders("Desktop")
programFiles = WshShell.Environment("Process")("ProgramFiles")
Set link = WshShell.CreateShortcut(desktop & "\�����¿�.lnk")
link.TargetPath = programFiles & "\Dashboard\�����¿�.exe"
link.WorkingDirectory = programFiles & "\Dashboard"
link.Save
