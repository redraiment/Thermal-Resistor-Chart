Set WshShell = Wscript.CreateObject("Wscript.Shell")
desktop = WshShell.SpecialFolders("Desktop")
programFiles = WshShell.Environment("Process")("ProgramFiles")
Set link = WshShell.CreateShortcut(desktop & "\电子温控.lnk")
link.TargetPath = programFiles & "\Dashboard\电子温控.exe"
link.WorkingDirectory = programFiles & "\Dashboard"
link.Save
