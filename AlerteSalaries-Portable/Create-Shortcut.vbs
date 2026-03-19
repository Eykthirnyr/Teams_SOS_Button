Dim shortcutPath
Dim targetPath
Dim arguments
Dim iconPath
Dim description
Dim shell
Dim shortcut
Dim fso

If WScript.Arguments.Count < 5 Then
    WScript.Quit 1
End If

shortcutPath = WScript.Arguments(0)
targetPath = WScript.Arguments(1)
arguments = WScript.Arguments(2)
iconPath = WScript.Arguments(3)
description = WScript.Arguments(4)

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set shortcut = shell.CreateShortcut(shortcutPath)

shortcut.TargetPath = targetPath
shortcut.Arguments = arguments
shortcut.WorkingDirectory = fso.GetParentFolderName(targetPath)
shortcut.IconLocation = iconPath
shortcut.Description = description
shortcut.Save
