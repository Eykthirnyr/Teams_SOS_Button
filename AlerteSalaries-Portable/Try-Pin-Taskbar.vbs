Dim shell
Dim desktopPath
Dim shortcutPath
Dim folder
Dim item
Dim verb
Dim pinned

Set shell = CreateObject("WScript.Shell")
desktopPath = shell.SpecialFolders("Desktop")
shortcutPath = desktopPath & "\Alerte Salaries.lnk"

If CreateObject("Scripting.FileSystemObject").FileExists(shortcutPath) = False Then
    WScript.Echo "Le raccourci Bureau 'Alerte Salaries.lnk' est introuvable."
    WScript.Quit 1
End If

Set folder = CreateObject("Shell.Application").Namespace(desktopPath)
Set item = folder.ParseName("Alerte Salaries.lnk")

pinned = False
For Each verb In item.Verbs
    If InStr(LCase(Replace(verb.Name, "&", "")), "taskbar") > 0 Or _
       InStr(LCase(Replace(verb.Name, "&", "")), "barre des taches") > 0 Then
        verb.DoIt
        pinned = True
        Exit For
    End If
Next

If pinned Then
    WScript.Echo "Tentative d'epinglage envoyee a Windows."
    WScript.Quit 0
Else
    WScript.Echo "Windows n'a pas expose de verbe d'epinglage utilisable sur ce poste. Epinglage manuel recommande."
    WScript.Quit 2
End If
