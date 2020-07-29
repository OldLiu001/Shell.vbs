Set fso = CreateObject("Scripting.FileSystemObject")
Set objShell=CreateObject("Shell.Application")

Sub UnZip(ByVal myZipFile,ByVal myTargetDir)
    If NOT fso.FolderExists(myTargetDir) Then
        fso.CreateFolder(myTargetDir)
    End If
    Set objSource = objShell.NameSpace(myZipFile)
    Set objFolderItem = objSource.Items()
    Set objTarget = objShell.NameSpace(myTargetDir)
    objTarget.CopyHere objFolderItem,256
End Sub
