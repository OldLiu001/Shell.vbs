Set FSO=CreateObject("scripting.filesystemobject")
Set ws=CreateObject("wscript.shell")

Function Fi_LS(ByVal oFolder)
	If oFolder=Null Then oFolder=ws.CurrentDirectory()
	If Not FSO.FolderExists(oFolder)Then 
		Set oFolder = fso.GetFolder(ws.CurrentDirectory)     '获取文件夹
	Else
		Set oFolder=FSO.GetFolder(oFolder)
	End If
	
	Set oSubFolders = oFolder.SubFolders    '获取子目录集合
	
	Set oFiles = oFolder.Files              '获取文件集合
	For Each file In oFiles
		Fi_LS=Fi_LS&(file.Name)&vbCrLf
	Next
	
	For Each folder In oSubFolders
		Fi_LS=Fi_LS&Fi_LS(folder)&vbCrLf
	Next

End Function

WScript.Echo Fi_LS("a")