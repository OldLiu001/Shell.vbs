Set FSO=CreateObject("scripting.filesystemobject")
Set ws=CreateObject("wscript.shell")

Function Fi_LS(ByVal oFolder)
	If oFolder=Null Then oFolder=ws.CurrentDirectory()
	If Not FSO.FolderExists(oFolder)Then 
		Set oFolder = fso.GetFolder(ws.CurrentDirectory)     '��ȡ�ļ���
	Else
		Set oFolder=FSO.GetFolder(oFolder)
	End If
	
	Set oSubFolders = oFolder.SubFolders    '��ȡ��Ŀ¼����
	
	Set oFiles = oFolder.Files              '��ȡ�ļ�����
	For Each file In oFiles
		Fi_LS=Fi_LS&(file.Name)&vbCrLf
	Next
	
	For Each folder In oSubFolders
		Fi_LS=Fi_LS&Fi_LS(folder)&vbCrLf
	Next

End Function

WScript.Echo Fi_LS("a")