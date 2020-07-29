Class objColors
	Public Name, Color, Profile
End Class
Public strColors()

Function GetColor(strObjName)
	For i = 0 to UBound(strColors,1) - 1
	    If strObjName = strColors(i).Name Then
	        GetColor = strColors(i).Color
	        Exit Function
	    End If
	Next
	'意味着没有找到相关配置信息，报错。
	err.Rise 2
End Function

Function ReadColorProfile(Path)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FileExists(objFso.getfolder(Wscript.scriptfullName) & "\" & Path) Then
		err.Raise 3 '没有找到配置文件
		Exit Function
	End If
	Set objColorProfile = objFSO.Opentextfile(objFso.getfolder(Wscript.scriptfullName) & "\" & Path)
	If objColorProfile.ReadLine() <> "Visual_Basic_Script_Shell ColorProfile" Then
		err.Raise 4 '配置文件不合规
	End If
	Redim strColors(1)
	Do While Not ColorProfile.EOF
		ColorProfileLine = ColorProfile.ReadLine()
		If ColorProfileLine = "" Then
		Else
			If InStr(ColorProfileLine,"丨") <> 0 Then
				strColors(UBound(strColors,1)).Name = Mid(ColorProfileLine,1,InStr(ColorProfileLine,"丨") - 1)
				strColors(UBound(strColors,1)).Color = Mid(ColorProfileLine,InStr(ColorProfileLine,"丨") + 1 , Len(ColorProfileLine) - InStr(ColorProfileLine,"丨"))
				strColors(UBound(strColors,1)).Profile = Path
				Redim Preserve strColors(UBound(strColors,1) + 1)
			End If
		End If
	Loop
	Redim Preserve strColors(UBound(strColors,1) - 1)
	Set objColorProfile = Nothing
	Set objFSO = Nothing
End Function
