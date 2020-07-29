Class objColors
	Public Name, Color, Profile
End Class

Public strColors()
ReadColorProfile("Preference\DefaultColor.crp")
Function GetColor(strObjName)
	For i = 1 to UBound(strColors) 
	    If strObjName = strColors(i).Name Then
	        GetColor = strColors(i).Color
	        Exit Function
	    End If
	Next
	'��ζ��û���ҵ����������Ϣ������
	err.Raise 2
End Function

Function ReadColorProfile(Path)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FileExists(Left(Wscript.scriptfullName,InStrRev(Wscript.scriptfullName,"\")) & Path) Then
		err.Raise 3,"û���ҵ������ļ�" 'û���ҵ������ļ�
		Exit Function
	End If
	Set objColorProfile = objFSO.Opentextfile(Left(Wscript.scriptfullName,InStrRev(Wscript.scriptfullName,"\")) & Path)
	If objColorProfile.ReadLine() <> "Visual_Basic_Script_Shell ColorProfile" Then
		err.Raise 4 '�����ļ����Ϲ�
	End If
	Redim strColors(1)
	Do While Not objColorProfile.atEndOfStream
		ColorProfileLine = objColorProfile.ReadLine()
		If ColorProfileLine = "" Then
		Else
			If InStr(ColorProfileLine,"ح") <> 0 Then
				Set strColors((UBound(strColors))) = New objColors
				strColors(UBound(strColors)).Name = Mid(ColorProfileLine,1,InStr(ColorProfileLine,"ح") - 1)
				strColors(UBound(strColors)).Color = Mid(ColorProfileLine,InStr(ColorProfileLine,"ح") + 1 , Len(ColorProfileLine) - InStr(ColorProfileLine,"ح"))
				strColors(UBound(strColors)).Profile = Path
				Redim Preserve strColors(UBound(strColors) + 1)
			End If
		End If
	Loop
	Redim Preserve strColors(UBound(strColors) - 1)
	Set objColorProfile = Nothing
	Set objFSO = Nothing
End Function
