Dim Code,NotNeedCheck
Dim FSO,WS,Args,WSH,ObjCmd,TLI
Set WSH=WScript

Set FSO= CreateObject("Scripting.FileSystemObject")
Set WS= CreateObject("Wscript.Shell")
'�����Ҫ�����
Dim LibFiles
Dim CMD_Input,IsOutPut,Help_Str,str,IsConfig,InputType,PromptStr,PromptSet,ControlNum,filepath,char,RootFolder,ScriptType,ScriptName,FuncList,file,ScriptFile,IsScript,LineCount,ScriptLine,NeedExecute
Dim Isrecv,RecvData,IsRemote'Socket����
IsOutPut=True
IsConfig=True
InputType=1
CMD_Input="'"
ControlNum=0
CurrentDirectory=ws.CurrentDirectory'��ͨ������ʼ��

Execute("Public Const Desktop= """&WS.ExpandEnvironmentStrings("%UserProFile%\Desktop")&Chr(34))
Execute("Public Const WinDir="""&WS.ExpandEnvironmentStrings("%Windir%")&Chr(34))
Execute("Public Const Path="""&WS.ExpandEnvironmentStrings("%path%")&Chr(34))
Execute("Public Const UserProFile="""&WS.ExpandEnvironmentStrings("%Userprofile%")&Chr(34))
Execute("Public Const FLib="""&WS.ExpandEnvironmentStrings("%VBSH_FLib%")&Chr(34))
Execute("Public Const ComLib="""&WS.ExpandEnvironmentStrings("%VBSH_Dll%")&Chr(34))
Dim rIsUAC:rIsUAC=IsUAC


'��ʼ����̬����
Public Const NL=Null
Public Const BK=1
Public Const RT=2'��Ҫ�ı�����ʼ��
Public Const KB=1024
Public Const MB=1048576
Public Const GB=1073741824
Public Const TB=1099511627776

Public Const VersionNum="0016"
Public Const VersionName="1.0.0.16"
Public Const ProductName="VBS Shell"
Public Const LastVersiondTime="202005142100" '�汾�Ͳ�Ʒ��Ϣ
'�ð汾ʹ��CScript���������С�
filearr=Split(ReadAll(WScript.ScriptFullName),vbCrLf)'��ʼ����̬����
For i=0 To UBound(filearr)
		If LCase(Trim(Left(filearr(i),9)))="function"Or LCase(Trim(Left(filearr(i),3)))="sub" Then
			Help_Str=Help_Str&filearr(i)&vbCrLf
		End If 
Next

'------------------------Preparing.....-------------------------------------

Function Start()
	WS.Run "%windir%\system32\Cscript //nologo "&WScript.ScriptFullName 
End Function

Sub GetCommand()'�г���������
	WScript.Echo(Help_Str)'���������������˵��
End Sub

Sub GetHelp(ByVal FuncStr)
	If InStr(FuncStr,"function")Or InStr(FuncStr,"sub")Then
		WSH.Echo("��Ӧ��'sub'����'function'�ڲ�ѯ�ַ��С�")
		Exit Sub
	End If

	If IsActiveX(Funcstr) Then 
		Set TLI=CreateObject("TLI.TLIApplication")
		Set O=CreateObject(Funcstr)
		Set Info = tli.ClassInfoFromObject(O)
		For Each Member In Info.DefaultInterface.Members
    		Echo Member.Name 
		Next
		Exit Sub
	End If
	
	If FSO.FileExists(Funcstr)Then
		If LCase(FSO.GetExtensionName(Funcstr))="dll"Or LCase(FSO.GetExtensionName(Funcstr))="ocx" Then
			Set TLI=CreateObject("TLI.TLIApplication")
			Set Info = tli.TypeLibInfoFromFile("scrrun.dll")
			For Each Interface In Info.Interfaces
   				Echo Interface.Name 
			Next
			Exit sub
		End If
	End If
	
	StrArr=Split(Help_Str,vbCrLf)
	For i=0 To UBound(StrArr)
		If InStr(LCase(StrArr(i)),LCase(FuncStr)) Then
			Echo(StrArr(i))
		End If
	Next
End Sub

Function IsActiveX(ByVal ComName)
	On Error Resume Next
    Dim O
    Set O = CreateObject(ComName)
    If Err.Number = 0 Then
        IsActiveX = True
    Else
        IsActiveX = False
    End If
    Set O = Nothing
    Err.Clear
    On Error Goto 0
End Function


Sub ErrorDealing()
	If Err.Number <> 0 Then 
		WScript.Echo "����" & Err.Number
		WScript.Echo Err.Description
		Select Case Err.Number
			Case 62  
			WScript.Echo "���������������Ctrl+Z����ġ�"
		Case Else
			WScript.Echo("��֧��vbscript�������ԣ�")
			WScript.Echo("  1.����������������" & Err.Description & "��")
			WScript.Echo("  2.���롱About������ϵ�����ߡ�")
		End Select
		ExecuteGlobal("input=Null")
	End If
End Sub

Sub CLS()
	For i=0 To 2000
		str=str&vbCrLf
	Next
	WScript.Echo str
	str=Null
End Sub

Function WhichType(ByVal varname)
	If varname=Null Then 
		WhichType="[null]"
	ElseIf IsArray(varname) Then 
		WhichType="[Array]"
	ElseIf IsObject(varname) Then
		WhichType="[Object]"
	ElseIf IsNumeric(varname) Then 
		If InStr(varname,".")Then 
			WhichType="[Float]"
		Else
			WhichType="[Int]"
		End If		
	End If
End Function

Sub Import(ByVal sInstFile) 
	If FSO.FileExists(FLib&"\"&sInstFile) Then sInstFile=Lib&"\"&sInstFile
	If FSO.FileExists(sInstFile&".vbs") Then sInstFile=sInstFile&".vbs"
	If Not FSO.FileExists(sInstFile) Then 
		WScript.Echo("Import:Can't find the moudle file("&sInstFile&").")
		Exit Sub
	End If
	LibFiles=LibFiles&sInstFile&vbCrLf
	FileStr=ReadAll(sInstFile)
	filearr=Split(FileStr,vbCrLf)
	Help_Str=Help_Str&sInstFile&"{"&vbCrLf
	For i=0 To UBound(filearr)
		If LCase(Trim(Left(filearr(i),9)))="function"Or LCase(Trim(Left(filearr(i),3)))="sub" Then
			
			Help_Str=Help_Str&"    "&filearr(i)&vbCrLf
		End If 
	Next
	Help_Str=Help_Str&"}"&vbCrLf&vbCrLf 
	On Error Resume Next
		ExecuteGlobal(FileStr)
		If Err.Number<>0 Then WS.Run FileStr:Exit sub
	On Error Goto 0	
End Sub 

Sub About()
	WScript.Echo("�����ڱ�����")
	WScript.Echo("")
	WScript.Echo(ProductName)
	WScript.Echo("�汾�ţ�" & VersionName)
	WScript.Echo("�ڲ��汾��" & VersionNum)
	WScript.Echo("���༭ʱ�䣺" & LastVersiondTime)
	WScript.Echo("")
	WScript.Echo("    �����������Ȩ��TECH_Noob��С�������С�����Ȩ��������")
	WScript.Echo("    �����򲻻��¼�����κθ�����Ϣ����������ʹ�á�")
	WScript.Echo("    ������Ҳ�������κζ�����룬�����ų�����������ڿ�����Ŀ��ܡ��������������ȡ�ó���")
	WScript.Echo("    �������Ҫ��ϵ������Ŀ����ߣ���ͨ������������")
	WScript.Echo("��վ��")
	WScript.Echo("	������2��yangruixian.icoc.me")	
	WScript.Echo("���䣺")
	WScript.Echo("	TECH_N00b��3464943410@qq.com")
	WScript.Echo("	������2��joengjeoijin@outlook.com")
	WScript.Echo("QQ��")
	WScript.Echo("TECH_N00b��3464943410")
	WScript.Echo("����������κ����ʣ���ӭ��������ϵ��")
End Sub

Function ReadAll(ByVal filepath)
	If Not fso.FileExists(filepath) Then Exit Function
	Set file=fso.OpenTextFile(filepath)
	ReadAll=file.ReadAll
	Set file=Nothing
End Function

Function Cat(ByVal filepath)
	If Not fso.FileExists(filepath) Then Exit Function
	Set file=fso.OpenTextFile(filepath)
	ReadAll=file.ReadAll
	Set file=Nothing
End Function


Function Trimpp(ByVal input)
	str=input
	strarr=Split(input,Chr(34))
	For i=0 To UBound(strarr)
		If Not IsInt((i+1)/2) Then
			str=Replace(str,strarr(i),"")
		Else
			str=Replace(str,strarr(i),LCase(strarr(i)))
		End If
	Next
	
	strarr=Split(str,vbCrLf)
	For i=0 To UBound(strarr)
		For intc=0 To Len(strarr(i))
			If Right(Left(strarr(i),intc),1)="'" Then
				outstr=Left(strarr(i),intc)
				MsgBox(outstr)
				Exit For
			End If
		Next
		str=str&outstr&vbCrLf
	Next
	Trimpp=str
	str=Null
End Function

Function InstallVBM(MoudlePath)
	If Not FSO.FolderExists(Lib) Then WScript.Echo("Can't find the 'lib', reinstall the program can solve the problem."):Exit Function
	If Not FSO.FileExists(MoudlePath) Then WScript.Echo("The Moudle file isn't exist."):Exit Function
	If LCase(FSO.GetExtensionName(MoudlePath))<>"vbs"Or LCase(FSO.GetExtensionName(MoudlePath))<>"vbe"Or LCase(FSO.GetExtensionName(MoudlePath))<>"vbm" Then WScript.Echo("The ExtensionName must be 'vbs,vbe or vbm'"):Exit Function
	FSO.MoveFile MoudlePath,Lib
	WScript.Echo("Finish!")
End Function

Function IsInt(ByVal num)
	If InStr(num,".") And IsNumeric(num)Then 
		IsInt=True
	Else
		IsInt=False
	End If
End Function

Function CD(ByVal PathStr)
	PathStr=Trim(PathStr)
	If PathStr=Null Then
		CD=WS.CurrentDirectory
		Exit Function
	End If
	If IsNumeric(PathStr) Then
		Select Case PathStr
			Case 1
			str=ws.CurrentDirectory
			str=Left(str,InStrRev(str,"\"))
			ws.CurrentDirectory=str
			str=Null
			Case 2
			str=Left(ws.CurrentDirectory,1)
			ws.CurrentDirectory=str&":\"
			str=Null
			Case Else
			WSH.Echo("Path isn't correct.....")
		End Select
	Else
		If FSO.FolderExists(ws.CurrentDirectory&"\"&PathStr) Then
			WS.CurrentDirectory=ws.CurrentDirectory&"\"&PathStr
			Exit Function
		End If
		If FSO.FolderExists(ws.CurrentDirectory&PathStr) Then
			WS.CurrentDirectory=ws.CurrentDirectory&PathStr
			CD=WS.CurrentDirectory
			Exit Function
		End If
		If Not fso.FolderExists(PathStr) Then
			WScript.Echo("Path isn't correct.....")
			CD=WS.CurrentDirectory
			Exit Function
		Else
			ws.CurrentDirectory=PathStr
		End If
	End If
	CD=WS.CurrentDirectory		
End Function

Function System(ByVal Cmd)
	Set ObjCmd=WS.Exec("cmd /c "&Cmd)
	Do While ObjCmd.StdOut.AtEndOfStream=False
		WSH.StdOut.Writeline(ObjCmd.StdOut.Readline)
	Loop
	Set ObjCmd=Nothing
End Function

Function IsCommand(ByVal str)
	Dim Strings
	str=str&" "
	char=Null
	Do Until char=" "
		i=i+1
		char=Right(Left(str,i),1)
		Strings=Left(str,i)
	Loop
	i=0
	Strings=LCase(Trim(Strings))
	
	If fso.FileExists(Strings) Then
		IsCommand=True
		Exit Function
	End If
	
	Paths=Split(ws.ExpandEnvironmentStrings("%path%"),";")
	For i=0 To UBound(Paths)
		If fso.FileExists(Paths(i)&"\"&Strings) Then
			IsCommand=True
			Exit Function
		End If
		
		If fso.FileExists(Paths(i)&"\"&Strings&".exe") Then
			IsCommand=True
			Exit Function
		End If
	Next
	
	
	If InStr(str,"if") Then
		If InStr(str,"/") Then IsCommand=True 
		If InStr(str,"%") Then IsCommand=True
		If InStr(str,"then") Then IsCommand=False
		If IsCommand Then Exit Function
	End If
	
	If InStr(str,"set") Then
		If InStr(str,".") Then 
			IsCommand=False
			Exit Function
		End If
		If InStr(str,"%") Then IsCommand=True
		If Not InStr(str,"/") Then IsCommand=True
		If IsCommand Then Exit Function
	End If
	
	
	If InStr(str,"for") Then
		If InStr(str,"/") Then IsCommand=True 
		If InStr(str,"%") Then IsCommand=True
		If InStr(str,"do") Then IsCommand=True
		If IsCommand Then Exit Function
	End If
	
	If Strings="powershell" Or Strings="cmd" Or Strings="choice" Or Strings="kill"Or Strings="shutdown"Or Strings="wget" Then
		IsCommand=False
		Exit Function
	End If
	
	If Strings = "md" Or Strings = "setx"Or Strings = "ren" Or Strings = "xcopy" Or Strings = "copy" Or Strings = "rd" Or Strings = "format" Or Strings = "del" Or Strings = "pushd" Or Strings = "popd" Or Strings = "type" Or Strings = "color" Or Strings = "call" Or Strings = "dir" Or Strings = "erase" Or Strings = "shift" Or Strings = "ftype" Or Strings = "assoc" Then
		IsCommand=True
		Exit Function
	End If
	
	For i=0 To UBound(Paths)
		If fso.FileExists(Paths(i)&"\"&Strings&".bat") Then
			IsCommand=True
			Exit Function
		End If
		If fso.FileExists(Paths(i)&"\"&Strings&".vbe") Then
			IsCommand=True
			Exit Function
		End If
		
		If fso.FileExists(Paths(i)&"\"&Strings&".vbs") Then
			IsCommand=True
			Exit Function
		End If
	Next
End Function

Function Deal(ByVal input)
	If input=Null Then Exit Function 
	If LCase(Trim(Left(input,1)))="'" Then Exit Function
	If IsCommand(input) Then
		ScriptType="bat"
	Else
		Replace input,"wscript.scriptfullname",ScriptName
	End If
	
	On Error Resume Next
		Err.Clear
		If InStr(input,"=") Then 
			
		Else
			TestEval=Eval(input)
		End If
	On Error Goto 0
		If Len(TestEval)<>0 Then
			WScript.Echo(TestEval)
			input=" "
		Else 
			Deal=input
		End If
		TestEval=null

	If InStr(input,"if") Then
		If(Right(Trim(LCase(input)),4)<>"then") Then Exit Function 
	End If 
	If InStr(input,"end") Then ControlNum=ControlNum-1
	input=Trimpp(input)
	If InputType<>2 And InStr(input,"do ")Or InStr(input,"if ")Or InStr(input,"sub ")Or InStr(input,"for ") Or InStr(input,"select ")Or InStr(input,"wend ")Or InStr(input,"function ") Then
		NotNeedCheck=False
		ControlNum=ControlNum+1
		InputType=2
	End If 
	input=Null  
End Function

Function IsUAC()
	On Error Resume Next
	FSO.CreateFolder("C:\Windows\TestUAC")
	IsUAC=True
	If Err.Number=70 Then IsUAC=False
	FSO.DeleteFolder("C:\Windows\TestUAC")
	On Error Goto 0 
End Function

Sub Echo(ByVal str)
	If Is_Remote<>True Then 
		WSH.StdOut.Write(str&vbCrLf)
	Else
		Sock.sendData(str&vbCrLf)
	End If
End Sub

Function Input()
	If Is_Remote<>True Then 
		Input=WScript.StdIn.ReadLine
	Else
		Do Until Isrecv=True
			WScript.Sleep 200
		Loop
		ExecuteGlobal("Isrecv=False")
		Input=RecvData
		RecvData=Null 
	End If
End Function

Sub Quit()
	WScript.Quit
End Sub


'---------------------------------function and sub end---------------------------------------------------------


If LCase(Right(WScript.FullName,11))<>"cscript.exe" Then
	Start()
	WScript.Quit
End If

Import("Coreutils.vbs")

If WScript.Arguments.Count=1 Then
	
	Dim Islogo
	Islogo=True
	Select Case Trim(LCase(WScript.Arguments(0)))
	Case "-nl"
		Islogo=False
	Case "-re"
		
	End Select
	If fso.FileExists(WScript.Arguments(0))=True Then
		IsScript=True
		ScriptName=WSH.Arguments(0)
		Set ScriptFile=fso.OpenTextFile(WScript.Arguments(0))
	End If
End If 

If IsScript<>True And Islogo=True Then
	WSH.Echo ProductName&" [�汾 " & VersionName & "]"
	WScript.Echo("�ڲ��汾 "& VersionNum)
	WScript.Echo("��Ȩ���� " & Chr(60) + Chr(99) + Chr(62) & " 2019 " + "TECH_N00b&С����" + Chr(32) + "��������Ȩ����")
	WScript.Echo("")
End If

'-------------------------------input part-----------------------------------------------------
Dim Count,ErrMode
ErrMode=0
InputType=1

Do
	If ErrMode=0 Then On Error Goto 0
	If ErrMode=1 Then On Error Resume Next
	
	Select Case InputType
		Case 1
		If IsScript Then
			If ScriptFile.AtEndOfStream Then WSH.Quit()
			CMD_Input=ScriptFile.ReadLine
		Else
			If rIsUAC= True Then WScript.StdOut.Write(WS.CurrentDirectory & "#>")
			If rIsUAC<>True Then WScript.StdOut.Write(WS.CurrentDirectory & "@>")
			CMD_Input=WScript.StdIn.ReadLine
		End If
		ErrorDealing()
		Case 2
		If IsScript=False Then
			output=Null
			For i=0 To ControlNum
				output=output&"----"
			Next
			WScript.StdOut.Write output
			str=Trim(LCase(WScript.StdIn.ReadLine))
		Else
			If IsAtEnd<>True Then
				str=ScriptFile.ReadLine	
			End If
			IsScript=True
		End If
		CMD_Input=str
		str=Trimpp(str)
		If InStr(str,"next")Or InStr(str,"loop")Or InStr(str,"end ")Then
			ControlNum=ControlNum-1
		End If   
		If InStr(str,"do ")Or InStr(str,"if ") Or InStr(str,"for") Or InStr(str,"select")Or InStr(str,"with ") Or InStr(str,"wend") Then
			If InStr(str,"end if") Or InStr(str,"end function") Then ControlNum=ControlNum-1
			If InStr(input,"if") Then
				If(Right(Trim(LCase(input)),4)<>"then") Then ControlNum=ControlNum-1
			End If 
			ControlNum=ControlNum+1
		End If
		If ControlNum=0 Then 
			NotNeedCheck=True
			NeedExecute=True
		End If
		CMD_Input=str
		str=Null
	End Select
	If InputType=2 Then
		Code=Code&vbCrLf&Deal(CMD_Input)
	Else
		Code=Deal(CMD_Input)
			
	End If
	If ControlNum=0 Then 
		InputType=1
	End If
	If InputType=1 And IsScript<>True  Then
		'--------------------------------------execute part------------------------------------------------------------------------------
		On Error Resume next
		If ScriptType<>"bat" Then 
			Code="'VBS Shell Code here:"&vbCrLf&Code
			Execute(Code)
		Else
			System(Code)
			ScriptType="vbs"
		End If
		CMD_Input=Null
		Code=Null
		If Err.Number<>0 Then 
			WSH.Echo("����:" & Err.Number)
			WScript.Echo(Err.Description)
			WScript.Echo("��֧��vbscript�������ԣ�")
			WScript.Echo("  1.����������������" & Err.Description & "��")
			WScript.Echo("  2.���롱About������ϵ�����ߡ�")
		End If
		On Error Goto 0
		'----------------------------------------execute part end -----------------------------------------------------------------------------
	End If
Loop