Dim HTTP,FSO,AStr
Set WMI=GetObject("WinMgmts:\\.\root\cimv2")

Set http = CreateObject("Msxml2.ServerXMLHTTP.6.0")

Set Astr = CreateObject("ADODB.Stream")
Set FSO=CreateObject("Scripting.FileSystemObject")

'------------------------------File Operation--------------------------------

Sub RM(File_Path)
	If fso.FileExists(File_Path) Then FSO.DeleteFile File_Path,True
	If fso.FolderExists(File_Path) Then fso.DeleteFolder File_Path,True 
End Sub

Sub WriteFile(FilePath,WriteStr)
	If FSO.FileExists(FilePath) Then 
		Set	file=FSO.OpenTextFile(FilePath)
	Else
		Set file=FSO.CreateTextFile(FilePath)
	End If 
	file.Write(WriteStr)
	file.Close
End Sub

Function LS()
	Set oFolder = fso.GetFolder(ws.CurrentDirectory)     '获取文件夹
	Set oSubFolders = oFolder.SubFolders    '获取子目录集合
	
	For Each folder In oSubFolders
		LS=LS&("#"&folder&"#")&vbCrLf
	Next
	
	Set oFiles = oFolder.Files              '获取文件集合
	For Each file In oFiles
		LS=LS&(file.Name&"	"& FSO.GetFile(file.Name).size)&vbCrLf
	Next
End Function

Sub CP(FileSpec,destination,NeedCovered)
	If destination=Null Then destination=PWD()
	If FSO.FolderExists(destination) Then destination=destination&"\"&FileSpec
	
	If NeedCovered=Null Then 
		WriteFile destination,Cat(FilePath)
	Else 
		If FSO.FileExists(destination) Then WriteFile destination&" - Copyed",Cat(FilePath)
	End If
End Sub

'--------------------------------Folder----------------------------------

Function MkDir(FolderName)
	FSO.CreateFolder(Folder)
End Function


'---------------------------------Network Tools-----------------------------------

Function Wget(ByVal url,FilePath)
	If FilePath=Null Then FilePath=FSO.GetFile(url)
	adTypeBinary = 1

	adSaveCreateOverWrite = 2

	HTTP.SetOption 2, 13056
	
http.open "GET",url,False
	
http.send
	
AStr.Type = adTypeBinary
	
AStr.Open

	AStr.Write http.responseBody
	
AStr.SaveToFile FilePath
	
AStr.Close
End Function

Function GetIP() '取本机IP
Set colItems = WMI.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
For Each objItem in colItems
    For Each objAddress in objItem.IPAddress
        If objAddress <> "" then
            GetIP = objAddress
            Exit For
        End If
    Next
Next
End Function

Function TestPort(IP,Port,Timeout)
	If Not IsActiveX("MSWinsock.Winsock") Then WScript.Echo("Plz install Mswinsck.ocx!"):Exit Function
	Set Sock=CreateObject("MSWinsock.Winsock")
	sock.Connect IP,Port
	WScript.Sleep Int(Timeout)
	If sock.State()=7 Then 
		TestPort=True
	Else
		TestPort=False
	End If
End Function

Function Ping(IP)'Ping(IP)
	Set colItems = WMI.ExecQuery("Select * From Win32_PingStatus Where Address='" & IP & "'") 
	For Each objItem In colItems 
		Ping=objItem.StatusCode
	Next 
End Function

'--------------------------------System Managemant----------------------------------

Sub Shutdown()
	Set objWMIService = GetObject("winmgmts:" _ 
	& "{impersonationLevel=impersonate,(Shutdown)}!\\.\root\cimv2") 
	Set colOperatingSystems = objWMIService.ExecQuery _ 
	("Select * from Win32_OperatingSystem") 
	For Each objOperatingSystem In colOperatingSystems 
		ObjOperatingSystem.ShutDown() 
	Next 
End Sub

Function IsActiveX(ComName)
	On Error Resume Next
    Dim O
    Set O = CreateObject(ComName)
    If Err.Number = 0 Then
        IsActiveX = True
    End If
    Set O = Nothing
    Err.Clear
    On Error Goto 0
End Function


Sub ReBoot()
	Set objWMIService = GetObject("winmgmts:" _ 
	& "{impersonationLevel=impersonate,(Shutdown)}!\\.\root\cimv2") 
	OperatingSystems = objWMIService.ExecQuery _ 
	("Select * from Win32_OperatingSystem") 
	For Each objOperatingSystem In colOperatingSystems 
		ObjOperatingSystem.Reboot() 
	Next 
End Sub

Function IsProc(ProcessTag)
	Dim Processes, Process
	IsProc = False
	Set Processes = wmi.ExecQuery("SELECT * FROM Win32_Process",,48) 
	For Each Process In Processes
		If IsNumeric(Process) Then
			If CStr(Process.ProcessId) = ProcessTag Then IsProc = True
		Else
			If Process.Name = ProcessTag Then IsProc = True
		End If
	Next
End Function


Function Kill(Process)'0:找不到。1:成功了。2：失败了。
	If InStr(Process,"*") Then NeedSearch=True

	If NeedSearch=True Then Kill=3
	If Kill=0 Then Exit Function 
	If Not IsNumeric(Process) Then
		If NeedSearch<>True Then 
			Set colItems = wmi.ExecQuery("SELECT * FROM Win32_Process where name='"&Process&"' ",,48) 
			For Each objItem In colItems  
				objItem.Terminate()
			Next
			Set colItems = Nothing
			Exit Function
		Else 
			Process=Replace(Process,"*","")
			Set WMI=GetObject("WinMgmts:")
			Set Objs=WMI.InstancesOf("Win32_Process")
			GetProcess=""
			For Each Obj In Objs
				If InStr(LCase(Obj.Description),LCase(Process)) Then Echo("Killing "&Obj.Description&", PID:"&Obj.ProcessID):Obj.Terminate
			Next
			NeedSearch=false
		End If 
	Else
		Set WMI=GetObject("WinMgmts:")
		Set Objs=WMI.InstancesOf("Win32_Process")
		GetProcess=""
		For Each Obj In Objs
			If Process=Obj.ProcessId Then Obj.Terminate()
		Next
	End If 
	If IsProc(Process) Then Kill=2
End Function

Function PS()'List out Process
	Set WMI=GetObject("WinMgmts:")
	Set Objs=WMI.InstancesOf("Win32_Process")
	GetProcess=""
	For Each Obj In Objs
		PS=PS&(GetProcess&Obj.Description&Chr(9)&Obj.ProcessId)&vbCrLf
	Next
End Function 

'------------------------------------Baisc Commands--------------------------------
Function WhoAmI()
	 WhoAmI=CreateObject( "WScript.Network" ).ComputerName & "\" & CreateObject( "WScript.Network" ).UserName
End Function

Function PWD()
	PWD=ws.CurrentDirectory
End Function

Function ReStart()
	ws.run WScript.ScriptFullName
	Quit()
End Function

Function Env(Var)
	Env=WS.ExpandEnvironmentStrings("%"&Var&"%")
End Function