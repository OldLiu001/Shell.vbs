Const OWN_PROCESS = &H10
Const ERR_CONTROL = &H2
Const INTERACTIVE = False

Set objWMIService = GetObject("winmgmts:" & "\\.\root\cimv2") 

Function Serv_List()
	Set colServices = objWMIService.ExecQuery("Select * from Win32_Service") 
	i=0
	For Each objService In colServices 
		Serv_List=Serv_List&objService.DisplayName & "|" & objService.State&vbCrLf
	Next
	Set colServices =Nothing 
End Function

Function Serv_Exist(Serv_Name)
	Set colServices = objWMIService.ExecQuery("Select * from Win32_Service Where Name='"&Serv_Name&"'") 
	Serv_Eist=false
	For Each objService In colServices 
		Serv_Exist=True
	Next
	Set colServices =Nothing
End Function


Function Serv_Create(ServiceName,DisPlayName,InstallPath)
	Set ObjWMI=objWMIService
	Set ObjSvr = ObjWMI.Get("Win32_Service")
	Serv_Return = ObjSvr.Create(ServiceName, DisplayName, InstallPath, OWN_PROCESS, ERR_CONTROL, "Automatic", INTERACTIVE, "LocalSystem", "")
	If(Serv_Return = 0) Then
		Serv_Create=True
	Else
		Serv_Create=Serv_Return
	End If
	
	Set ObjWMI=Nothing
	Set ObjSvr=Nothing
End Function

Function Serv_Delete(Serv_Name)
	Set colListOfServices = objWMIService.ExecQuery("Select * from Win32_Service Where Name = '"&Serv_Name&"'")
	For Each objService in colListOfServices
    	objService.StopService()
    	objService.Delete()
	Next
	Set colListOfServices = Nothing
End Function