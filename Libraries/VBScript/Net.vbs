Function HTTP_TestPort(ByVal IP,ByVal Port)
	On Error Resume Next
	Set x=CreateObject("msxml2.serverXMLHTTP")
	x.open "post","http://"&IP&":"&Port
	x.send("hello")
	If Err.NuMbEr=0 Or Err.NuMbEr=-2147012866 Or Err.NuMbEr=-2147012894 Or Err.NuMbEr=-2147012744 Or Err.NuMbEr=-2147467259 Then
		HTTP_TestPort=True
	Else
		Http_TestPort=False
	End If
	On Error Goto 0
End Function

Function TCP_TestPort(ByVal ip,ByVal port,Int_TimeOut)
	Set sock=CreateObject("MSWinsock.Winsock")
	sock.Connect IP,port
	WScript.Sleep(Int_TimeOut)
	If sock.State()=7 Then 
		TCP_TestPort=True
	Else
		TCP_TestPort=False
	End If
	Set sock=Nothing
End Function

Sub Client(IP,Port)
	
End Sub
