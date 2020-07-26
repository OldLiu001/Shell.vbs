If Not IsActiveX("MSWinsock.Winsock") Then WScript.Echo("Plz install Mswinsck.ocx!"):Exit Function
Public Const Is_Remote=True
Dim Socketobj
Set Socketobj=CreateObject("MSWinsock.Winsock")


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
