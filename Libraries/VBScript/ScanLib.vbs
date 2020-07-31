Set WMI=GetObject("WinMgmts:\\.\root\cimv2")
Function Ping(IP)'Ping(IP)
	Set colItems = WMI.ExecQuery("Select * From Win32_PingStatus Where Address='" & IP & "'") 
	For Each objItem In colItems 
		Ping=objItem.StatusCode
	Next 
End Function

Function GetIP() 'È¡±¾»úIP
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

Function Scan_INP()
	Dim GW,SIP
	GW=Split(GetIP(),".")
	For i=0 To 3
		SIP=SIP&GW(i)&"."
	Next
	MsgBox(SIP)
	
	For i=0 To 255
		Scan_INP=Scan_INP&SIP&i&"	"&Ping(SIP&i)&vbCrLf
	Next
End Function

WScript.Echo(Scan_INP())