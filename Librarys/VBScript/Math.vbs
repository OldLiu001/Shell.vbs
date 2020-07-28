Function IsDevidedEventally(NumberBeDevided,DevideNum)
	If(NumberBeDevided Mod DevideNum)=0 Then IsDevidedEventally=True:Exit Function
	IsDevidedEventally=False
End Function

Function DevideInInt(ByVal Num)
	Num=Abs(Num)
	Do Until i=1
		For i=2 To Int(Num/2)
			If IsDevidedEventally(Num,i) Then 
				DevideInInt=i&"*"&DevideInInt
				Num=Num/i
				MsgBox(Num)
				Exit For
			End If
		Next
		If i=Int(Num/2) Then i=1
	Loop
	DevideInInt=Split(DevideInInt,"*")
End Function

For intc=0 To UBound(DevideInInt(12))
	WScript.Echo(DevideInInt(12)(intc))
Next