Function InstrCount(ByVal Str,ByVal Findstr)
	If Len(str)<Len(Fisnstr) Then Exit Function	
	intc=0
	i=0
	Do Until i=Len(str)-Len(Findstr)+1
		If Cut(str,i,Len(Findstr))=Findstr Then 
			intc=intc+1
			i=i+Len(Findstr)
		Else
			i=i+1
		End If
	Loop
	InstrCount=intc
End Function

Function Cut(ByVal Str,StartN,Length)
	Cut=Left(Right(Str,Len(str)-StartN),Length)
End Function 
