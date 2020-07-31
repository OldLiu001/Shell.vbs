
Function fDecode(sStringToDecode)  
	'This function will decode a Base64 encoded string and returns the decoded string.  
	'This becomes usefull when attempting to hide passwords from prying eyes.  
	Const CharList = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"  
	Dim iDataLength, sOutputString, iGroupInitialCharacter  
	sStringToDecode = Replace(Replace(Replace(sStringToDecode, vbCrLf, ""), vbTab, ""), " ", "")  
	iDataLength = Len(sStringToDecode)  
	If iDataLength Mod 4 <> 0 Then  
		fDecode = "Bad string passed to fDecode() function."  
		Exit Function  
	End If  
	For iGroupInitialCharacter = 1 To iDataLength Step 4  
		Dim iDataByteCount, iCharacterCounter, sCharacter, iData, iGroup, sPreliminaryOutString  
		iDataByteCount = 3  
		iGroup = 0  
		For iCharacterCounter = 0 To 3  
			sCharacter = Mid(sStringToDecode, iGroupInitialCharacter + iCharacterCounter, 1)  
			If sCharacter = "=" Then  
				iDataByteCount = iDataByteCount - 1  
				iData = 0  
			Else  
				iData = InStr(1, CharList, sCharacter, 0) - 1  
				If iData = -1 Then  
					fDecode = "Bad string passed to fDecode() function."  
					Exit Function  
				End If  
			End If  
			iGroup = 64 * iGroup + iData  
		Next  
		iGroup = Hex(iGroup)  
		iGroup = String(6 - Len(iGroup), "0") & iGroup  
		sPreliminaryOutString = Chr(CByte("&H" & Mid(iGroup, 1, 2))) & Chr(CByte("&H" & Mid(iGroup, 3, 2))) & Chr(CByte("&H" & Mid(iGroup, 5, 2)))  
		sOutputString = sOutputString & Left(sPreliminaryOutString, iDataByteCount)  
	Next  
	fDecode = sOutputString  
End Function
