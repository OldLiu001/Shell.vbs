Function HighLight(strToHighlight)
	Rem Vbs����������������д��
	Rem http://www.bathome.net/thread-47323-1-1.html
	Rem HTML�����ַ�����ǩ�����лbbaa��
	Rem ��� From Demon's Vbs-Beautifier
	Rem �˰汾Ϊ����0x5�ַ����������޸ġ�
	
	Rem ��������
	Dim STRING_FLAG,COMMENT_FLAG,BLANK_FLAG,SPECIAL_CHAR_FLAG
	Dim [���ż���],[�����ּ���],[���ú�������],[���ó�������]
	STRING_FLAG = Chr(1)
	COMMENT_FLAG = Chr(2)
	BLANK_FLAG = Chr(3)
	SPECIAL_CHAR_FLAG = Chr(4)
	CURSOR_FLAG = Chr(5) 'ΪVBS Shell�����Ĺ����
	[���ż���] = ",./\()<=>+-*^&"
	[�����ּ���] = Split("And As Boolean ByRef Byte ByVal Call Case Class Const Currency Debug Dim Do Double Each Else ElseIf Empty End EndIf Enum Eqv Event Exit Explicit False For Function Get Goto If Imp Implements In Integer Is Let Like Long Loop LSet Me Mod New Next Not Nothing Null On Option Optional Or ParamArray Preserve Private Property Public RaiseEvent ReDim Resume RSet Select Set Shared Single Static Stop Sub Then To True Type TypeOf Until Variant WEnd While With Xor"," ")
	[���ú�������] = Split("Abs Array Asc Atn CBool CByte CCur CDate CDbl CInt CLng CSng CStr Chr Cos CreateObject Date DateAdd DateDiff DatePart DateSerial DateValue Day Escape Eval Exp Filter Fix FormatCurrency FormatDateTime FormatNumber FormatPercent GetLocale GetObject GetRef Hex Hour InStr InStrRev InputBox Int IsArray IsDate IsEmpty IsNull IsNumeric IsObject Join LBound LCase LTrim Left Len LoadPicture Log Mid Minute Month MonthName MsgBox Now Oct Randomize RGB RTrim Replace Right Rnd Round ScriptEngine ScriptEngineBuildVersion ScriptEngineMajorVersion ScriptEngineMinorVersion Second SetLocale Sgn Sin Space Split Sqr StrComp StrReverse String Tan Time TimeSerial TimeValue Timer Trim TypeName UBound UCase Unescape VarType Weekday WeekdayName Year"," ")
	[���ó�������] = Split("vbBlack vbRed vbGreen vbYellow vbBlue vbMagenta vbCyan vbWhite vbBinaryCompare vbTextCompare vbSunday vbMonday vbTuesday vbWednesday vbThursday vbFriday vbSaturday vbUseSystemDayOfWeek vbFirstJan1 vbFirstFourDays vbFirstFullWeek vbGeneralDate vbLongDate vbShortDate vbLongTime vbShortTime vbObjectError vbOKOnly vbOKCancel vbAbortRetryIgnore vbYesNoCancel vbYesNo vbRetryCancel vbCritical vbQuestion vbExclamation vbInformation vbDefaultButton1 vbDefaultButton2 vbDefaultButton3 vbDefaultButton4 vbApplicationModal vbSystemModal vbOK vbCancel vbAbort vbRetry vbIgnore vbYes vbNo vbCr vbCrLf vbFormFeed vbLf vbNewLine vbNullChar vbNullString vbTab vbVerticalTab vbUseDefault vbTrue vbFalse vbEmpty vbNull vbInteger vbLong vbSingle vbDouble vbCurrency vbDate vbString vbObject vbError vbBoolean vbVariant vbDataObject vbDecimal vbByte vbArray WScript Wsh"," ")
	
	Rem ��������ʼ��
	Dim re
	Set re = New RegExp
	re.Global = True
	re.IgnoreCase = True
	re.MultiLine = False
	
	Rem ��ȡ����
	Dim strCode
	strCode = strToHighlight
	
	Rem HTML�������Ԥ����
	Dim [��ɫ��ǩ],[���б�ǩ],[�հ��ַ�]
	[��ɫ��ǩ] = "<span style=""color:|ReplaceHere|;"">$1</span>"
	[���б�ǩ] = "<br>"
	[�հ��ַ�] = "&nbsp;"
	strCode = Replace(strCode,"&",SPECIAL_CHAR_FLAG&"amp;")
	strCode = Replace(strCode,">",SPECIAL_CHAR_FLAG&"gt;")
	strCode = Replace(strCode,"<",SPECIAL_CHAR_FLAG&"lt;")
	
	Rem Ԥ�����ַ���
	Dim [�ַ�������]
	re.Pattern = """[.\x05]*?"""
	Set [�ַ�������] = re.Execute(strCode)
	strCode = re.Replace(strCode, STRING_FLAG)
	
	Rem Ԥ������ַ�
	strCode = Replace(strCode,Chr(9),"    ")
	strCode = Replace(strCode," ",BLANK_FLAG)
	
	Rem Ԥ������
	strCode = Replace(strCode,vbNewLine,vbCr)
	strCode = Replace(strCode,vbLf,vbCr)
	
	Rem Ԥ����ע��
	Dim [ע�ͼ���]
	re.Pattern = "((?:\x05?[\x03\x05]*R\x05?e\x05?m\x05?\x03+\x05?|'\x05?)[^\r]*)" '�ڴ����صĸ�лbbaaָ��
	Set [ע�ͼ���] = re.Execute(strCode)
	strCode = re.Replace(strCode, COMMENT_FLAG)
	
	Rem �����ɫ��ǩ�Լ�HTML������Ŵ���
	With re
		Rem ͵�������������򽫷��ż����滻Ϊ������ʽ�������滻������������ʽ����strCode��
		Rem �������еĴ�������� ",./\()<=>+-*&^" ==> "(\,|\.|\/|\\|\(|\)|\<|\=|\>|\+|\-|\*|\&|\^)"
		.Pattern = ""
		.Pattern = re.Replace([���ż���],"|\")
		.Pattern = "(" & Left(Right(.Pattern,Len(.Pattern) - 1),Len(.Pattern) - 3) & ")"
		strCode = .Replace(StrCode,Replace([��ɫ��ǩ],"|ReplaceHere|",GetColor("typePunctuation")))
	End With
	
	strCode = Replace(strCode,SPECIAL_CHAR_FLAG & "amp;", "<span style=""color:OrangeRed;"">"&SPECIAL_CHAR_FLAG&"amp;"&"</span>")
	strCode=Replace(strCode,SPECIAL_CHAR_FLAG&"gt;","<span style=""color:OrangeRed;"">"&SPECIAL_CHAR_FLAG&"gt;"&"</span>")
	strCode=Replace(strCode,SPECIAL_CHAR_FLAG&"lt;","<span style=""color:OrangeRed;"">"&SPECIAL_CHAR_FLAG&"lt;"&"</span>")
	
	Dim [������]
	For Each [������] In [�����ּ���]
		re.Pattern = "\b("&CursorSupport([������])&")\b"
		strCode = re.Replace(strCode, Replace([��ɫ��ǩ],"|ReplaceHere|",GetColor("typePreserved")))
	Next
	
	Dim [���ú���]
	For Each [���ú���] In [���ú�������]
		re.Pattern = "\b("&CursorSupport([���ú���])&")\b"
		strCode = re.Replace(strCode, Replace([��ɫ��ǩ],"|ReplaceHere|",GetColor("typeFunction")))
	Next
	
	Dim [���ó���]
	For Each [���ó���] In [���ó�������]
		re.Pattern = "\b("&CursorSupport([���ó���])&")\b"
		strCode = re.Replace(strCode, Replace([��ɫ��ǩ],"|ReplaceHere|",GetColor("typeConst")))
	Next
	
	Rem ����ע��
	Dim [ע��]
	For Each [ע��] In [ע�ͼ���]
		strCode = Replace(strCode, COMMENT_FLAG, _
		Replace(Replace([��ɫ��ǩ],"|ReplaceHere|",GetColor("typeComment")),"$1",[ע��]), 1, 1) 'or #00ff00
	Next
	
	Rem �����ַ���
	Dim [�ַ���]
	For Each [�ַ���] In [�ַ�������]
		strCode = Replace(strCode, STRING_FLAG, _
		Replace(Replace([��ɫ��ǩ],"|ReplaceHere|",GetColor("typeString")),"$1",[�ַ���]), 1, 1)
	Next
	
	Rem �����кͿ��ַ�
	strCode = Replace(strCode,vbCr,[���б�ǩ])
	strCode = Replace(strCode,BLANK_FLAG,[�հ��ַ�])
	strCode = Replace(strCode,SPECIAL_CHAR_FLAG,Chr(&H26))
	
	Rem ������ɡ�
	Highlight = strCode
End Function

Function CursorSupport(strKeyword) 'Ϊ�˱�����λ�ö���Ƶ������޸ĺ���
	CursorSupport = "\x05?"
	Dim lngPtr
	For lngPtr = 1 To Len(strKeyword)
		CursorSupport = CursorSupport & Mid(strKeyword,lngPtr,1) & "\x05?"
	Next
End Function