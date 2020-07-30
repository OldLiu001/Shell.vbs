Function HighLight(strToHighlight)
	Rem Vbs高亮函数，老刘编写。
	Rem http://www.bathome.net/thread-47323-1-1.html
	Rem HTML特殊字符及标签处理感谢bbaa。
	Rem 灵感 From Demon's Vbs-Beautifier
	Rem 此版本为保存0x5字符做了特殊修改。
	
	Rem 常量设置
	Dim STRING_FLAG,COMMENT_FLAG,BLANK_FLAG,SPECIAL_CHAR_FLAG
	Dim [符号集合],[保留字集合],[内置函数集合],[内置常量集合]
	STRING_FLAG = Chr(1)
	COMMENT_FLAG = Chr(2)
	BLANK_FLAG = Chr(3)
	SPECIAL_CHAR_FLAG = Chr(4)
	CURSOR_FLAG = Chr(5) '为VBS Shell新增的光标标记
	[符号集合] = ",./\()<=>+-*^&"
	[保留字集合] = Split("And As Boolean ByRef Byte ByVal Call Case Class Const Currency Debug Dim Do Double Each Else ElseIf Empty End EndIf Enum Eqv Event Exit Explicit False For Function Get Goto If Imp Implements In Integer Is Let Like Long Loop LSet Me Mod New Next Not Nothing Null On Option Optional Or ParamArray Preserve Private Property Public RaiseEvent ReDim Resume RSet Select Set Shared Single Static Stop Sub Then To True Type TypeOf Until Variant WEnd While With Xor"," ")
	[内置函数集合] = Split("Abs Array Asc Atn CBool CByte CCur CDate CDbl CInt CLng CSng CStr Chr Cos CreateObject Date DateAdd DateDiff DatePart DateSerial DateValue Day Escape Eval Exp Filter Fix FormatCurrency FormatDateTime FormatNumber FormatPercent GetLocale GetObject GetRef Hex Hour InStr InStrRev InputBox Int IsArray IsDate IsEmpty IsNull IsNumeric IsObject Join LBound LCase LTrim Left Len LoadPicture Log Mid Minute Month MonthName MsgBox Now Oct Randomize RGB RTrim Replace Right Rnd Round ScriptEngine ScriptEngineBuildVersion ScriptEngineMajorVersion ScriptEngineMinorVersion Second SetLocale Sgn Sin Space Split Sqr StrComp StrReverse String Tan Time TimeSerial TimeValue Timer Trim TypeName UBound UCase Unescape VarType Weekday WeekdayName Year"," ")
	[内置常量集合] = Split("vbBlack vbRed vbGreen vbYellow vbBlue vbMagenta vbCyan vbWhite vbBinaryCompare vbTextCompare vbSunday vbMonday vbTuesday vbWednesday vbThursday vbFriday vbSaturday vbUseSystemDayOfWeek vbFirstJan1 vbFirstFourDays vbFirstFullWeek vbGeneralDate vbLongDate vbShortDate vbLongTime vbShortTime vbObjectError vbOKOnly vbOKCancel vbAbortRetryIgnore vbYesNoCancel vbYesNo vbRetryCancel vbCritical vbQuestion vbExclamation vbInformation vbDefaultButton1 vbDefaultButton2 vbDefaultButton3 vbDefaultButton4 vbApplicationModal vbSystemModal vbOK vbCancel vbAbort vbRetry vbIgnore vbYes vbNo vbCr vbCrLf vbFormFeed vbLf vbNewLine vbNullChar vbNullString vbTab vbVerticalTab vbUseDefault vbTrue vbFalse vbEmpty vbNull vbInteger vbLong vbSingle vbDouble vbCurrency vbDate vbString vbObject vbError vbBoolean vbVariant vbDataObject vbDecimal vbByte vbArray WScript Wsh"," ")
	
	Rem 正则对象初始化
	Dim re
	Set re = New RegExp
	re.Global = True
	re.IgnoreCase = True
	re.MultiLine = False
	
	Rem 读取代码
	Dim strCode
	strCode = strToHighlight
	
	Rem HTML特殊符号预处理
	Dim [着色标签],[换行标签],[空白字符]
	[着色标签] = "<span style=""color:|ReplaceHere|;"">$1</span>"
	[换行标签] = "<br>"
	[空白字符] = "&nbsp;"
	strCode = Replace(strCode,"&",SPECIAL_CHAR_FLAG&"amp;")
	strCode = Replace(strCode,">",SPECIAL_CHAR_FLAG&"gt;")
	strCode = Replace(strCode,"<",SPECIAL_CHAR_FLAG&"lt;")
	
	Rem 预处理字符串
	Dim [字符串集合]
	re.Pattern = """[.\x05]*?"""
	Set [字符串集合] = re.Execute(strCode)
	strCode = re.Replace(strCode, STRING_FLAG)
	
	Rem 预处理空字符
	strCode = Replace(strCode,Chr(9),"    ")
	strCode = Replace(strCode," ",BLANK_FLAG)
	
	Rem 预处理换行
	strCode = Replace(strCode,vbNewLine,vbCr)
	strCode = Replace(strCode,vbLf,vbCr)
	
	Rem 预处理注释
	Dim [注释集合]
	re.Pattern = "((?:\x05?[\x03\x05]*R\x05?e\x05?m\x05?\x03+\x05?|'\x05?)[^\r]*)" '在此严重的感谢bbaa指导
	Set [注释集合] = re.Execute(strCode)
	strCode = re.Replace(strCode, COMMENT_FLAG)
	
	Rem 添加着色标签以及HTML特殊符号处理
	With re
		Rem 偷懒操作，用正则将符号集合替换为正则表达式，再用替换出来的正则表达式处理strCode。
		Rem 下面三行的代码完成了 ",./\()<=>+-*&^" ==> "(\,|\.|\/|\\|\(|\)|\<|\=|\>|\+|\-|\*|\&|\^)"
		.Pattern = ""
		.Pattern = re.Replace([符号集合],"|\")
		.Pattern = "(" & Left(Right(.Pattern,Len(.Pattern) - 1),Len(.Pattern) - 3) & ")"
		strCode = .Replace(StrCode,Replace([着色标签],"|ReplaceHere|",GetColor("typePunctuation")))
	End With
	
	strCode = Replace(strCode,SPECIAL_CHAR_FLAG & "amp;", "<span style=""color:OrangeRed;"">"&SPECIAL_CHAR_FLAG&"amp;"&"</span>")
	strCode=Replace(strCode,SPECIAL_CHAR_FLAG&"gt;","<span style=""color:OrangeRed;"">"&SPECIAL_CHAR_FLAG&"gt;"&"</span>")
	strCode=Replace(strCode,SPECIAL_CHAR_FLAG&"lt;","<span style=""color:OrangeRed;"">"&SPECIAL_CHAR_FLAG&"lt;"&"</span>")
	
	Dim [保留字]
	For Each [保留字] In [保留字集合]
		re.Pattern = "\b("&CursorSupport([保留字])&")\b"
		strCode = re.Replace(strCode, Replace([着色标签],"|ReplaceHere|",GetColor("typePreserved")))
	Next
	
	Dim [内置函数]
	For Each [内置函数] In [内置函数集合]
		re.Pattern = "\b("&CursorSupport([内置函数])&")\b"
		strCode = re.Replace(strCode, Replace([着色标签],"|ReplaceHere|",GetColor("typeFunction")))
	Next
	
	Dim [内置常量]
	For Each [内置常量] In [内置常量集合]
		re.Pattern = "\b("&CursorSupport([内置常量])&")\b"
		strCode = re.Replace(strCode, Replace([着色标签],"|ReplaceHere|",GetColor("typeConst")))
	Next
	
	Rem 处理注释
	Dim [注释]
	For Each [注释] In [注释集合]
		strCode = Replace(strCode, COMMENT_FLAG, _
		Replace(Replace([着色标签],"|ReplaceHere|",GetColor("typeComment")),"$1",[注释]), 1, 1) 'or #00ff00
	Next
	
	Rem 处理字符串
	Dim [字符串]
	For Each [字符串] In [字符串集合]
		strCode = Replace(strCode, STRING_FLAG, _
		Replace(Replace([着色标签],"|ReplaceHere|",GetColor("typeString")),"$1",[字符串]), 1, 1)
	Next
	
	Rem 处理换行和空字符
	strCode = Replace(strCode,vbCr,[换行标签])
	strCode = Replace(strCode,BLANK_FLAG,[空白字符])
	strCode = Replace(strCode,SPECIAL_CHAR_FLAG,Chr(&H26))
	
	Rem 处理完成。
	Highlight = strCode
End Function

Function CursorSupport(strKeyword) '为了保存光标位置而设计的正则修改函数
	CursorSupport = "\x05?"
	Dim lngPtr
	For lngPtr = 1 To Len(strKeyword)
		CursorSupport = CursorSupport & Mid(strKeyword,lngPtr,1) & "\x05?"
	Next
End Function