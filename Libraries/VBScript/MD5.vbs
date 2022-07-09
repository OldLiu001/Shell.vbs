Function GetFileHash(file_name)
	Set wi = CreateObject("WindowsInstaller.Installer")
	Dim file_hash
	Dim hash_value
	Dim i
	Set file_hash = wi.FileHash(file_name, 0)
	hash_value = ""
	For i = 1 To file_hash.FieldCount
		hash_value = hash_value & BigEndianHex(file_hash.IntegerData(i))
	Next
	GetFileHash = hash_value
	Set file_hash = Nothing
End Function

Function BigEndianHex(Int)
	Dim result
	Dim b1, b2, b3, b4
	result = Hex(Int)
	b1 = Mid(result, 7, 2)
	b2 = Mid(result, 5, 2)
	b3 = Mid(result, 3, 2)
	b4 = Mid(result, 1, 2)
	BigEndianHex = b1 & b2 & b3 & b4
End Function