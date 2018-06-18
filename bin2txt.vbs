writeTxt readBinary(WScript.Arguments.Item(0)),WScript.Arguments.Item(0) & ".txt"

Function readBinary(strPath)
  Dim oFSO
  Dim oFile

	Set oFSO = CreateObject("Scripting.FileSystemObject")
	Set oFile = oFSO.GetFile(strPath)
	If IsNull(oFile) Then
  	MsgBox("File not found: " & strPath)
	Else
		With oFile.OpenAsTextStream()
      readBinary = .Read(oFile.Size)
    	.Close
  	End With
  End If
End Function

Function writeTxt(strBinary, strPath)
  Dim oFSO
  Dim oTxtStream
  Dim length
  Dim Str1
  Dim Str2

  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oTxtStream = oFSO.createTextFile(strPath)

  length = len(strBinary)
	For i = 0 to length - 1
		If i mod 16 = 0 Then
			If i <> 0 Then
				oTxtStream.Write(chr(13) & chr(10))
			End If
			oTxtStream.Write("block" & Hex(i) & "=")
		End If

		Str1=Mid(strBinary, i + 1, 1)
		Str2=Asc(Str1)
		If i mod 16 = 15 Or i = length - 1 Then
			oTxtStream.Write(Hex(Str2))
		else
			oTxtStream.Write(Hex(Str2) & ",")
		End If

	Next
	oTxtStream.Write(chr(13) & chr(10))
	oTxtStream.Close
  Set oTxtStream = Nothing
End Function