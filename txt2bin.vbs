writeBinary readTxt(WScript.Arguments.Item(0)), Left(WScript.Arguments.Item(0), Len(WScript.Arguments.Item(0)) - 4)

Function readTxt(strPath)
  Dim oFSO
  Dim oFile
	Set oFSO = CreateObject("Scripting.FileSystemObject")
	Set oFile = oFSO.GetFile(strPath)
  If IsNull(oFile) Then MsgBox("File not found: " & strPath) : Exit Function

  With oFile.OpenAsTextStream()
		readTxt = .Read(oFile.Size)
    .Close
  End With
End Function

Function writeBinary(strText, strPath)
  const vbBinaryCompare = 0

  Dim strBinary
  Dim oFSO
  Dim oTxtStream
  Dim i, j, k
  Dim strBinTxtLine
  Dim myLong

  Set oFSO = CreateObject("Scripting.FileSystemObject")
  Set oTxtStream = oFSO.createTextFile(strPath)
  j = 1

  Do
  	i = instr(j, strText, "=", vbBinaryCompare)
  	j = instr(i + 1, strText, chr(10), vbBinaryCompare)
  	if i > 0 Then
	  	strBinTxtLine = mid(strText, i + 1, j - i - 1)

			'wscript.echo i & ":" & k & ":" & strBinTxtLine
			For Each strHex in Split(strBinTxtLine, ",")
			  myLong = CLng("&h" & strHex)
    		strBinary = strBinary & Chr(myLong)
 			Next
    End If
  Loop Until i = 0

  oTxtStream.Write(strBinary)
	oTxtStream.Close
  Set oTxtStream = Nothing

End Function

