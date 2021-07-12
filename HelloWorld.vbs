filename = CurDir + "\random.fxml"

Function readTextFile()
	Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(filename,1)
	readTextFile = objFileToRead.ReadAll()
	objFileToRead.Close
	Set objFileToRead = Nothing
End Function

Function clearEmptyLines()
	Const ForReading = 1
	Const ForWriting = 2

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile(filename, ForReading)
	Do Until objFile.AtEndOfStream
		strLine = objFile.Readline
		strLine = Trim(strLine)
		If Len(strLine) > 0 Then
			strNewContents = strNewContents & strLine & vbCrLf
		End If
	Loop

	objFile.Close
	Set objFile = objFSO.OpenTextFile(filename, ForWriting)
	objFile.Write strNewContents
	objFile.Close
End Function

Function encode(text)
	for b = 1 to len(text)
		enText = Mid(text, b, 1)
		enText = Chr(Asc(enText)+3)
		encodedText = encodedText + enText
	next
	encode= encodedText
End Function

Function decode(text)
	for b = 1 to len(text)
		deText = Mid(text, b, 1)
		deText = Chr(Asc(deText)-3)
		decodedText = decodedText + deText
	next
	decode= decodedText
End Function

Function appendToFile(sText)
	Const ForAppending = 8
	Dim OutputFileName, FileSystemObject, InputFile, OutputFile
	Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
	OutputFileName = filename
	If Not FileSystemObject.FileExists(OutputFileName) Then
        msgbox "Output File Not Found"
        WScript.Quit
    End If
	Set OutputFile = FileSystemObject.OpenTextFile(OutputFileName, ForAppending, True)
	OutputFile.WriteLine sText
	OutputFile.Close
	clearEmptyLines()
End Function

Function writeToTextFile(data)
	Set objFileToWrite = CreateObject("Scripting.FileSystemObject").OpenTextFile(filename,2,true)
	objFileToWrite.WriteLine data
	objFileToWrite.Close
	Set objFileToWrite = Nothing
	clearEmptyLines()
End Function

Function getAppName(arrData)
	fieldNumber = ""
	for a = 1 to Ubound(arrData)-1
		strName = decode(Replace(Split(arrData(a),"~")(0), vbCrLf, ""))
		if len(fieldNumber) > 0 then
			fieldNumber = fieldNumber & " " & a & "." & strName
		else
			fieldNumber = a & "." & strName
		end if
	next
	getAppName = fieldNumber
End Function

sCheck = InputBox("Enter Password :", "Welcome to Password Manager")
If sCheck = decode("3:37<3") then

	sAction = InputBox("Enter Number 1. Add 2. View 3. Update 4. Delete", "Your Password Manager")

	Select case sAction
		case "1":
			sName = InputBox ("Enter the Application Name:", "Your Password Repository")
			sPassword = InputBox("Enter the Application Password:", "Add to your repository")
			appendToFile (encode(sName) & "~" & encode(sPassword) & "#")
			If Err.Number = 0 then
				Msgbox "Data Added Successfully"
			Else
				msgbox Err.Description
			End If
		case "2":
			data = readTextFile()
			arrData = Split(data, "#")
			sField = InputBox ("Enter App Name for getting password: " & getAppName(arrData), "View your password")
			strPass = Replace(Split(arrData(sField),"~")(1), vbCrLf, "")
			msgbox decode(strPass)
		case "3":
			data = readTextFile()
			arrData = Split(data, "#")
			sField = InputBox ("Enter App Name for updating password: " & getAppName(arrData), "Update password in your repository")
			strPass = Replace(Split(arrData(sField),"~")(1), vbCrLf, "")
			newPass = encode(InputBox("Enter New Password for " & decode(Replace(Split(arrData(sField),"~")(0), vbCrLf, "")), "Update Password"))
			fName = Split(arrData(sField),"~")(0)
			fVal = Split(arrData(sField),"~")(1)
			data = replace(data, fName &"~" & fVal, fName &"~" &newPass)
			writeToTextFile(data)
			If Err.Number = 0 then
				Msgbox "Data Updated Successfully"
			Else
				msgbox Err.Description
			End If
		case "4":
			data = readTextFile()
			arrData = Split(data, "#")
			sField = InputBox ("Enter App Name to delete: " & getAppName(arrData), "Delete your records")
			fName = Split(arrData(sField),"~")(0)
			fVal = Split(arrData(sField),"~")(1)
			data = replace(data, fName &"~" & fVal &"#","")
			writeToTextFile(data)
			If Err.Number = 0 then
				Msgbox "Data Deleted Successfully"
			Else
				msgbox Err.Description
			End If
		case else:
			Msgbox "Invalid Input"
	End Select
Else
	Msgbox "Hello Rajeev!"
End If
 No newline at end of file
