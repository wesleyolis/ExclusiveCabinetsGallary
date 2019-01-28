
<% 
'-------------------------------------------------------------------------------
' Functions for default date format
' ANamedFormat = 0-7, where 0-4 same as VBScript
' 5 = "yyyy/mm/dd"
' 6 = "mm/dd/yyyy"
' 7 = "dd/mm/yyyy"

Const EW_DATE_SEPARATOR = "/"

Function EW_FormatDateTime(ADate, ANamedFormat)
  If IsDate(ADate) Then
		If ANamedFormat >= 0 And ANamedFormat <= 4 Then
			EW_FormatDateTime = FormatDateTime(ADate, ANamedFormat)
		ElseIf ANamedFormat = 5 Then
			EW_FormatDateTime = Year(ADate) & EW_DATE_SEPARATOR & Month(ADate) & EW_DATE_SEPARATOR & Day(ADate)
		ElseIf ANamedFormat = 6 Then
			EW_FormatDateTime = Month(ADate) & EW_DATE_SEPARATOR & Day(ADate) & EW_DATE_SEPARATOR & Year(ADate)
		ElseIf ANamedFormat = 7 Then
			EW_FormatDateTime = Day(ADate) & EW_DATE_SEPARATOR & Month(ADate) & EW_DATE_SEPARATOR & Year(ADate)
		Else
			EW_FormatDateTime = ADate
		End If
	Else
		EW_FormatDateTime = ADate
  End If
End Function

Function EW_UnFormatDateTime(ADate, ANamedFormat)

	Dim arDateTime, arDate, AYear, AMonth, ADay

	ADate = Trim(ADate)
	While Instr(ADate, "  ") > 0
		ADate = Replace(ADate, "  ", " ")
	Wend
	arDateTime = Split(ADate, " ")
	If UBound(arDateTime) < 0 Then
		EW_UnFormatDateTime = ADate
		Exit Function
	End If
	arDate = Split(arDateTime(0), EW_DATE_SEPARATOR)
	If UBound(arDate) = 2 Then
		If ANamedFormat = 6 Then
			EW_UnFormatDateTime = arDate(2) & "/" & arDate(0) & "/" & arDate(1)
		ElseIf ANamedFormat = 7 Then
			EW_UnFormatDateTime = arDate(2) & "/" & arDate(1) & "/" & arDate(0)
		Else ' ANamedFormat = 5 or other
			EW_UnFormatDateTime = arDateTime(0)
		End If
		If UBound(arDateTime) > 0 Then
			If IsDate(arDateTime(1)) Then ' Is time
				EW_UnFormatDateTime = EW_UnFormatDateTime & " " & arDateTime(1)
			End If
		End If
	Else
		EW_UnFormatDateTime = ADate
	End If
End Function

'-------------------------------------------------------------------------------
' Function for debug
Sub Trace(aMsg)
	On Error Resume Next

	Dim fso, ts

	Set fso = Server.Createobject("Scripting.FileSystemObject")
	Set ts = fso.OpenTextFile(Server.MapPath("debug.txt"), 8, True)
	ts.writeline(aMsg)
	ts.Close
	Set ts = Nothing
	Set fso = Nothing
End Sub
%>
<%
'-------------------------------------------------------------------------------
' Functions for file upload

Function stringToByte(toConv)

	Dim i, tempChar

	 For i = 1 to Len(toConv)
	 	tempChar = Mid(toConv, i, 1)
		stringToByte = stringToByte & chrB(AscB(tempChar))
	 Next
	 
End Function

Function byteToString(toConv)

	Dim i, byteord, nextbyteord

	For i = 1 to LenB(toConv)
		byteord = AscB(MidB(toConv, i, 1))
		If byteord < &H80 Then ' Ascii
			byteToString = byteToString & Chr(byteord)
		Else ' Double-byte characters?
			If i < LenB(toConv) Then
				nextbyteord = AscB(MidB(toConv, i+1, 1))
				On Error Resume Next
				' Note: This line does NOT work on all systems due to limitation of the
				' Chr() function
	      byteToString = byteToString & Chr(CInt(byteord) * &H100 + CInt(nextbyteord))
				If Err.Number <> 0 Then
					On Error GoTo 0
					byteToString = byteToString & Chr(byteord) & Chr(nextbyteord)
				End If
				i = i + 1
			ElseIf i = LenB(toConv) Then
				byteToString = byteToString & Chr(byteord)
			End If
		End If
	Next
End Function

Function ConvertToBinary(ByRef rawData)

	Dim oRs

	Set oRs = Server.CreateObject("ADODB.Recordset")
		
	' Create field in an empty RecordSet
	Call oRs.Fields.Append("Blob", 205, LenB(rawData)) ' Add field with type adLongVarBinary
	Call oRs.Open()
	Call oRs.AddNew()
	Call oRs.Fields("Blob").AppendChunk(rawData & ChrB(0))
	Call oRs.Update()
		
	' Save Blob Data
	ConvertToBinary = oRs.Fields("Blob").GetChunk(LenB(rawData))
		
	' Close RecordSet
	Call oRs.Close()
	Set oRs = Nothing
		
End Function

Function ConvertToUnicode(ByRef rawData)

	Dim oRs
		
	Set oRs = Server.CreateObject("ADODB.Recordset")
		
	' Create field in an empty recordset
	Call oRs.Fields.Append("Text", 201, LenB(rawData)) ' Add field with type adLongVarChar
	Call oRs.Open()
	Call oRs.AddNew()
	Call oRs.Fields("Text").AppendChunk(rawData & ChrB(0))
	Call oRs.Update()
		
	' Save Unicode Data
	ConvertToUnicode = oRs.Fields("Text").Value
		
	' Close recordset
	Call oRs.Close()
	Set oRs = Nothing
		
End Function

Function ewUploadPath(parm)

	If parm = 0 Then
		ewUploadPath = ""
	Else
		ewUploadPath = Server.MapPath("/")
	End If

	' Customize the upload path here
	' Check the last delimiter
	If parm = 0 Then
		If Right(ewUploadPath, 1) <> "/" Then ewUploadPath = ewUploadPath & "/"
	Else
		If Right(ewUploadPath, 1) <> "\" Then ewUploadPath = ewUploadPath & "\"
	End If
End Function 

Function ewUploadFileName(sFileName)

	Dim sOutFileName

	' Amend your logic here
	sOutFileName = sFileName

	' Return computed output file name
	ewUploadFileName = sOutFileName
End Function

Function getValue(dict, name)

	Dim gv

	If dict.Exists(name) Then
		gv = CStr(dict(name).Item("Value"))	
		gv = Left(gv,Len(gv)-2)
		getValue = gv
	Else
		getValue = ""
	End If
End Function

Function getFileData(dict, name)
	If dict.Exists(name) Then
		getFileData = dict(name).Item("Value")
		If LenB(getFileData) Mod 2 = 1 Then
			getFileData = getfileData & ChrB(0)
		End If
	Else
		getFileData = ""
	End If
End Function

Function getFileName(dict, name)

	Dim temp, tempPos

	If dict.Exists(name) Then
		temp = dict(name).Item("FileName")
		tempPos = 1 + InStrRev(temp, "\")
		getFileName = Mid(temp, tempPos)
	Else
		getFileName = ""
	End If
End Function

Function getFileSize(dict, name)
	If dict.Exists(name) Then
		getFileSize = LenB(dict(name).Item("Value"))
	Else
		getFileSize = 0
	End If
End Function

Function getFileContentType(dict, name)
	If dict.Exists(name) Then
		getFileContentType = dict(name).Item("ContentType")
	Else
		getFileContentType = ""
	End If
End Function

%>
<%
' Function to Adjust SQL
Function AdjustSql(str)

	Dim sWrk

	sWrk = Trim(str&"")
	sWrk = Replace(sWrk, "'", "''") ' Adjust for Single Quote

	sWrk = Replace(sWrk, "[", "[[]") ' Adjust for Open Square Bracket

	AdjustSql = sWrk

End Function
%>
<%
' Function to Load Email Content from input file name
' - Content Loaded to the following variables
' - Subject: sEmailSubject
' - From: sEmailFrom
' - To: sEmailTo
' - Cc: sEmailCc
' - Bcc: sEmailBcc
' - Format: sEmailFormat
' - Content: sEmailContent
'
Sub LoadEmail(fn)

	Dim sWrk, sHeader, arrHeader
	Dim sName, sValue
	Dim i, j
	sWrk = LoadTxt(fn) ' Load text file content
	If sWrk <> "" Then
		' Locate Header & Mail Content
		i = InStr(sWrk, vbCrLf&vbCrLf)
		If i > 0 Then
			sHeader = Mid(sWrk, 1, i)
			sEmailContent = Mid(sWrk, i+4)
			arrHeader = Split(sHeader, vbCrLf)
			For j = 0 to UBound(arrHeader)
				i = InStr(arrHeader(j), ":")
				If i > 0 Then
					sName = Trim(Mid(arrHeader(j), 1, i-1))
					sValue = Trim(Mid(arrHeader(j), i+1))
					Select Case LCase(sName)
						Case "subject": sEmailSubject = sValue
						Case "from": sEmailFrom = sValue
						Case "to": sEmailTo = sValue
						Case "cc": sEmailCc = sValue
						Case "bcc": sEmailBcc = sValue
						Case "format": sEmailFormat = sValue
					End Select
				End If
			Next 
		End If
	End If

End Sub

' Function to Load a Text File
Function LoadTxt(fn)

	Dim fso, fobj

	' Get text file content
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	Set fobj = fso.OpenTextFile(Server.MapPath(fn))

	LoadTxt = fobj.ReadAll ' Read all Content

	fobj.Close
	Set fobj = Nothing

End Function

' Function to Send out Email
Sub Send_Email(sFrEmail, sToEmail, sCcEmail, sBccEmail, sSubject, sMail, sFormat)

	Dim objMail, objConfig, sServerVersion, i, sIISVer
	Dim sSmtpServer, iSmtpServerPort

	sServerVersion = Request.ServerVariables("SERVER_SOFTWARE")
	If InStr(sServerVersion, "Microsoft-IIS") > 0 Then
		i = InStr(sServerVersion, "/")
		If i > 0 Then
			sIISVer = Trim(Mid(sServerVersion, i+1))
		End If
	End If

	If sIISVer <= "5.0" Then
		' NT / 2000 using CDONTS
		Set objMail = Server.CreateObject("CDONTS.NewMail")
		objMail.From = sFrEmail
		objMail.To = sToEmail
		If sCcEmail <> "" Then
			objMail.Cc = sCcEmail
		End If
		If sBccEmail <> "" Then
			objMail.Bcc = sBccEmail
		End If
		If LCase(sFormat) = "html" Then
			objMail.BodyFormat = 0  ' 0 means HTML format, 1 means text
			objMail.MailFormat = 0  ' 0 means MIME, 1 means text
		End If
		objMail.Subject = sSubject
		objMail.Body = sMail
		objMail.Send
		Set objMail = Nothing
	Else
		' XP / 2003 using CDO
		' Set up Mail
		Set objMail = Server.CreateObject("CDO.Message")
		sSmtpServer = "localhost"
		iSmtpServerPort = 25
		If (sIISVer < "6.0") Or (sSmtpServer <> "" And LCase(sSmtpServer) <> "localhost") Then ' XP or not localhost
			' Set up Configuration
			Set objConfig = CreateObject("CDO.Configuration")
			objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' cdoSendUsingMethod = cdoSendUsingPort
			objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver")  = sSmtpServer ' cdoSMTPServer
			objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = iSmtpServerPort ' cdoSMTPServerPort
			objConfig.Fields.Update
			Set objMail.Configuration = objConfig ' Use Configuration
		End If
		objMail.From = sFrEmail
		objMail.To = sToEmail
		If sCcEmail <> "" Then
			objMail.Cc = sCcEmail
		End If
		If sBccEmail <> "" Then
			objMail.Bcc = sBccEmail
		End If
		If LCase(sFormat) = "html" Then
			objMail.HtmlBody = sMail
		Else
			objMail.TextBody = sMail
		End If
		objMail.Subject = sSubject
		objMail.Send
		Set objMail = Nothing
		Set objConfig = Nothing
	End If

End Sub
%>
<%
' Function to generate Value Separator based on current row count
' rowcnt - zero based row count
'
Function ValueSeparator(rowcnt)

	ValueSeparator = ", "

End Function

' Function to generate View Option Separator based on current row count (Multi-Select / CheckBox)
' rowcnt - zero based row count
'
Function ViewOptionSeparator(rowcnt)

	ViewOptionSeparator = ", "
	' Sample code to adjust 2 options per row
	'If ((rowcnt + 1) Mod 2 = 0) Then ' 2 options per row
		'ViewOptionSeparator = ViewOptionSeparator & "<br>"
	'End If

End Function

' Function to generate Edit Option Separator based on current row count (Radio / CheckBox)
' rowcnt - zero based row count
'
Function EditOptionSeparator(rowcnt)

	EditOptionSeparator = "&nbsp;"
	' Sample code to adjust 2 options per row
	'If ((rowcnt + 1) Mod 2 = 0) Then ' 2 options per row
		'EditOptionSeparator = EditOptionSeparator & "<br>"
	'End If

End Function

' Function to truncate Memo Field based on specified length, string truncated to nearest space or CrLf
'
Function TruncateMemo(str, ln)

	Dim i, j, k

	If Len(str) > 0 And Len(str) > ln Then
		k = 1
		Do While k > 0 And k < Len(str)
			i = InStr(k, str, " ", 1)
			j = InStr(k, str, vbCrLf, 1)
			If i < 0 And j < 0 Then ' Not able to truncate
				TruncateMemo = str
				Exit Function
			Else
				' Get nearest space or CrLf
				If i > 0 And j > 0 Then
					If i < j Then
						k = i
					Else
						k = j
					End If
				ElseIf i > 0 Then
					k = i
				ElseIf j > 0 Then
					k = j
				End If
				' Get truncated text
				If k >= ln Then
					TruncateMemo = Mid(str, 1, k-1) & " ..."
					Exit Function
				Else
					k = k + 1
				End If
			End If
		Loop
	Else
		TruncateMemo = str
	End If

End Function
%>
