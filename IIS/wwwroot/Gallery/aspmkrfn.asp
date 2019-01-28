<%
' ASPMaker functions for ASPMaker 5+
' (C)2006 e.World Technology Ltd.

' Common constants
Const EW_DATE_SEPARATOR = "/"
Const EW_SMTPSERVER = "localhost"
Const EW_SMTPSERVER_PORT = 25
Const EW_SMTPSERVER_USERNAME = ""
Const EW_SMTPSERVER_PASSWORD = ""


'-------------------------------------------------------------------------------
' Functions for default date format
' ANamedFormat = 0-8, where 0-4 same as VBScript
' 5 = "yyyy/mm/dd"
' 6 = "mm/dd/yyyy"
' 7 = "dd/mm/yyyy"
' 8 = Short Date & " " & Short Time

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
		ElseIf ANamedFormat = 8 Then
			EW_FormatDateTime = FormatDateTime(ADate, 2)
			If Hour(ADate) <> 0 Or Minute(ADate) <> 0 Or Second(ADate) <> 0 Then
				EW_FormatDateTime = EW_FormatDateTime & " " & FormatDateTime(ADate, 4) & ":" & ewZeroPad(Second(ADate), 2)
			End If
		Else
			EW_FormatDateTime = ADate
		End If
	Else
		EW_FormatDateTime = ADate
  End If
End Function

Function EW_UnFormatDateTime(ADate, ANamedFormat)
	Dim arDateTime, arDate
	ADate = Trim(ADate & "")
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
			EW_UnFormatDateTime = arDate(0) & "/" & arDate(1) & "/" & arDate(2)
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
' Function for format percent

Function EW_FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
	On Error Resume Next
	EW_FormatPercent = FormatPercent(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
	If Err.Number <> 0 Then
		EW_FormatPercent = FormatNumber(Expression*100, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits) & "%"
	End If
End Function

' Note: Object "conn" is required
Function ewExecuteScalar(SQL)
	ewExecuteScalar = Null
	If Trim(SQL&"") = "" Then	Exit Function
	Dim rs
	Set rs = conn.Execute(SQL)
	If Not rs.Eof Then ewExecuteScalar = rs(0)
	rs.Close
	Set rs = Nothing
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

sEmailFrom = "": sEmailTo = "": sEmailCc = "": sEmailBcc = "": sEmailSubject = "": sEmailFormat = "": sEmailContent = ""

Sub LoadEmail(fn)

	Dim sWrk, sHeader, arrHeader
	Dim sName, sValue
	Dim i, j

	' Initialize
	sEmailFrom = "": sEmailTo = "": sEmailCc = "": sEmailBcc = "": sEmailSubject = "": sEmailFormat = "": sEmailContent = ""

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

	If sIISVer < "5.0" Then
		' NT using CDONTS
		Set objMail = Server.CreateObject("CDONTS.NewMail")
		objMail.From = sFrEmail
		objMail.To = Replace(sToEmail, ",", ";")
		If sCcEmail <> "" Then
			objMail.Cc = Replace(sCcEmail, ",", ";")
		End If
		If sBccEmail <> "" Then
			objMail.Bcc = Replace(sBccEmail, ",", ";")
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
		' 2000 / XP / 2003 using CDO
		' Set up Mail
		Set objMail = Server.CreateObject("CDO.Message")
		sSmtpServer = EW_SMTPSERVER
		iSmtpServerPort = EW_SMTPSERVER_PORT
		If (sIISVer < "6.0") Or (sSmtpServer <> "" And LCase(sSmtpServer) <> "localhost") Then ' XP or not localhost
			' Set up Configuration
			Set objConfig = CreateObject("CDO.Configuration")
			objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 ' cdoSendUsingMethod = cdoSendUsingPort
			objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver")  = sSmtpServer ' cdoSMTPServer
			objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = iSmtpServerPort ' cdoSMTPServerPort
			If EW_SMTPSERVER_USERNAME <> "" And EW_SMTPSERVER_PASSWORD <> "" Then
				objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'cdoBasic (clear text)
				objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername") = EW_SMTPSERVER_USERNAME
				objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword") = EW_SMTPSERVER_PASSWORD
			End If
			objConfig.Fields.Update
			Set objMail.Configuration = objConfig ' Use Configuration
		End If
		objMail.From = sFrEmail
		objMail.To = Replace(sToEmail, ",", ";")
		If sCcEmail <> "" Then
			objMail.Cc = Replace(sCcEmail, ",", ";")
		End If
		If sBccEmail <> "" Then
			objMail.Bcc = Replace(sBccEmail, ",", ";")
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

' Function to render repeat column table
' rowcnt - zero based row count
'
Function RenderControl(totcnt, rowcnt, repeatcnt, rendertype)

	Dim sWrk
	sWrk = ""

	' Render control start
	If rendertype = 1 Then

		If rowcnt = 0 Then sWrk = sWrk & "<table class=""aspmakerlist"">"
		If (rowcnt mod repeatcnt = 0) Then sWrk = sWrk & "<tr>"
		sWrk = sWrk & "<td>"

	' Render control end
	ElseIf rendertype = 2 Then

		sWrk = sWrk & "</td>"
		If (rowcnt mod repeatcnt = repeatcnt -1) Then
			sWrk = sWrk & "</tr>"
		ElseIf rowcnt = totcnt Then
			For i = ((rowcnt mod repeatcnt) + 1) to repeatcnt - 1
				sWrk = sWrk & "<td>&nbsp;</td>"
			Next
			sWrk = sWrk & "</tr>"
		End If
		If rowcnt = totcnt Then sWrk = sWrk & "</table>"

	End If

	RenderControl = sWrk

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
					TruncateMemo = Mid(str, 1, k-1) & "..."
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

<%
'-------------------------------------------------------------------------------
' Function for Writing audit trail
'
Sub ewWriteAuditTrail(pfx, curDate, curTime, id, user, action, table, field, keyvalue, oldvalue, newvalue)
	On Error Resume Next
	Dim fso, ts, sMsg, sFn, sFolder
	Dim bWriteHeader, sHeader
	Dim userwrk
	userwrk = user
	If userwrk = "" Then userwrk = "-1" ' assume Administrator if no user
	sHeader = "date" & vbTab & _
		"time" & vbTab & _
		"id" & vbTab & _
		"user" & vbTab & _
		"action" & vbTab & _
		"table" & vbTab & _
		"field" & vbTab & _
		"key value" & vbTab & _
		"old value" & vbTab & _
		"new value"
	sMsg = curDate & vbTab & _
		curTime & vbTab & _
		id & vbTab & _
		userwrk & vbTab & _
		action & vbTab & _
		table & vbTab & _
		field & vbTab & _
		keyvalue & vbTab & _
		oldvalue & vbTab & _
		newvalue
	sFolder = ""
	sFn = pfx & "_" & ewZeroPad(Year(Date), 4) & ewZeroPad(Month(Date), 2) & ewZeroPad(Day(Date), 2) & ".txt"
	Set fso = Server.Createobject("Scripting.FileSystemObject")
	bWriteHeader = Not fso.FileExists(Server.MapPath(sFolder & sFn))
	Set ts = fso.OpenTextFile(Server.MapPath(sFolder & sFn), 8, True)
	If bWriteHeader Then
		ts.writeline(sHeader)
	End If
	ts.writeline(sMsg)
	ts.Close
	Set ts = Nothing
	Set fso = Nothing
End Sub

' Pad zeros before number
Function ewZeroPad(m, t)
  ewZeroPad = String(t - Len(m), "0") & m
End Function

' IIf function
Function ewIIf(cond, v1, v2)
	On Error Resume Next
	If CBool(cond) Then
		ewIIf = v1
	Else
		ewIIf = v2
	End If
End Function

' Convert different data type value
Function ewConv(v, t)

	Select Case t

	' adBigInt/adUnsignedBigInt
	Case 20, 21
		If IsNull(v) Then
			ewConv = Null
		Else
			ewConv = CLng(v)
		End If

	' adSmallInt/adInteger/adTinyInt/adUnsignedTinyInt/adUnsignedSmallInt/adUnsignedInt/adBinary
	Case 2, 3, 16, 17, 18, 19, 128
		If IsNull(v) Then
			ewConv = Null
		Else
			ewConv = CInt(v)
		End If

	' adSingle
	Case 4
		If IsNull(v) Then
			ewConv = Null
		Else
			ewConv = CSng(v)
		End If

	' adDouble/adCurrency/adNumeric
	Case 5, 6, 131
		If IsNull(v) Then
			ewConv = Null
		Else
			ewConv = CDbl(v)
		End If

	Case Else
		ewConv = v

	End Select

End Function
%>
<script language="JScript" runat="server">
// Server-side JScript functions for ASPMaker 5+ (Requires script engine 5.5.+)

function EW_Encode(str) {	
	return encodeURIComponent(str);
}

function EW_Decode(str) {	
	return decodeURIComponent(str);	
}

// encrytion key
EW_RANDOM_KEY = 'WMG1%&Er99UCJKr5';

// JavaScript implementation of Block TEA by Chris Veness
// http://www.movable-type.co.uk/scripts/TEAblock.html
//
// TEAencrypt: Use Corrected Block TEA to encrypt plaintext using password
//            (note plaintext & password must be strings not string objects)
//
// Return encrypted text as string
//
function TEAencrypt(plaintext, password)
{
    if (plaintext.length == 0) return('');  // nothing to encrypt
    // 'escape' plaintext so chars outside ISO-8859-1 work in single-byte packing, but  
    // keep spaces as spaces (not '%20') so encrypted text doesn't grow too long, and 
    // convert result to longs
    var v = strToLongs(escape(plaintext).replace(/%20/g,' '));
    if (v.length == 1) v[1] = 0;  // algorithm doesn't work for n<2 so fudge by adding nulls
    var k = strToLongs(password.slice(0,16));  // simply convert first 16 chars of password as key
    var n = v.length;

    var z = v[n-1], y = v[0], delta = 0x9E3779B9;
    var mx, e, q = Math.floor(6 + 52/n), sum = 0;

    while (q-- > 0) {  // 6 + 52/n operations gives between 6 & 32 mixes on each word
        sum += delta;
        e = sum>>>2 & 3;
        for (var p = 0; p < n-1; p++) {
            y = v[p+1];
            mx = (z>>>5 ^ y<<2) + (y>>>3 ^ z<<4) ^ (sum^y) + (k[p&3 ^ e] ^ z)
            z = v[p] += mx;
        }
        y = v[0];
        mx = (z>>>5 ^ y<<2) + (y>>>3 ^ z<<4) ^ (sum^y) + (k[p&3 ^ e] ^ z)
        z = v[n-1] += mx;
    }
    // note use of >>> in place of >> due to lack of 'unsigned' type in JavaScript 

    return escCtrlCh(longsToStr(v));
}

//
// TEAdecrypt: Use Corrected Block TEA to decrypt ciphertext using password
//
function TEAdecrypt(ciphertext, password)
{
    if (ciphertext.length == 0) return('');
    var v = strToLongs(unescCtrlCh(ciphertext));
    var k = strToLongs(password.slice(0,16)); 
    var n = v.length;

    var z = v[n-1], y = v[0], delta = 0x9E3779B9;
    var mx, e, q = Math.floor(6 + 52/n), sum = q*delta;

    while (sum != 0) {
        e = sum>>>2 & 3;
        for (var p = n-1; p > 0; p--) {
            z = v[p-1];
            mx = (z>>>5 ^ y<<2) + (y>>>3 ^ z<<4) ^ (sum^y) + (k[p&3 ^ e] ^ z)
            y = v[p] -= mx;
        }
        z = v[n-1];
        mx = (z>>>5 ^ y<<2) + (y>>>3 ^ z<<4) ^ (sum^y) + (k[p&3 ^ e] ^ z)
        y = v[0] -= mx;
        sum -= delta;
    }

    var plaintext = longsToStr(v);
    // strip trailing null chars resulting from filling 4-char blocks:
    if (plaintext.search(/\0/) != -1) plaintext = plaintext.slice(0, plaintext.search(/\0/));

    return unescape(plaintext);
}


// supporting functions

function strToLongs(s) {  // convert string to array of longs, each containing 4 chars
    // note chars must be within ISO-8859-1 (with Unicode code-point < 256) to fit 4/long
    var l = new Array(Math.ceil(s.length/4))
    for (var i=0; i<l.length; i++) {
        // note little-endian encoding - endianness is irrelevant as long as 
        // it is the same in longsToStr() 
        l[i] = s.charCodeAt(i*4) + (s.charCodeAt(i*4+1)<<8) + 
               (s.charCodeAt(i*4+2)<<16) + (s.charCodeAt(i*4+3)<<24);
    }
    return l;  // note running off the end of the string generates nulls since 
}              // bitwise operators treat NaN as 0

function longsToStr(l) {  // convert array of longs back to string
    var a = new Array(l.length);
    for (var i=0; i<l.length; i++) {
        a[i] = String.fromCharCode(l[i] & 0xFF, l[i]>>>8 & 0xFF, 
                                   l[i]>>>16 & 0xFF, l[i]>>>24 & 0xFF);
    }
    return a.join('');  // use Array.join() rather than repeated string appends for efficiency
}

function escCtrlCh(str) {  // escape control chars which might cause problems with encrypted texts
    return str.replace(/[\0\n\v\f\r!]/g, function(c) { return '!' + c.charCodeAt(0) + '!'; });
}

function unescCtrlCh(str) {  // unescape potentially problematic nulls and control characters
    return str.replace(/!\d\d?!/g, function(c) { return String.fromCharCode(c.slice(1,-1)); });
}

</script>
