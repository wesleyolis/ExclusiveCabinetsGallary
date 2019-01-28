<%@ EnableSessionState=False %>
<!--#include file="ewconfig.asp"-->
<!--#include file="db.asp"-->
<!--#include file="aspmkrfn.asp"-->
<%
On Error Resume Next
Dim LeftQuote, RightQuote, QS, Sql, Where, FieldList, ValueList
Dim LookupTableName, LinkFieldName, DisplayFieldName, DisplayField2Name
Dim LinkField, DisplayField, DisplayField2, LinkFieldQuote, DisplayFieldQuote, DisplayField2Quote
Dim bError
Dim bUseLinkField, bUseDisplayField, bUseDisplayField2
LeftQuote = "["
RightQuote = "]"
bError = False
QS = Split(Request.Querystring, "&")
If IsArray(QS) Then
	If UBound(QS) >= 0 Then
		LookupTableName = EW_GetValue("ltn")
		LinkFieldName = EW_GetValue("lfn")
		DisplayFieldName = EW_GetValue("dfn")
		DisplayField2Name = EW_GetValue("df2n")
		LinkField = EW_GetValue("lf")
		If DisplayFieldName = LinkFieldName Then
			DisplayField = LinkField
		Else
			DisplayField = EW_GetValue("df")
		End If
		If DisplayField2Name = LinkFieldName Then
			DisplayField2 = LinkField
		ElseIf DisplayField2Name = DisplayFieldName Then
			DisplayField2 = DisplayField
		Else
			DisplayField2 = EW_GetValue("df2")
		End If
		LinkFieldQuote = EW_GetValue("lfq")
		DisplayFieldQuote = EW_GetValue("dfq")
		DisplayField2Quote = EW_GetValue("df2q")
	Else
		Response.Write "Invalid Parameter"
		Response.End
	End If
Else
	Response.Write "Invalid Parameter"
	Response.End
End If
If LookupTableName = "" Then
	Response.Write "Missing lookup table name"
	Response.End
End If
If DisplayFieldName = "" Then
	Response.Write "Missing display field name"
	Response.End
End If
bUseLinkField = (LinkFieldName <> "" And LinkField <> "")
bUseDisplayField = (DisplayFieldName <> "" And DisplayFieldName <> LinkFieldName And DisplayField <> "")
bUseDisplayField2 = (DisplayField2Name <> "" And DisplayField2Name <> LinkFieldName And DisplayField2Name <> DisplayFieldName And DisplayField2 <> "")
Sql = ""
If bUseLinkField Then
	Sql = Sql & LeftQuote & LinkFieldName & RightQuote
End If
If bUseDisplayField Then
	If Sql <> "" Then Sql = Sql & ","
	Sql = Sql & LeftQuote & DisplayFieldName & RightQuote
End If
If bUseDisplayField2 Then
	If Sql <> "" Then Sql = Sql & ","
	Sql = Sql & LeftQuote & DisplayField2Name & RightQuote
End If
Sql = "SELECT DISTINCT " & Sql & " FROM " & LeftQuote & LookupTableName & RightQuote
Where = ""
If bUseLinkField Then
	Where = LeftQuote & LinkFieldName & RightQuote & "=" & LinkFieldQuote & AdjustSql(LinkField) & LinkFieldQuote
End If
If bUseDisplayField Then
	If Where <> "" Then Where = Where & " AND "
	Where = Where & LeftQuote & DisplayFieldName & RightQuote & "=" & DisplayFieldQuote & AdjustSql(DisplayField) & DisplayFieldQuote
End If
If bUseDisplayField2 Then
	If Where <> "" Then Where = Where & " AND "
	Where = Where & LeftQuote & DisplayField2Name & RightQuote & "=" & DisplayField2Quote & AdjustSql(DisplayField2) & DisplayField2Quote
End If
Sql = Sql & " WHERE " & Where
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Set rs = conn.Execute(Sql)
If Err.Number <> 0 Then
	Response.Write Err.Description
	bError = True
End If
If Not bError Then
	If rs.Eof Then ' Add new option
		FieldList = ""
		ValueList = ""
		If bUseLinkField Then
			FieldList = FieldList & LeftQuote & LinkFieldName & RightQuote
			ValueList = ValueList & LinkFieldQuote & AdjustSql(LinkField) & LinkFieldQuote
		End If
		If bUseDisplayField Then
			If FieldList <> "" Then FieldList = FieldList & ","
			FieldList = FieldList & LeftQuote & DisplayFieldName & RightQuote
			If ValueList <> "" Then ValueList = ValueList & ","
			ValueList = ValueList & DisplayFieldQuote & AdjustSql(DisplayField) & DisplayFieldQuote
		End If
		If bUseDisplayField2 Then
			If FieldList <> "" Then FieldList = FieldList & ","
			FieldList = FieldList & LeftQuote & DisplayField2Name & RightQuote
			If ValueList <> "" Then ValueList = ValueList & ","
			ValueList = ValueList & DisplayField2Quote & AdjustSql(DisplayField2) & DisplayField2Quote
		End If
		conn.Execute("INSERT INTO " & LeftQuote & LookupTableName & RightQuote & " (" & FieldList & ") VALUES (" & ValueList & ")")
		If Err.Number <> 0 Then
			Response.Write Err.Description
			bError = True
		End If
	Else
		Response.Write "Option already exists"
		bError = True
	End If
End If
rs.Close
Set rs = Nothing
If Not bError Then
	If LinkField = "" Then ' Get new link field value
		Sql = "SELECT " & LeftQuote & LinkFieldName & RightQuote & " FROM " & LeftQuote & LookupTableName & RightQuote & " WHERE " & Where
		Set rs = conn.Execute(Sql)
		If Not rs.Eof Then
			LinkField = rs(0)
			If DisplayFieldName = LinkFieldName Then DisplayField = LinkField
			If DisplayField2Name = LinkFieldName Then DisplayField2 = LinkField
		End If
		rs.Close
		Set rs = Nothing
	End If
End If
conn.Close
Set conn = Nothing
If Not bError Then
	Response.Clear
	Response.Write "OK" & vbCr
	Response.Write LinkField & vbCr
	Response.Write DisplayField & vbCr
	Response.Write DisplayField2
End If
Response.End
Function EW_GetValue(Key)
	Dim kv
	For I = 0 To UBound(QS)
		kv = Split(QS(I), "=")
		If (kv(0) = Key) Then
			EW_GetValue = EW_Decode(kv(1))
			Exit Function
		End If
	Next
	EW_GetValue = ""
End Function
%>
