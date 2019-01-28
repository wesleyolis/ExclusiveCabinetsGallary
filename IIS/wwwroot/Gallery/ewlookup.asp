<%@ EnableSessionState=False %>
<!--#include file="ewconfig.asp"-->
<!--#include file="db.asp"-->
<!--#include file="aspmkrfn.asp"-->
<%
Response.Expires = 0
Response.ExpiresAbsolute = #1/1/1980# ' Expired
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private, no-cache, no-store, must-revalidate"
%>
<%
Dim QS, Sql, Value
QS = Split(Request.Querystring, "&")
If IsArray(QS) Then
	If UBound(QS) >= 0 Then
		Sql = EW_GetValue("s")
		Sql = TEAdecrypt(Sql, EW_RANDOM_KEY)
		Value = EW_GetValue("q")
		Value = AdjustSql(Value)
		If Sql <> "" Then
			If Value <> "" Then Sql = Replace(Sql, "@FILTER_VALUE", Value)
			EW_GetLookupValues(Sql)
		End If
	End If
End If
Function EW_GetValue(Key)
	Dim kv, I
	For I = 0 To UBound(QS)
		kv = Split(QS(I), "=")
		If (kv(0) = Key) Then
			EW_GetValue = EW_Decode(kv(1))
			Exit Function
		End If
	Next
	EW_GetValue = ""
End Function
Sub EW_GetLookupValues(Sql)

	' Connect to database
	Dim conn, rs, rsarr, str, I, J
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open xDb_Conn_Str
	Set rs = conn.Execute(Sql)
	If Not rs.EOF Then
		rsarr = rs.GetRows
	End If

	' Close database
	rs.Close
	Set rs = Nothing
	conn.Close
	Set conn = Nothing

	' Output
	If IsArray(rsarr) Then
		For J = 0 To UBound(rsarr, 2)
			For I = 0 To UBound(rsarr, 1)
		    str = rsarr(I, J)
				If Len(str) > 0 Then
					str = Replace(str, vbCrLf, " ")
					str = Replace(str, vbCr, " ")
					str = Replace(str, vbLf, " ")
				End If 
		    Response.write str & vbCr
		  Next
		Next
	End If
End Sub
%>
