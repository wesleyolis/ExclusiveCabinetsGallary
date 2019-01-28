<%@ LANGUAGE = VBScript %>
<!--METADATA TYPE="typelib" 
uuid="00000206-0000-0010-8000-00AA006D2EA4" -->
<% Session.Timeout = 20 %>
<%
Response.expires = -1
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "private"
Response.CacheControl = "no-cache"
%>
<!--#include file="db.asp"-->

<%
Response.Buffer = true


If Session("Admin") Then
Session("Admin") = true 
admin=true
'str = str & "&admin"
Else
Session("Admin") = false
admin=false
End IF

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str



if Trim(Request.QueryString("Ulti")) <> ""  Then
Ulti = Request.QueryString("Ulti")


	Dim rs, SQL

'SQL = "UPDATE Images SET Images.Price = " & Price & " WHERE (((Images.Code)=UCase('" & Ulti & "')))"
SQL = "SELECT * From Images WHERE (((Images.Code)=UCase('" & Ulti & "')))"



	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open SQL, conn, 1, 2
	rs.PageSize = 1
If rs.Eof Then
%>
<input name="suc" value="false">
<%
Else

%>
<input name="suc" value="true">
<input name="count" value="<%=rs.PageCount%>">
<%
End If
'Conn.Execute(SQL)
%>

<%
Else
%><input name="suc" value="false"><%
End IF


conn.Close ' Close Connection
Set conn = Nothing

%>