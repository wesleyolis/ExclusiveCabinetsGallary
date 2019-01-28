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
Dim dir,udir,str,email,admin, rs,rsi,FilRange, FIL,Group1,Group2, Range,Description, Page

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



if Request.QueryString("Ulti") <> "" And Request.QueryString("Price") <> "" Then
Ulti = Request.QueryString("Ulti")
Price = Request.QueryString("Price")

SQL = "UPDATE Images SET Images.Price = " & Price & " WHERE (((Images.Code)=UCase('" & Ulti & "')))"

Conn.Execute(SQL)
%>
Found
<%
Else
%>Missing<%
End IF


conn.Close ' Close Connection
Set conn = Nothing

%>