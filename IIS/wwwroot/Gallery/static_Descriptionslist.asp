<%@ CodePage = 1252 LCID = 7177 %>
<% Session.Timeout = 20 %>
<%
Response.ContentType = "text/html; charset=windows-1252"
Response.expires = 0
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "private"
Response.CacheControl = "no-cache"
%>
<!--#include file="db.asp"-->
<% 
Dim rs
rs = NULL
' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str


' Build SQL
sSql = "SELECT * FROM [Descriptions] Order By [Name]"

' Set up Record Set
Set rs = conn.Execute(sSql)

%>
<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Exclusive Cabinets Gallary</title>
</head>
<link rel="stylesheet" href="gallery.css" type="text/css">
<body  background="images/creamGradient.jpg" style="background-repeat:repeat-x; background-attachment:fixed" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0">
<div align="center">
<table border="0" cellspacing="0" cellpadding="2">
<tr><td class="header2"><b><font size="5">Please Select Description</font></b></td></tr>
<tr height="15px"><td></td></tr>
<%
Do While Not rs.eof
%>
<tr align="center"><td><a href="SimpleBrowser.asp?Des=<%=rs("Description")%>&Type.html"></a><a href="SimpleBrowser.asp@Des=<%=rs("Description")%>&Type.html"><%=Server.HTMLEncode(rs("Name"))%></a></td></tr><%
rs.MoveNext
Loop

%>

</table>




	<table border="0" width="100%" cellspacing="0" cellpadding="0" height="193px" id="table1">
	<tr>
		<td>
		<img border="0" src="images/email_footer_span.JPG" width="100%" height="193"></td>
		<td width="752">
		<img border="0" src="images/email%20footer.JPG" width="752" height="193"></td>
		<td>
		<img border="0" src="images/email_footer_span.JPG" width="100%" height="193"></td>
	</tr>
	</table>




</div>
<p><font face="Times New Roman"></font></p>
</body>
<%

' Close recordset and connection
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>