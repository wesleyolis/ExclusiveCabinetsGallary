<%@ LANGUAGE = VBScript %>
<!--METADATA TYPE="typelib" 
uuid="00000206-0000-0010-8000-00AA006D2EA4" -->
<% Session.Timeout = 20 %>
<%
Response.ContentType = "text/html; charset=windows-1252"
Response.expires = -1
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "private"
Response.CacheControl = "no-cache"
%>
<!--#include file="db.asp"-->

<%
Response.Buffer = true
Dim dir,udir,str,email,admin, rs,rsi,FilRange, FIL,Group1,Group2, Range,Description, Page,Rows,RCount,Rangetxt,Destxt

Rangetxt =""
Destxt =""
Page = 1
Group1 =""
Group2 =""
Range = ""
Description = ""
FIL=""
FilRange = -1
str=""


Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str

'Response.write(FIL)
sql = "SELECT Images.Image FROM Images;"

Set rsi = conn.Execute(sql)
%>

<html>
<head>
<title>Exclusive Cabinets Gallary</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
</head>
<link rel="stylesheet" href="gallery.css" type="text/css">

<body background="images/creamGradient.jpg" style="background-repeat:repeat-x; background-attachment:fixed">

<%Do While Not rsi.Eof%>
<img  border="0" src="images/thumbs/image.asp?I=<%=rsi.Fields("Image")%>&Type.jpg">
<img border="0" src="images/large/image.asp?I=<%=rsi.Fields("Image")%>&Type.jpg">
<a href="images/html/image.asp?I=<%=rsi.Fields("Image")%>&Type.html"></a>
<%
rsi.MoveNext
Loop
rsi.Close
%>


<%
sql = "SELECT Images.Image, Sub_Image.ID FROM Images INNER JOIN Sub_Image ON Images.Image = Sub_Image.Image;"

Set rsi = conn.Execute(sql)

Do While Not rsi.Eof%>
<img src="images/large/image.asp?I=<%=rsi.Fields("Image")%>_<%=rsi.Fields("ID")%>&Type.jpg">
<%
rsi.MoveNext
Loop
rsi.Close
%><%
conn.Close ' Close Connection
Set conn = Nothing

%>