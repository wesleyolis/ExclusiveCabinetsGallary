<%@ CodePage = 1252 LCID = 7177 %>
<% Session.Timeout = 20 %>
<%
Response.expires = -1
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "no-cache"
Response.CacheControl = "no-cache"
%>
<!--#include file="db.asp"-->

<% Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str

%>

<html>
<head>
<title>Exclusive Cabinets Gallary</title>
</head>
<link rel="stylesheet" href="gallery.css" type="text/css">
<style>
 .udot {border-bottom-style: dashed; border-bottom-width: 1px; padding: 2px; }
 .t1{font-size:11pt;border: 2px solid #cc907e; padding: 0}
</style>
<body background="images/creamGradient.jpg" style="background-repeat:repeat-x; background-attachment:fixed"  topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">




<table width="100%" height="193px" border="1" cellspacing="0" cellpadding="0">
<tr><td background="/images/email_footer_span.JPG" ><td></tr>
</Table>
