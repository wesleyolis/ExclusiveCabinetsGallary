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
'sSql = "SELECT Ranges.Range, Descriptions.Description, Descriptions.Name, Ranges.Name FROM Descriptions, Ranges;"

sSql = "SELECT Ranges.Range, Descriptions.Description FROM Images INNER JOIN Descriptions ON Images.Description = Descriptions.Description, Images AS Images_1 INNER JOIN Ranges ON Images_1.Range = Ranges.Range GROUP BY Ranges.Range, Descriptions.Description ORDER BY Ranges.Range, Descriptions.Description;"

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
<tr><td class="header2"><b><font size="5">Combinations</font></b></td></tr>
<tr height="15px"><td></td></tr>
<tr align="center"><td><a href="SimpleBrowser.asp?Type.html"></a></td></tr>
<%
Do While Not rs.eof
%>
<tr align="center"><td><a href="SimpleBrowser.asp?Range=<%=rs("Range")%>&Des=<%=rs("Description")%>&Type.html"></a></td></tr><%
rs.MoveNext
Loop
rs.Close()
%>


<%
sSql = "SELECT Ranges.Range, Ranges.Name FROM Ranges  INNER JOIN Images ON Ranges.Range = Images.Range GROUP BY Ranges.Range, Ranges.Name ORDER BY Ranges.Name"
Set rs = conn.Execute(sSql)

Do While Not rs.eof
%>
<tr align="center"><td><a href="SimpleBrowser.asp?Range=<%=rs("Range")%>&Type.html"></a></td></tr><%
rs.MoveNext
Loop
rs.Close()
%>

<%
sSql = "SELECT Descriptions.Description, Descriptions.Name FROM Descriptions   INNER JOIN Images ON Descriptions.Description = Images.Description GROUP BY Descriptions.Description, Descriptions.Name Order by Descriptions.Name"
Set rs = conn.Execute(sSql)
Do While Not rs.eof
%>
<tr align="center"><td><a href="SimpleBrowser.asp?Des=<%=rs("Description")%>&Type.html"></a></td></tr><%
rs.MoveNext
Loop
rs.Close()
%>




<%
sSql = "SELECT Image_Groups.Image, Image_Groups.Group FROM Image_Groups ORDER BY Image_Groups.Image, Image_Groups.Group;"
Set rs = conn.Execute(sSql)
Img = ""

While Not rs.eof

	Img = rs("Image")
	GRPS = "Grp=" & rs("Group")
	OK = 1
	If Not rs.eof Then
	rs.Movenext
	End IF

	While Not rs.eof And OK = 1

		If Img = rs("Image") Then
		%><%
			GRPS = GRPS & "&Grp=" & rs("Group")
			rs.MoveNext
		Else
		
			OK = 0
		End IF

	Wend 

%>

<tr align="center"><td><a href="Color_Ranges/Available_Colors.asp?<%=GRPS%>&Type.html"></a></td></tr><%
'rs.MoveNext
Wend
rs.Close()
%>

<tr align="center"><td><a href="Color_Ranges/Available_Colors.asp?&Type.html"></a></td></tr>
</table>




	




</div>
<p><font face="Times New Roman"></font></p>
</body>
<%

' Close recordset and connection

Set rs = Nothing
conn.Close
Set conn = Nothing
%>