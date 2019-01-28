<%@ LANGUAGE = VBScript %>

<html>

<head>
<meta http-equiv="Content-Language" content="en-za">
<title>Exclusive Cabinets Gallary - All Available Board Colours</title>
<%
Response.Expires = 0
Response.ExpiresAbsolute = #1/1/1980# ' Expired
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private, no-cache, no-store, must-revalidate"
%>
<!--#include file="db.asp"-->
<%
Dim rs,coloms,colom

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	path = Server.MapPath("thumbs")

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str

Set rs = conn.Execute("SELECT Colors.Color, Colors.Name FROM [Colors] ORDER BY Colors.Name;")

Coloms = 4
%>
</head>

<body background="images/creamGradient.jpg" style="background-repeat:repeat-x; background-attachment:fixed">
<div align="center">
<b><font size="5">All Our Available Board Colours</font></b><p>Please not 
that not all products are offered in every colour.<font size="5"><br></font>We 
ask you to look at the products &quot;Available Colours&quot; link for further information 
and availability<font size="5">.</font></p>
<Table cellspacing="0" cellpadding="20">
<%While Not rs.eof
Colom = 1
%>
<tr>
<%While (colom <= coloms) And (Not rs.eof)
%>
<td>

<table border="2" bordercolor="#800000" style="border-collapse: collapse" cellspacing="0" cellpadding="0" bgcolor="#EEDDD7">
<tr><td width="190" height="200" align="center"><%
	 If (objFSO.FileExists(path & "/Color_" & rs("Color") & ".jpg"))=True Then
		%><a name="jump<%=rs("Color")%>" href="getimage.asp?I=Color_<%=rs("Color")%>&Resample=True"><img border="0" src="thumbs/Color_<%=rs("Color")%>.jpg"></a><%
	Else  %><b>No Photo Available</b><%
		End IF
%></td></tr>
<tr><td align="center" style="padding-top: 3px; padding-bottom: 3px"><%=rs("Name")%> </td></tr>
</table>

</td>
<%
rs.MoveNext
Colom = colom + 1
Wend
If colom < (coloms +1) Then%>
<td colspan="<%=(coloms-colom)%>"></td>
<%End IF%>
</tr>

<%
Wend
rs.Close
rs = NULL

%>
</Table>
</div>
</body>
</html>