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
Dim rs,coloms,colom,Group, GP

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str

'Set rs = conn.Execute("SELECT Image_Groups.Group AS GRP FROM Image_Groups WHERE (((Image_Groups.Image)=78));")

'Set rs = conn.Execute("SELECT [Image_Groups].[Group] AS GRP, [Colors].[Color] AS Color, [Colors].[Name] AS CName FROM (Image_Groups INNER JOIN Colors_Group ON Image_Groups.Group = Colors_Group.Grp) INNER JOIN Colors ON Colors_Group.Colours = Colors.Color WHERE (((Image_Groups.Image)=78)) ORDER BY Image_Groups.Group, Colors.Name ;")


Set rs = conn.Execute("SELECT Image_Groups.Group AS GRP, Color_Groups.Name AS GName, Color_Groups.Memo AS GMemo, Colors.Color AS Color, Colors.Name AS CName FROM ((Image_Groups INNER JOIN Color_Groups ON Image_Groups.Group = Color_Groups.Index) INNER JOIN Colors_Group ON Image_Groups.Group = Colors_Group.Grp) INNER JOIN Colors ON Colors_Group.Colours = Colors.Color WHERE (((Image_Groups.Image)=78)) ORDER BY Color_Groups.Name, Colors.Name;")



Coloms = 4
Group = ""
GP = ""
%>
</head>

<body background="images/creamGradient.jpg" style="background-repeat:repeat-x; background-attachment:fixed">
<div align="center">
<font size="5"><b>Items Available Colors</b></font><b><font size="5"><br></font></b>
<%If rs.eof Then%>
<br><br>
<b><font size="4">Sorry, The available Colours for this Item are currently unavailable</font></b>
<%End If%>

<%

While Not rs.eof

GP = rs("GRP")

If Not (Group = GP) Then

Group = rs("GRP")
%>
<%=rs("GName")%> And memo <%=rs("GMemo")%>
<%End IF%>

<Table cellspacing="0" cellpadding="20">
<%While Not rs.eof
If Group = rs("GRP") Then
Colom = 1
%>
<tr>
<%While (colom <= coloms) And Not rs.eof
If Group = rs("GRP") Then
%>
<td>

<table border="2" bordercolor="#800000" style="border-collapse: collapse" cellspacing="0" cellpadding="0" bgcolor="#EEDDD7">
<tr><td width="190" height="200" align="center"><a name="jump<%=rs("Color")%>" href="getimage.asp?I=Color_<%=rs("Color")%>&Resample=True"><img border="0" src="thumbs/Color_<%=rs("Color")%>.jpg"></a></td></tr>
<tr><td align="center" style="padding-top: 3px; padding-bottom: 3px"><%=rs("CName")%> </td></tr>
</table>

</td>
<%
rs.MoveNext
Colom = colom + 1
End IF
Wend
If colom < coloms Then%>
<td colspan="<%=(coloms-colom)%>"></td>
<%End IF%>
</tr>

<%
End IF
Wend
rs.Close
rs = NULL

%>
</Table>

<%Wend%>


</div>
</body>
</html>