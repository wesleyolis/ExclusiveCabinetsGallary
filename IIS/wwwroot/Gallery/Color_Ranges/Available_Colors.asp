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

<%
Dim rs,coloms,colom,Group

	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
	path = Server.MapPath("../thumbs")


xDb_Conn_Str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../Gallary.mdb") & ";"
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str

'Set rs = conn.Execute("SELECT Image_Groups.Group AS GRP FROM Image_Groups WHERE (((Image_Groups.Image)=78));")

'Set rs = conn.Execute("SELECT [Image_Groups].[Group] AS GRP, [Colors].[Color] AS Color, [Colors].[Name] AS CName FROM (Image_Groups INNER JOIN Colors_Group ON Image_Groups.Group = Colors_Group.Grp) INNER JOIN Colors ON Colors_Group.Colours = Colors.Color WHERE (((Image_Groups.Image)=78)) ORDER BY Image_Groups.Group, Colors.Name ;")



C=1
Sql = "(1=2)"
Cmax = Request.QueryString("Grp").Count
While C <= Cmax
Sql = Sql & " Or ((Colors_Group.Grp)= " & Request.QueryString("Grp")(C) & ") "
C = C + 1
Wend

Set rs = conn.Execute("SELECT Colors_Group.Grp As GRP, Color_Groups.Name AS GName, Color_Groups.Memo AS GMemo, Colors.Color AS Color, Colors.Name AS CName FROM Color_Groups INNER JOIN (Colors_Group INNER JOIN Colors ON Colors_Group.Colours = Colors.Color) ON Color_Groups.Index = Colors_Group.Grp WHERE (" & Sql & ") ORDER BY Color_Groups.Name, Colors.Name, Colors_Group.Grp;")

Coloms = 4

strGP = 0
%>
</head>

<body background="../images/creamGradient.jpg" style="background-repeat:repeat-x; background-attachment:fixed">
<div align="center">
<font size="5"><b>Items Available Colours</b></font><b><font size="5"><br></font></b>
<%If rs.eof Then%>
<br>
<br>
<b><font size="4">Sorry, The available Colours for this Item are currently unavailable</font></b>
<%End If%>

<%
While Not rs.eof
strGP = rs("GRP") 
If Not (Group = strGP) Then
Group = strGP
%>
<br>
<b><font size="4"><%=rs("GName")%> Range</font></b><br><br><%=rs("GMemo")%>
<%End IF%>

<Table cellspacing="0" cellpadding="20">
<%
LP1 = True
While LP1 AND Not rs.eof
strGP = rs("GRP") 
If (Group = strGP) Then
Colom = 1
%>
<tr>
<%
LP2 = True
While LP2 And(colom <= coloms) And Not rs.eof
If Group = rs("GRP") Then
%>
<td>

<table border="2" bordercolor="#800000" style="border-collapse: collapse" cellspacing="0" cellpadding="0" bgcolor="#EEDDD7">
<tr><td width="190px" height="200px" align="center"><%
		If (objFSO.FileExists(path & "/Color_" & rs("Color") & ".jpg"))=True Then
		%><a name="jump<%=rs("Color")%>" href="../images/large/image.asp@I=Color_<%=rs("Color")%>&Resample=True&Type.jpg"><img border="0" src="../thumbs/Color_<%=rs("Color")%>.jpg"></a><%
			Else
			%><b>No Photo Available</b><%	End IF
%></td></tr>
<tr><td align="center" style="padding-top: 3px; padding-bottom: 3px"><%=rs("CName")%> </td></tr>
</table>

</td>
<%
rs.MoveNext
Colom = colom + 1
Else
LP2 = False

End IF
Wend
If colom < (coloms + 1) Then%>
<td colspan="<%=(coloms-colom)%>"></td>
<%End IF%>
</tr>

<%
Else
LP1 = False
End IF
Wend


%>
</Table>
<br>
<%Wend
rs.Close
rs = NULL
%>


</div>
</body>
</html>