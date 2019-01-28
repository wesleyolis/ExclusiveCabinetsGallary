<%@ LANGUAGE = VBScript %>

<% Session.Timeout = 20 %>
<%
Response.expires = -1
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "no-cache"
Response.CacheControl = "no-cache"
%>
<!--#include file="db.asp"-->

<%
Response.Buffer = True
Dim rsi,img
Dim Des,Width,Height,Depth,Range,Color,Edge,Code,Price,Info

imags = ""
pos = 1
While pos <= Request.QueryString("I").Count
imgs = imgs & " OR ([Images.Image]=" & Request.QueryString("I")(pos) & ")"
pos = pos + 1
Wend

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str

sql = "SELECT Images.Image, Images.Code, Ranges.Name, Descriptions.Name, Colors.Name, Images.Width, Images.Height, Images.Depth, IIf([Images.Price]=0,'P.O.A','R ' & [Images.Price]) AS Price, Images.Edge, IIf(IsNull([Info]),' ',[Info]) AS Info2, Sub_Image.ID, Sub_Image.Des"_
&" FROM (Descriptions RIGHT JOIN (Colors RIGHT JOIN (Ranges RIGHT JOIN Images ON Ranges.Range = Images.Range) ON Colors.Color = Images.Color) ON Descriptions.Description = Images.Description) LEFT JOIN Sub_Image ON Images.Image = Sub_Image.Image"_
&" WHERE ((([Images.Image])=-1)" & imgs & ");"



set rsi = Server.CreateObject("ADODB.recordset")
rsi.Open sql, conn
%>

<html>
<head>
<title>Exclusive Cabinets Gallary</title>
</head>
<style>
 .t1{font-size:11pt;border: 2px solid #cc907e; padding: 0; color:#333E64}
  .t2{font-size:11pt; padding: 0; color:#333E64}
  .udot {border-bottom-style: dashed; border-bottom-width: 1px; padding: 2px; }
 .ldot {border-bottom-width: 1px; border-left-style:dashed; border-left-width:1px; border-right-width:1px; border-top-width:1px; padding-left:5px; padding-right:2px; padding-top:2px; padding-bottom:2px }

 .body{ background-repeat:repeat-x; background-attachment:fixed}

.box {
	border-right: 1px solid #000000;
	border-left: 1px solid #000000;
}
.box2 {
	border:1px; solid #CC907E; height:29; background-image:url('images/bar_back.gif');
}
BODY {
scrollbar-face-color:#3a4775;
scrollbar-highlight-color:#3a4775;
scrollbar-3dlight-color:#ffffff;
scrollbar-darkshadow-color:#3a4775;
scrollbar-shadow-color:#ffffff;
scrollbar-arrow-color:#ffffff;
scrollbar-track-color:#3a4775;
}

a {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-weight: bold;
	color: #3c4976;
	text-decoration: none;
}
a:hover {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-weight: bold;
	color: #de4010;
	text-decoration: underline;
}
a:active {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 12px;
	font-weight: bold;
	color: #de4010;
	text-decoration: underline;
}


</style>
<body background="images/creamGradient.jpg" style="background-repeat:repeat-x; background-attachment:fixed" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">
<div align="center">
<b><font size="4"><br>Products as found in E-mail<br><br>Please note information 
below doesn't apply to your quotation.<br></font></b>

<table border="0" cellpadding="0" cellspacing="0">
<tr><td>
<%
Rows = 1
RCount = 0
%>
<table border="0" cellpadding="0" cellspacing="20"><%

Do While Not rsi.eof
img = rsi.Fields("Image")
Des = rsi.Fields("Descriptions.Name")
Width = rsi.Fields("Width")
Height = rsi.Fields("Height")
Depth = rsi.Fields("Depth")
Range = rsi.Fields("Ranges.Name")
Color = rsi.Fields("Colors.Name")
Edge = rsi.Fields("Edge")
Code = rsi.Fields("Code")
Price = rsi.Fields("Price")
Info = replace(rsi.Fields("Info2"),"<br>","  ")
If Trim(Info) = "" Then
Info = "No additional information available"
End IF

If Rcount = 0 Then%><tr>
<% End IF%>
<td valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="640" bgcolor="#EEDDD7" class="t1">
<tr><td colspan="2" height="188" class="udot" align="center">
<a href="images/html/image.asp?I=<%=rsi.Fields("Image")%>"><img border="1" src="images/thumbs/image.asp?I=<%=rsi.Fields("Image")%>"></a></td>
<td rowspan="8" align="center" valign="top" width="52%" class="ldot">

<table border="0" width="311" cellspacing="0" cellpadding="0" class="t2">
<tr><td height="0" width="55%"></td><td></td></tr><%

MR = 0
MC = 0
break = True

Do While (Not rsi.eof) AND (MC <=2) And (break)
IF (Not IsNull(rsi.Fields("ID"))) And (img = rsi.Fields("Image")) Then
%>
<%If MC = 0 Then%><tr>
<% End IF%>
<td>
<Table  border="0" width="151" height="170" class="t2">
<tr><td  height="117" width="151" align="center"><img border="1" src="images/thumbs/image.asp?I=<%=img%>_<%=rsi.Fields("ID")%>"></td></tr>
<tr><td align="center"  valign="top"><%=rsi.Fields("Des")%><br></td></tr>
</Table>
</td>
<%
MC = MC + 1
If MC >= 2 Then
MC = 0
MR = MR +1
%>
</tr>
<%End IF
img = rsi.Fields("Image")
rsi.MoveNext
Else
break = False
End IF
Loop

If MR = 0 And MC = 0 Then%>
<tr><td align="center" valign="top" height="150px" colspan="2"><br>No other Images found of this element</td></tr>
<tr><td colspan="2"></td></tr><%
End IF
%>
</table>

</td></tr><tr>
<td colspan="2"><b>&nbsp;Description:</b> <%=des%></td></tr>
<tr><td colspan="2"><b>&nbsp;Dimensions: </b><%=Width%> <b>X</b> <%=Height%> <b>X</b> 
<%=Depth%><font size="2"> in mm</font></td></tr>
<tr><td colspan="2"><b>&nbsp;Range:</b>&nbsp; <%=Range%></td></tr>
<tr><td colspan="2"><b>&nbsp;Colour:</b> <%=Color%></td></tr>
<tr><td width="194"><b>&nbsp;Code:</b>&nbsp;&nbsp;&nbsp; <%=Code%>&nbsp; </td>
<td height="22" width="122"><b>G.C:</b> <%=Image%></td></tr>
<%If MR =< 1  Then%>
<tr><td colspan="2" height="1"></td></tr>
<%Else%>

<%End IF%>
</table>
</td>
<%
RCount = RCount + 1
If RCount >= Rows Then
RCount = 0%>
</tr>
<%End IF%>

<%
if Not rsi.eof Then
if img = rsi.Fields("Image") Then
rsi.MoveNext
End IF
End IF
Loop
rsi.Close()
%>
<%If RCount < Rows Then%>
<td colspan="<%=(Rows-Rcount)%>"></td>
</tr>
<%End IF%>
</table>
</td></tr>
</table>
<br>
</div>
</td></tr>
</table>



<%
'Avalible colors script



colors = ""
pos = 1
While pos <= Request.QueryString("C").Count
colors = colors & " OR ([Colors.Color]=" & Request.QueryString("C")(pos) & ")"
pos = pos + 1
Wend


Dim coloms, colom


Set rs = conn.Execute("SELECT Colors.Color, Colors.Name FROM [Colors] WHERE (([Colors.Color]= -1) " & colors & ")  ORDER BY Colors.Name;")

Coloms = 3

If Not rs.eof Then
%>

<div align="center">

<Table cellspacing="0" cellpadding="20">
<tr><td colspan="<%=Coloms%>" align="center"><font size="5">Items are offer in the following colors</font></td></tr>
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
		%><a name="jump<%=rs("Color")%>" href="images/large/image.asp?I=Color_<%=rs("Color")%>&Resample=True"><img border="0" src="Color_Ranges/thumbs.asp?I=<%=rs("Color")%>"></a><%
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
<%End IF%>



<table border="0" width="100%" id="table2" cellspacing="0" cellpadding="0" height="193">
<tr>
<td><img border="0" src="images/email_footer_span.JPG" width="100%" height="193"></td>
<td width="752">
<img border="0" src="images/email%20footer.JPG" width="752" height="193"></td><td>
<img border="0" src="images/email_footer_span.JPG" width="100%" height="193"></td></tr>
</table>
<%
conn.Close ' Close Connection
Set conn = Nothing

%>