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
<div align="left">
<table border="0" width="100%" id="table3" style="font-size:11pt;color:#333E64;"  cellspacing="0" cellpadding="0">
<tr><td align="left" valign="top">

<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0" style="font-size: 11pt; color: #333E64; font-weight: bold">
<tr><td width="95" height="25"><b>Proposal for:</b></td>
<td height="23">&nbsp;</td></tr>

<tr><td width="95" height="25"><b>Attention:</b></td>
<td height="21"></td></tr>

<tr><td width="95" valign="top"><b>Comments:</b></td>
<td height="164" align="left" valign="top">&nbsp;</td></tr></table>
			
</td>
<td width="286"><img border="0" src="images/logo2.JPG" width="315" height="211"></td></tr>
</table>
</div>

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
<img border="1" src="thumbs/<%=img%>.jpg"></td>
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
<tr><td  height="117" width="151" align="center"><img border="1" src="thumbs/<%=img%>_<%=rsi.Fields("ID")%>.jpg"></td></tr>
<tr><td align="center"  valign="top"><%=rsi.Fields("Des")%><br> test</td></tr>
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
If MR =< 1  Then%>
<tr><td valign="top" height="25px" colspan="2"><u><b>More Info</b></u></td></tr>
<tr><td colspan="2"><%=Info%></td></tr><%
End IF%>
</table>

</td></tr><tr>
<td colspan="2"><b>&nbsp;Description:</b> <%=des%></td></tr>
<tr><td colspan="2"><b>&nbsp;Dimensions: </b><%=Width%> <b>X</b> <%=Height%> <b>X</b> 
<%=Depth%><font size="2"> in mm</font></td></tr>
<tr><td colspan="2"><b>&nbsp;Range:</b>&nbsp; <%=Range%></td></tr>
<tr><td colspan="2"><b>&nbsp;Colour:</b> <%=Color%></td></tr>
<tr><td colspan="2"><b>&nbsp;Edging:</b> <%=Edge%></td></tr>
<tr><td width="194"><b>&nbsp;Code:</b>&nbsp;&nbsp;&nbsp; <%=Code%>&nbsp; </td>
<td height="22" width="122"><b>Price:</b> <%=Price%></td></tr>
<%If MR =< 1  Then%>
<tr><td colspan="2" height="1"></td></tr>
<%Else%>
<tr><td colspan="2" height="35"><b><a title="<%=Info%>" href="">&nbsp; More Info - Hover Mouse Here</a></b></td></tr>
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