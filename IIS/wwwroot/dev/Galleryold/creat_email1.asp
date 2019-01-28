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
Dim rsi
imags = ""
pos = 1
While pos <= Request.QueryString("I").Count
imgs = imgs & " OR ([Images.Image]=" & Request.QueryString("I")(pos) & ")"
pos = pos + 1
Wend

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str


sql = "SELECT Images.Image, Images.Code, Ranges.Name, Descriptions.Name, Colors.Name, Images.Width, Images.Height, Images.Depth,IIF([Images.Price]=0,'P.O.A','R ' & [Images.Price]) As Price,Images.Edge,  IIf(IsNull([Info]),' ',[Info]) As Info2 "_
&" FROM Descriptions RIGHT JOIN (Colors RIGHT JOIN (Ranges RIGHT JOIN Images ON Ranges.Range = Images.Range) ON Colors.Color = Images.Color) ON Descriptions.Description = Images.Description"_
&" WHERE (([Images.Image]= -1) " & imgs & ")"



set rsi = Server.CreateObject("ADODB.recordset")
rsi.Open sql, conn
%>

<html>
<head>
<title>Exclusive Cabinets Gallary</title>
</head>
<style>
 .t1{font-size:11pt;border: 2px solid #cc907e; padding: 0; color:#333E64}
  .udot {border-bottom-style: dashed; border-bottom-width: 1px; padding: 2px; }

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

.enter a {
	font-family: arial;
	font-size: 12px;
	font-weight: bold;
	color: #000000;
	text-decoration: none;
	align: left;
	margin: 5px 50px 5px 50px;	
	}

.enter a:hover {
	color: #1e2b6f;
	text-decoration: none;
	border-bottom: 1px dotted #1e2b6f;
	padding-bottom: 2px;
	}

.online {
	border-right: 1px dotted #1e2b6f;
	border-left: 1px dotted #1e2b6f;
	}
</style>
<body background="images/creamGradient.jpg" style="background-repeat:repeat-x; background-attachment:fixed" topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">

<div align="center">

<div align="left">
	<table border="0" width="100%" id="table3" style="font-size:11pt;color:#333E64;"  cellspacing="0" cellpadding="0">
		<tr>
			<td align="left" valign="top">
			<table border="0" width="100%" id="table4" cellspacing="0" cellpadding="0" style="font-size: 11pt; color: #333E64; font-weight: bold">
				<tr>
					<td width="95" height="25"><b>Proposal for:</b></td>
					<td height="23">&nbsp;</td>
				</tr>
				<tr>
					<td width="95" height="25"><b>Attention:</b></td>
					<td height="21"></td>
				</tr>
				<tr>
					<td width="95" valign="top"><b>Comments:</b></td>
					<td height="164" align="left" valign="top">&nbsp;</td>
				</tr>
			</table>
			</td>
			<td width="286">
			<img border="0" src="images/logo2.JPG" width="315" height="211"></td>
		</tr>
	</table>
</div>

<table border="0" cellpadding="0" cellspacing="0">
<tr><td>
<%
Rows = 2
RCount = 0
%>
<table border="0" cellpadding="0" cellspacing="20">
<%
Do While Not rsi.eof
If Rcount = 0 Then%><tr>
<% End IF%>
<td valign="top">
<table border="0" cellpadding="0" cellspacing="0" width="300" bgcolor="#EEDDD7" class="t1">
<tr><td colspan="2" height="188" class="udot" align="center">
<img border="1" src="thumbs/<%=rsi.Fields("Image")%>.jpg"></td></tr><tr>
<td colspan="2"><b>&nbsp;Description:</b> <%=rsi.Fields("Descriptions.Name")%></td></tr>
<tr><td colspan="2"><b>&nbsp;Dimensions: </b><%=rsi.Fields("Width")%> <b>X</b> <%=rsi.Fields("Height")%> <b>X</b> 
<%=rsi.Fields("Depth")%><font size="2"> in mm</font></td></tr>
<tr><td colspan="2"><b>&nbsp;Range:</b>&nbsp; <%=rsi.Fields("Ranges.Name")%></td></tr>
<tr><td colspan="2"><b>&nbsp;Colour:</b> <%=rsi.Fields("Colors.Name")%></td></tr>
<tr><td colspan="2"><b>&nbsp;Edging:</b> <%=rsi.Fields("Edge")%></td></tr>
<tr><td width="162"><b>&nbsp;Code:</b>&nbsp;&nbsp;&nbsp; <%=rsi.Fields("Code")%>&nbsp; </td>
<td height="22" width="134"><b>Price:</b> <%=rsi.Fields("Price")%></td></tr>

<tr><td colspan="2" height="30"><b><a title="<%=replace(rsi.Fields("Info2"),"<br>","  ")%>" href="">&nbsp; More Info - Hover Mouse Here</a></b></td>
</tr>
</table>
</td>
<%
RCount = RCount + 1
If RCount >= Rows Then
RCount = 0%>
</tr>
<%End IF%>

<%
rsi.MoveNext
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
		<td>
		<img border="0" src="images/email_footer_span.JPG" width="100%" height="193"></td>
		<td width="752">
		<img border="0" src="images/email%20footer.JPG" width="752" height="193"></td>
		<td>
		<img border="0" src="images/email_footer_span.JPG" width="100%" height="193"></td>
	</tr>
</table>



<%
conn.Close ' Close Connection
Set conn = Nothing

%>