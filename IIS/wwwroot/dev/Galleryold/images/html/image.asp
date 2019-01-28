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
<%

xDb_Conn_Str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../../Gallary.mdb") & ";"


Response.Buffer = true

Set conn = Server.CreateObject("ADODB.Connection")
Dim rs, Img, conn, sql
Img = Request.QueryString("I")
conn.Open xDb_Conn_Str


sql="SELECT Images.Image, Images.Code, Images.Client, Ranges.Name, Descriptions.Name, Colors.Name, Images.Width, Images.Height, Images.Depth, IIf([Images.Price]=0,'P.O.A','R ' & [Images.Price]) AS Price, Images.Edge, Images.Info, Images.Image, Descriptions.Description, Ranges.Range FROM Descriptions RIGHT JOIN (Colors RIGHT JOIN (Ranges RIGHT JOIN Images ON Ranges.Range = Images.Range) ON Colors.Color = Images.Color) ON Descriptions.Description = Images.Description WHERE (((Images.Image)=" & Img & "))"
Set rs = Conn.Execute(sql)
%>

<html>
<head>
<title>Exclusive Cabinets Gallary</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<style>
<!--
 .t1{font-size:11pt;border: 2px solid #cc907e; padding: 0}

 .udot {border-bottom-style: dashed; border-bottom-width: 1px; padding: 2px; }
-->
</style>
</head>
<link rel="stylesheet" href="../../gallery.css" type="text/css">
<body background="../../images/creamGradient.jpg" style="background-repeat:repeat-x; background-attachment:fixed" leftmargin="0" rightmargin="0">
<div align="center">
<a href="http://www.exclusivecabinets.co.za">
<img border="2" src="../logo2.JPG" width="315" height="211" align="left"></a><br>
<table class="t1" cellSpacing="0" cellPadding="2" width="410" bgColor="#E4CAC0" border="0" id="table2">
	<tr>
		<td colSpan="2" height="25"><b><font size="3">Description:</font></b><font size="3"> 
		<%=rs.Fields("Descriptions.Name")%>&nbsp;&nbsp;&nbsp;(<a href="../../SimpleBrowser.asp@Des=<%=rs.Fields("Description")%>&Type.html">All Items</a>)</font></td>
	</tr>
	<tr>
		<td colSpan="2" height="25"><b><font size="3">Dimensions: </font></b>
		<font size="3"><%=rs.Fields("Width")%><b>&nbsp;X&nbsp;</b><%=rs.Fields("Height")%> 
		<b>&nbsp;X&nbsp;</b><%=rs.Fields("Depth")%></font><font size="2"> </font>
		<font size="2"> mm&nbsp;&nbsp; (</font><font color="#3C4976" size="2"><b>Alternative 
		Size Available</b></font><font size="2">)</font></td>
	</tr>
	<tr>
		<td colSpan="2" height="25"><b><font size="3">Range:</font></b><font size="3">&nbsp;<%=rs.Fields("Ranges.Name")%>
		&nbsp;&nbsp;&nbsp;&nbsp; (<a href="../../SimpleBrowser.asp@Range=<%=rs.Fields("Range")%>&Des=<%=rs.Fields("Description")%>&Type.html">In Current 
		Range</a>)</font></td>
	</tr>
	<tr>
		<td colSpan="2" height="25"><b><font size="3">Colour:</font></b><font size="3">&nbsp; 
		<%=rs.Fields("Colors.Name")%>&nbsp;&nbsp;&nbsp;&nbsp; (<a href="../../Available%20Colours.html">Available 
		Colours</a>)</font></td>
	</tr>
	<tr>
		<td colSpan="2" height="25"><b><font size="3">Edging:</font></b><font size="3"> 
		<%=rs.Fields("Edge")%></font></td>
	</tr>
	<tr>
		<td width="208" height="25"><b><font size="3">Code:</font></b><font size="3">&nbsp;
		<%=rs.Fields("Code")%></font></td>
		<td width="198" height="25"><b><font size="3">Price:</font></b><font size="3">&nbsp; 
		<%=rs.Fields("Price")%></font></td>
	</tr>
	<tr>
		<td width="406" colSpan="2" height="25"><b><font size="3">Client Ref:</font></b><font size="3">&nbsp;<%=rs.Fields("Client")%>
		</font></td>
	</tr>
	<tr>
		<td colSpan="2" style="padding-left: 10px; padding-right: 10px">
		<p align="left"><%=rs.Fields("Info")%></td>
	</tr>
</table>
<p align="left">&nbsp;
<a href="../../../index.html">Gallery Home</a>
</p>
<table border="0" id="table1" cellspacing="0" cellpadding="0" width="100%">

<%
rs.Close
sql="SELECT  Sub_Image.ID, Sub_Image.Des FROM Sub_Image WHERE ((Sub_Image.Image)=" & Img & ");"
Set rs = Conn.Execute(sql)
%>
<tr><td align="center"><img border="2" src="../large/image.asp@I=<%=Request.QueryString("I")%>&Type.jpg"></td></tr>
<tr><td height="31">&nbsp;</td></tr>
<%while Not rs.eof	%>	
<tr><td align="center">
<img border="2" src="../large/image.asp@I=<%=Img%>_<%=rs(0)%>&Type.jpg"></td></tr>
<tr><td align="center"><b><%=rs(1)%></b></td></tr>
<tr><td>&nbsp;<p>&nbsp;</td></tr>
	<%
	rs.MoveNext()
	Wend
	%>
	<tr height="0px">
		<td>
	
	<table border="0" width="100%" cellspacing="0" cellpadding="0" height="193px" id="table3">
	<tr>
		<td>
		<img border="0" src="../email_footer_span.JPG" width="100%" height="193"></td>
		<td width="752">
		<a href="http://www.exclusivecabinets.co.za/">
		<img border="0" src="../email%20footer.JPG" width="752" height="193"></a></td>
		<td>
		<img border="0" src="../email_footer_span.JPG" width="100%" height="193"></td>
	</tr>
	</table>

		</td>
	</tr>
</table>

</div>
</body>