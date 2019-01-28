<%@ CodePage = 1252 LCID = 7177 %>
<% Session.Timeout = 20 %>
<%
Response.expires = 60
xDb_Conn_Str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("Gallery/Gallary.mdb") & ";"

%>


<% Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str

%>

<html>
<head>
<title>Exclusive Cabinets Gallary</title>
</head>
<link rel="stylesheet" href="Gallery/gallery.css" type="text/css">
<style>
 .udot {border-bottom-style: dashed; border-bottom-width: 1px; padding: 2px; }
 .t1{font-size:11pt;border: 2px solid #cc907e; padding: 0}
</style>
<body background="Gallery/images/creamGradient.jpg" style="background-repeat:repeat-x; background-attachment:fixed"  topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">



<table border="0" width="100%" height="95%" cellspacing="0" cellpadding="0">
<tr>
<td align="center" valign="top">
<table border="0" width="100%" height="100%" cellspacing="0" cellpadding="0"><tr><td width="30%"></td><td width="40%" align="center"><img border="0" src="Gallery/images/logo2.JPG" width="315" height="211"></td><td width ="30%" align="right" valign="top">

</td></tr>
</Table>
</td>
<tr><td>
<table width="100%"><tr><td width="40%" align="right" valign="top">
</td>
<td width="20%" align="center" valign="top">
<script type="text/javascript">
<!--
function goto()
{
url="";
if(document.all.Des.value == -1 & document.all.Range.value == -1)
{
parent.document.location = ("Gallery/SimpleBrowser.asp@Type.html");
}
else
{
if(document.all.Range.value != -1)
{url="Range=" + document.all.Range.value}

if(document.all.Des.value != -1)
{if(url!=""){url+="&"};
url+="Des=" + document.all.Des.value}
parent.document.location = ("Gallery/SimpleBrowser.asp@" + url + "&Type.html");
}


}
//-->
</script>
<font size="3">
	<a href="../welcome.htm">Home</a> </font>
<Form>
<%If Request.QueryString("Email") = "true" Then%><input  type="hidden" name="Email" value="true"><%END IF%>
<table border="1" bordercolor="#CC907E"  cellpadding="3" cellspacing="0">
<tr><td>Description</td><td>

<select width="200px" class="input" id="Des" name="Des" size="1" style="width: 200" >
<option value="-1">No Filter</option>
<% Set rs = conn.Execute("SELECT Descriptions.Description, Descriptions.Name FROM Descriptions  INNER JOIN Images ON Descriptions.Description = Images.Description GROUP BY Descriptions.Description, Descriptions.Name Order by Descriptions.Name")
 Do While Not rs.Eof %>
<option value="<%=rs.Fields(0)%>"><%=rs.Fields(1)%></option>
<% rs.MoveNext
Loop %>
</select>

</td></tr>
<tr><td>Range</td><td>

<select width="200px" class="input" id="Range" name="Range" size="1" style="width: 200" >
<option value="-1">No Filter</option>
<% Set rs = conn.Execute("SELECT Ranges.Range, Ranges.Name FROM Ranges  INNER JOIN Images ON Ranges.Range = Images.Range GROUP BY Ranges.Range, Ranges.Name ORDER BY Ranges.Name")
 Do While Not rs.Eof %>
<option value="<%=rs.Fields(0)%>"><%=rs.Fields(1)%></option>
<% rs.MoveNext
Loop %>
</select>


</td></tr>
<tr><td colspan="2" align="right"><font size="3">
	<a href="Gallery/Color_Ranges/Colors.asp@&Type.html">All Board Colours</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font><input onclick="goto()" type="button" value="Filter"></td></tr>
</Table>
</Form>

</td>


<td width="40%"  valign="top">
<td></table>



</td></tr>


<tr height="193px"><td>
	
	<table border="0" width="100%" cellspacing="0" cellpadding="0" height="193px">
	<tr>
		<td>
		<img border="0" src="Gallery/images/email_footer_span.JPG" width="100%" height="193"></td>
		<td width="752">
		<img border="0" src="Gallery/images/email%20footer.JPG" width="752" height="193"></td>
		<td>
		<img border="0" src="Gallery/images/email_footer_span.JPG" width="100%" height="193"></td>
	</tr>
	</table>

</td></tr>
</table>



</body>
</html>






