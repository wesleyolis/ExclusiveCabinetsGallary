<%@ CodePage = 1252 LCID = 7177 %>
<% Session.Timeout = 20 %>
<%
Response.expires = 60

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



<table border="0" width="100%" height="95%" cellspacing="0" cellpadding="0">
<tr>
<td align="center" valign="top">
<table border="0" width="100%" height="100%" cellspacing="0" cellpadding="0"><tr><td width="30%"></td><td width="40%" align="center"><img border="0" src="images/logo2.JPG" width="315" height="211"></td><td width ="30%" align="right" valign="top">


<Form action="login.asp" method="post">
<table border="1" bordercolor="#CC907E" cellpadding="2" cellspacing="0"><%if Not Session("Admin") = True THEN%>
<tr><td>UserName</td><td><input type="text" name="user" id="user" size="16" maxlength="16"></td></tr>
<tr><td>Password</td><td>
	<input type="password" name="pass" id="pass" size="16" maxlength="16"></td></tr>
<tr><td colspan="2" align="right"><%IF Session("msg")<> "" THEN%><%=Session("msg") &"&nbsp;&nbsp;"%> <%END IF%>
<%If Request.QueryString("Email") = "true" Then%><input  type="hidden" name="Email" value="true"><%END IF%><input type="submit" name="Action" value="Login"></td></tr>
<%ELSE%>
<%If Request.QueryString("Email") = "true" Then%><input  type="hidden" name="Email" value="true"><%END IF%><input type="submit" name="Action" value="Logout"></td></tr>
<%END IF%>
</Table>
</Form>

</td></tr>
</Table>
</td>
<tr><td>
<table width="100%"><tr><td width="40%" align="right" valign="top">
<Form action="Dirbrowser.asp" method="get">
<%If Request.QueryString("Email")  = "true" Then%><input type="hidden" name="Email" value="true"><%END IF%>
<table border="1" bordercolor="#CC907E" cellpadding="3" cellspacing="0">
<tr><td>Ulti Sales</td><td><input type="text" name="Ulti" id="Ulti" size="16" maxlength="16"></td></tr>
<tr><td>Client Ref</td><td><input type="text" name="Client" id="Client" size="16" maxlength="16"></td></tr>
<tr><td colspan="2" align="right"><input type="submit" value="Search"></td></tr>
</Table>
</Form>

</td>
<td width="20%" align="center" valign="top"><A href="DirBrowser.asp?d=-1<%If Request.QueryString("Email") = "true" Then%>&Email=true<%END IF%>">Directory Browser</a><br><br>
<%If Not Request.QueryString("Email") = "true" Then%>
<A href="gallary_browser.htm">Email Browser</a><br><br>
<%END IF%>
<%if Session("Admin") = True THEN%><a href="Colorslist.asp?cmd=resetall">Edit Selection Categories</a><%END IF%>
</td>


<td width="40%"  valign="top">

<Form action="Dirbrowser.asp" method="get">
<%If Request.QueryString("Email") = "true" Then%><input  type="hidden" name="Email" value="true"><%END IF%>
<table border="1" bordercolor="#CC907E"  cellpadding="3" cellspacing="0">
<tr><td>Description</td><td>

<select class="input" id="Des" name="Des" size="1" >
<option value="-1">No Filter</option>
<% Set rs = conn.Execute("SELECT Descriptions.Description, Descriptions.Name FROM Descriptions ORDER BY Descriptions.Name;")
 Do While Not rs.Eof %>
<option value="<%=rs.Fields(0)%>"><%=rs.Fields(1)%></option>
<% rs.MoveNext
Loop %>
</select>

</td></tr>
<tr><td>Range</td><td>

<select class="input" id="Range" name="Range" size="1" >
<option value="-1">No Filter</option>
<% Set rs = conn.Execute("SELECT Ranges.Range, Ranges.Name FROM Ranges ORDER BY Ranges.Name;")
 Do While Not rs.Eof %>
<option value="<%=rs.Fields(0)%>"><%=rs.Fields(1)%></option>
<% rs.MoveNext
Loop %>
</select>


</td></tr>
<tr><td>Edging</td><td>

<select class="input" id="Color" name="Color" size="1" >
<option value="-1">No Filter</option>
<% Set rs = conn.Execute("SELECT Colors.Color, Colors.Name FROM Colors ORDER BY Colors.Name;")
 Do While Not rs.Eof %>
<option value="<%=rs.Fields(0)%>"><%=rs.Fields(1)%></option>
<% rs.MoveNext
Loop %>
</select>

</td></tr>
<tr><td colspan="2" align="right"><input type="submit" value="Filter"></td></tr>
</Table>
</Form>
<td></table>



</td></tr>


<tr height="193px"><td>
	
	<table border="0" width="100%" cellspacing="0" cellpadding="0" height="193px">
	<tr>
		<td>
		<img border="0" src="images/email_footer_span.JPG" width="100%" height="193"></td>
		<td width="752">
		<img border="0" src="images/email%20footer.JPG" width="752" height="193"></td>
		<td>
		<img border="0" src="images/email_footer_span.JPG" width="100%" height="193"></td>
	</tr>
	</table>

</td></tr>
</table>



</body>
</html>






