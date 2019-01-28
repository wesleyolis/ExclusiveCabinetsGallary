<%@ CodePage = 1252 LCID = 7177 %>
<% Session.Timeout = 20 %>

<%
Response.expires = -1
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "private"
Response.CacheControl = "no-cache"
%>
<!--#include file="db.asp"-->

<% Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str

%>

<html>
<head>
<meta http-equiv="Content-Language" content="en-us">
<title>Exclusive Cabinets Gallary</title>
</head>
<link rel="stylesheet" href="gallery.css" type="text/css">
<style>
 .udot {border-bottom-style: dashed; border-bottom-width: 1px; padding: 2px; }
 .t1{font-size:11pt;border: 2px solid #cc907e; padding: 0}
</style>
<script language="javascript">
<!--
var i,cap;
function setimage(img,c)
{
i = img;
cap = c;
setimage2();
}

function setimage2()
{
//alert(i);
downloaded = Strip.imgs[i].complete;
//alert(downloaded);
if(!downloaded)
{
disimg.src = "images/logo2.jpg";
dis.innerHTML="<b>" + cap + "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Downloading Image..</b>";
setTimeout("setimage2()",500)
}
else
{
dis.innerHTML="<b>" + cap + "</b>";
disimg.src = Strip.imgs[i].src;
}
}

function addsub(img)
{
document.location.replace("Images_SubAdd.asp?I="+img);
}

//-->
</script>

<body background="images/creamGradient.jpg" style="background-repeat:repeat-x; background-attachment:fixed"  topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">



<table border="0" width="100%" id="table1" cellspacing="0" cellpadding="0" height="100%">
	<tr>
		<td width="100%" align="left" valign="top" colspan="2" height="10" >
		</td>
	</tr>
	<tr >
		<td height="395" width="50%" align="center" valign="middle" ><!-- getimage.asp?I=<%=Request.QueryString("I")%>&Resample=true&Quality=80&Width=495&Height=395 -->
		&nbsp;<a href="getimage.asp?I=<%=Request.QueryString("I")%>&Resample=true&Quality=80" target="_top"><img border="2" src="thumbs\<%=Request.QueryString("I")%>_Large.jpg"></a></td>
		<td height="385" width="47%" align="center" valign="middle">
		<img name="disimg" border="2" src="images/logo2.JPG"></td>
	</tr>
	<tr>
		<td align="left" >
		&nbsp;&nbsp;
		<a href="DirBrowser.asp?d=-1">Root Dir</a>&nbsp;&nbsp;
		<%if Session("Admin") = true Then%>	
			<input onclick="addsub('<%=Request.QueryString("I")%>')" type="button" value="Add Sub" name="Sub Image"><%End IF%>
		</td>
		<td width="47%" align="center" id="dis" class="header2">
		&nbsp;</td>
	</tr>
	<tr>
		<td height="160px" colspan="2">
		<iframe name="Strip" width="100%" height="157" src="Sub_Film_Strip.asp?I=<%=Request.QueryString("I")%>" scrolling="yes" style="border: 1px solid #FFFFFF" marginwidth="1" marginheight="1" border="0" frameborder="0">
		Your browser does not support inline frames or is currently configured not to display inline frames.
		</iframe></td>
	</tr>
	<tr><td colspan="2">&nbsp;</td></tr>
	</table>



</body>
</html>
