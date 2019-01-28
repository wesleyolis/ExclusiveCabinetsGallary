<%@ CodePage = 1252 LCID = 7177 %>

<%
Response.expires = -1
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "private"
Response.CacheControl = "no-cache"
%>

<!--#include file="db.asp"-->
<% Set conn = Server.CreateObject("ADODB.Connection")
Dim rs, Img, conn, sql
Img = Request.QueryString("I")
conn.Open xDb_Conn_Str

sql="SELECT  Sub_Image.ID, Sub_Image.Des FROM Sub_Image WHERE ((Sub_Image.Image)=" & Img & ");"
Set rs = Conn.Execute(sql)
Dim preload
preload = ""
pos = 0
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Film Strip</title>
</head>
<%If Session("Admin") Then%>
<script language="javascript">
<!--
function edit(img)
{
parent.document.location.replace("Images_SubAdd.asp?Key="+img);
}

function image(img)
{
parent.document.location.replace("ImagesSubAdd.asp?I=<%=Img%>&Key="+img);
}

function del(img)
{
var answer = confirm ("Delete this Entry?") 
if (answer)
{
	var answer2 = confirm ("Delete the Original Image as well") 
	if(answer2){
	parent.document.location.replace("ImagesdeleteSub.asp?I=<%=Img%>&Key="+img+"&o=");}
	else{
	parent.document.location.replace("ImagesdeleteSub.asp?I=<%=Img%>&Key="+img);}
}
}
//-->
</script>
<%End IF%>
<body  background="images/creamGradient.jpg" style="background-repeat:repeat-x; background-attachment:fixed"  topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0" marginwidth="0" marginheight="0">
<table border="0" width="100%" id="table5" cellspacing="0" cellpadding="0" height="135">
	<tr>
		<td background="images/email_footer_span.JPG" height="135" align="center" width="20">
		&nbsp;</td>
		<td background="images/email_footer_span.JPG" height="135" align="center">
		<table border="0" id="table6" cellspacing="0" cellpadding="0" height="123">
			<tr><td width="20" height="123">&nbsp;</td>
			<%while Not rs.eof		
	preload = preload & " imgs[" & pos & "] =new Image();"
	preload = preload & " imgs[" & pos & "] .src='getimage.asp?I=" & Img & "_" & rs(0) & "&Resample=true&Quality=100&Width=495&Height=385';"
	preload = preload & " info[" & pos & "] ='" & rs(1) & "';"
%>
<td width="155" height="123" align="center"><a href="getimage.asp?I=<%=Img%>_<%=rs(0)%>&Resample=true&Quality=80" target="left">
<img border="0"  onmouseover="parent.setimage(<%=pos%>,info[<%=pos%>]);" src="thumbs/<%=Img%>_<%=rs(0)%>.jpg" style="border: 2px solid #000000"></a></td>
				
				<%If Session("Admin") Then%><td onmouseover="parent.setimage(<%=pos%>,info[<%=pos%>]);" width="60" valign="top" align="center">
				<input onclick="image('<%=rs(0)%>')" type="button" value="Image" name="Image">
				<input onclick="edit('<%=rs(0)%>')" type="button" value="Edit" name="Edit">
				<input onclick="del('<%=rs(0)%>')" type="button" value="Delete" name="Delete">
				</td><%End IF%>
				<td width="20">&nbsp;</td>
				<%
				pos = pos+1
				rs.MoveNext()
				Wend%>
				
			</tr>
		</table>
		</td>
		<td background="images/email_footer_span.JPG" height="135" align="center" width="20">&nbsp;</td>
	</tr>
</table>
<script language="javascript">
<!--
if(document.images)
{
var imgs = new Array();
var info = new Array();
<%=preload%>
if(imgs[0]!=null)
{
parent.setimage(0,info[0]);
}
}else{alert("Sorry your browser version is too old please upgrade or enable java script");}
//-->
</script>
</body>

</html>