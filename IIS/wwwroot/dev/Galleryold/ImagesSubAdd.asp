<%
Response.Expires = 0
Response.ExpiresAbsolute = #1/1/1980# ' Expired
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private, no-cache, no-store, must-revalidate"

If Request.QueryString("I").Count > 0 And Request.QueryString("key").Count > 0 Then
%>

<!--#include file="header.asp"-->
<script language="javascript">
<!--
function cheack_filetype()
{

file = document.all.File1.value;
if(file != "")
{
extypes = new Array("jpg", "gif", "bmp", "png", "pcx", "psd", "tif", "wbmp");
extension = file.substring(file.lastIndexOf(".")+1,file.length).toLowerCase();

	for(i = 0;i<8;i++)
	{
		if(extension  == extypes[i])
		{return true;} 
	}
document.forms(0).reset();
alert("Invalied Extension '"+extension+"', should be\n jpg, gif, bmp, png, pcx, psd, tif, wbmp");
}
else
{
alert("No image selected, should be\n jpg, gif, bmp, png, pcx, psd, tif, wbmp");
}
return false

}
//-->
</script>
<form method="post" encType="multipart/form-data" onSubmit="return cheack_filetype();" action="ToFileSystemSub.asp?I=<%=Request.QueryString("I")%>&Key=<%=Request.QueryString("key")%>">
<table border="1" cellpadding="3" style="border-collapse: collapse" class="text">
	<tr>
		<td bgcolor="#3366CC" class="box2" colspan="2">Select image to bundle 
		with info</td>
	</tr>
	<tr>
		<td bgcolor="#3366CC" class="box2" width="66">Image</td>
		<td bgcolor="#F5F5F5">
	&nbsp;<INPUT  onchange="cheack_filetype();" type="File" name="File1" size="89" class="text"></td>
	</tr>
	<tr>
		<td bgcolor="#F5F5F5" colspan="2" align="right">
		<%If Request.QueryString("f").count = 0 Then%>
		<input  class="text" onclick="document.location.replace('Sub_Browser.asp?I=<%=Request.QueryString("I")%>')" type="button" value="Cancel" name="Cancel">&nbsp;&nbsp;
		<%End If%>
		<input type="submit" value="Upload" name="B1" class="text"></td>
	</tr>
</table>


</form>
	<!--#include file="footer.asp"-->
<%
Else
Response.Redirect("Error.asp")
End If%>