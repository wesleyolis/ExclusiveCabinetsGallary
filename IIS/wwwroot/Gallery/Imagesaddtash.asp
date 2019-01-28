<%
Response.Expires = 0
Response.ExpiresAbsolute = #1/1/1980# ' Expired
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private, no-cache, no-store, must-revalidate"
%>

<%
ewCurSec = 0 ' Initialise

' User levels
Const ewAllowAdd = 1
Const ewAllowDelete = 2
Const ewAllowEdit = 4
Const ewAllowView = 8
Const ewAllowList = 8
Const ewAllowReport = 8
Const ewAllowSearch = 8
Const ewAllowAdmin = 16
%>
<%

' Initialize common variables
x_Image = Null
x_Dir = Null
x_Code = Null
x_Client = Null
x_Range = Null
x_Description = Null
x_Color = Null
x_Width = Null
x_Height = Null
x_Depth = Null
x_Price = Null
x_Edge = Null
x_Info = Null
%>
<!--#include file="db.asp"-->
<!--#include file="aspmkrfn.asp"-->
<%
Response.Buffer = True

' Get action
sAction = Request.Form("a_add")
If (sAction = "" Or IsNull(sAction)) Then
	sKey = Request.Querystring("key")
	If sKey <> "" Then
		sAction = "C" ' Copy record
	Else
		sAction = "I" ' Display blank record
	End If
Else

	' Get fields from form
	x_Image = Request.Form("x_Image")
	x_Dir = Request.Form("x_Dir")
	x_Code = Request.Form("x_Code")
	x_Client = Request.Form("x_Client")
	x_Range = Request.Form("x_Range")
	x_Description = Request.Form("x_Description")
	x_Color = Request.Form("x_Color")
	x_Width = Request.Form("x_Width")
	x_Height = Request.Form("x_Height")
	x_Depth = Request.Form("x_Depth")
	x_Price = Request.Form("x_Price")
	x_Edge = Request.Form("x_Edge")
	x_Info = Request.Form("x_Info")
End If
image = ""
' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case sAction
	Case "C": ' Get a record to display
		If Not LoadData(sKey) Then ' Load Record based on key
			Session("ewmsg") = "No Record Found for Key = " & sKey
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "dirbrowser.asp?d=" & Request.QueryString("d")
		End If
		
	Case "A": ' Add
		If AddData() Then ' Add New Record
			Session("ewmsg") = "Add New Record Successful"
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "ImagesAdd2.asp?d=" & x_Dir & "&Key=" & image & "&f="
		End If
		
	Case "E":
		If 	EditData(x_Image) Then ' Add New Record
			Session("ewmsg") = "Update Record Successful"
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "dirbrowser.asp?d=" & x_Dir
		End If
		

	Case "D":
		If AddData() Then ' Add New Record
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		path = Server.MapPath("thumbs")
		objFSO.CopyFile path & "/" & x_Image & ".jpg", path & "/" & image & ".jpg"
		path = Server.MapPath("orig")
		objFSO.CopyFile path & "/" & x_Image & ".jpg", path & "/" & image & ".jpg"

		Set objFSO = Nothing
		
			Session("ewmsg") = "Add New Record Successful"
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "ImagesAdd2.asp?d=" & x_Dir & "&Key=" & image
		End If
		
		
	End Select
%>
<!--#include file="header.asp"-->
<script type="text/javascript" src="ew.js"></script>
<script type="text/javascript">
<!--
EW_dateSep = "/"; // set date separator	
//-->
</script>
<script type="text/javascript">
<!--
function EW_checkMyForm(EW_this) {
if (EW_this.x_Dir && !EW_hasValue(EW_this.x_Dir, "TEXT" )) {
	if (!EW_onError(EW_this, EW_this.x_Dir, "TEXT", "Please enter required field - Dir"))
		return false;
}
if (EW_this.x_Dir && !EW_checkinteger(EW_this.x_Dir.value)) {
	if (!EW_onError(EW_this, EW_this.x_Dir, "TEXT", "Incorrect integer - Dir"))
		return false; 
}
if (EW_this.x_Range && !EW_hasValue(EW_this.x_Range, "TEXT" )) {
	if (!EW_onError(EW_this, EW_this.x_Range, "TEXT", "Please enter required field - Range"))
		return false;
}
if (EW_this.x_Range && !EW_checkinteger(EW_this.x_Range.value)) {
	if (!EW_onError(EW_this, EW_this.x_Range, "TEXT", "Incorrect integer - Range"))
		return false; 
}
if (EW_this.x_Description && !EW_hasValue(EW_this.x_Description, "TEXT" )) {
	if (!EW_onError(EW_this, EW_this.x_Description, "TEXT", "Please enter required field - Description"))
		return false;
}
if (EW_this.x_Description && !EW_checkinteger(EW_this.x_Description.value)) {
	if (!EW_onError(EW_this, EW_this.x_Description, "TEXT", "Incorrect integer - Description"))
		return false; 
}
if (EW_this.x_Color && !EW_hasValue(EW_this.x_Color, "TEXT" )) {
	if (!EW_onError(EW_this, EW_this.x_Color, "TEXT", "Please enter required field - Color"))
		return false;
}
if (EW_this.x_Color && !EW_checkinteger(EW_this.x_Color.value)) {
	if (!EW_onError(EW_this, EW_this.x_Color, "TEXT", "Incorrect integer - Color"))
		return false; 
}
if (EW_this.x_Width && !EW_checkinteger(EW_this.x_Width.value)) {
	if (!EW_onError(EW_this, EW_this.x_Width, "TEXT", "Incorrect integer - Width"))
		return false; 
}
if (EW_this.x_Height && !EW_checkinteger(EW_this.x_Height.value)) {
	if (!EW_onError(EW_this, EW_this.x_Height, "TEXT", "Incorrect integer - Height"))
		return false; 
}
if (EW_this.x_Depth && !EW_checkinteger(EW_this.x_Depth.value)) {
	if (!EW_onError(EW_this, EW_this.x_Depth, "TEXT", "Incorrect integer - Depth"))
		return false; 
}
if (EW_this.x_Price && !EW_checkinteger(EW_this.x_Price.value)) {
	if (!EW_onError(EW_this, EW_this.x_Price, "TEXT", "Incorrect integer - Price"))
		return false; 
}
return true;
}
//-->
</script>

<link rel="stylesheet" href="gallery.css" type="text/css">

<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.x_Code.value.length > 16)
  {
    alert("Please enter at most 16 characters in the \"x_Code\" field.");
    theForm.x_Code.focus();
    return (false);
  }

  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzƒŠŒŽšœžŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ0123456789-";
  var checkStr = theForm.x_Code.value;
  var allValid = true;
  var validGroups = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Please enter only letter and digit characters in the \"x_Code\" field.");
    theForm.x_Code.focus();
    return (false);
  }

  if (theForm.x_Client.value.length > 16)
  {
    alert("Please enter at most 16 characters in the \"x_Client\" field.");
    theForm.x_Client.focus();
    return (false);
  }

  var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzƒŠŒŽšœžŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ0123456789-";
  var checkStr = theForm.x_Client.value;
  var allValid = true;
  var validGroups = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Please enter only letter and digit characters in the \"x_Client\" field.");
    theForm.x_Client.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form name="FrontPage_Form1" id="Imagesadd" action="Imagesadd.asp" method="post" onSubmit="return FrontPage_Form1_Validator(this)" language="JavaScript">
<p>
<input type="hidden" name="a_add" value="A">
<input hidden type="hidden" name="x_Dir" id="x_Dir" value="<%=Request.QueryString("d")%>">
<input hidden type="hidden" name="x_Image" id="x_Image" value="<%=sKey%>">
<table border="1" cellspacing="1" cellpadding="4" bgcolor="#F5F5F5" width="692" class="text" height="344" style="border-collapse: collapse">
	<tr>
		<td width="619" colspan="4" class="box2">
		<Span class="header">Add Image to gallery</Span></td>
		
	</tr>
	<tr>
		<td width="619" height="10" colspan="4"></td>
		
	</tr>
	<tr>
		<td height="40"  class="box2" width="58">Code</td>
		
		<td bgcolor="#F5F5F5" width="200">
&nbsp;<!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" i-maximum-length="16" --><input  class="input" type="text" name="x_Code" id="x_Code" size="30" maxlength="16" value="<%= Server.HTMLEncode(x_Code&"") %>" tabindex="1">
</td>
<td  bgcolor=#F5F5F5 width="29" rowspan="9">&nbsp;</td>
		<td  bgcolor="#3366CC" width="369"  class="box2" height="40">Info</td>
	</tr>
	<tr>
		<td bgcolor="#3366CC" height="29" class="box2" width="58">Client Ref:</td>
		<td bgcolor="#F5F5F5" height="29" width="194">
		<!--webbot bot="Validation" s-data-type="String" b-allow-letters="TRUE" b-allow-digits="TRUE" i-maximum-length="16" --><input  class="input" type="text" name="x_Client" id="x_Client" size="30" maxlength="16" value="<%= Server.HTMLEncode(x_Client&"") %>" tabindex="2"></td>
		<td bgcolor="#F5F5F5" rowspan="8" width="369"><span class="aspmaker">
<textarea  class="text" cols="43" rows="15" id="x_Info" name="x_Info" style="width: 369; height: 250" tabindex="11"><%= x_Info %></textarea></span></td>
	</tr>
	<tr>
		<td bgcolor="#3366CC" height="29" class="box2" width="58">Range</td>
		<td bgcolor="#F5F5F5" height="29" width="194">
<% If IsNull(x_Range) or x_Range = "" Then x_Range = 0 ' Set default value %>

<!--#include file="ranges.asp"-->
<%
Call List_ranges(x_Range)
%>
</td>
	</tr>
	<tr>
		<td bgcolor="#3366CC" height="29"  class="box2" width="58">
		Description</td>
		<td bgcolor="#F5F5F5" width="194">
<% If IsNull(x_Description) or x_Description = "" Then x_Description = 0 ' Set default value %>
<!--#include file="Descriptions.asp"-->
<%
Call List_Descriptions(x_Description)
%>
</td>
	</tr>
	<tr>
		<td bgcolor="#3366CC" height="29"  class="box2" width="58">
		Colour</td>
		<td bgcolor="#F5F5F5" width="194">
<% If IsNull(x_Color) or x_Color = "" Then x_Color = 0 ' Set default value %>
<!--#include file="Colours.asp"-->
<%
Call List_Colours(x_Color)
%>

</td>
	</tr>
	<tr>
		<td bgcolor="#3366CC" height="29"  class="box2" width="58">
		Width</td>
		<td bgcolor="#F5F5F5" width="194">
<% If IsNull(x_Width) or x_Width = "" Then x_Width = 0 ' Set default value %>
<input  class="text" type="text" name="x_Width" id="x_Width" size="5" value="<%= Server.HTMLEncode(x_Width&"") %>" style="height: 18" tabindex="6"> 
<b>X</b>
<% If IsNull(x_Height) or x_Height = "" Then x_Height = 0 ' Set default value %>
<input  class="text" type="text" name="x_Height" id="x_Height" size="5" value="<%= Server.HTMLEncode(x_Height&"") %>" style="height: 18" tabindex="7">
<b>X</b>
<% If IsNull(x_Depth) or x_Depth = "" Then x_Depth = 0 ' Set default value %>
<input  class="text" type="text" name="x_Depth" id="x_Depth" size="5" value="<%= Server.HTMLEncode(x_Depth&"") %>" style="height: 18" tabindex="8"></td>
	</tr>
	<tr>
		<td bgcolor="#3366CC" height="29"  class="box2" width="58">Price</td>
		<td bgcolor="#F5F5F5" width="194">
<% If IsNull(x_Price) or x_Price = "" Then x_Price = 0 ' Set default value %><input  class="input" type="text" name="x_Price" id="x_Price" size="30" value="<%= Server.HTMLEncode(x_Price&"") %>" tabindex="9">
</td>
	</tr>
	<tr>
		<td bgcolor="#3366CC" height="29"  class="box2" width="58">Edge</td>
		<td bgcolor="#F5F5F5" width="194">
		<input class="input" type="text" name="x_Edge" id="x_Edge" size="30" maxlength="25" value="<%= Server.HTMLEncode(x_Edge&"") %>" tabindex="10"></td>
	</tr>
	<tr>
		<td bgcolor="#F5F5F5" width="252" colspan="2" style="height: 34px">
		<input  class="text" onclick="document.location.replace('DirBrowser.asp?d=<%=Request.QueryString("d")%>')" type="button" value="Cancel" name="Cancel" tabindex="15">&nbsp;&nbsp;
		<%If sAction = "I" Then%>
		<input  class="text" type="submit" name="Action" value="ADD" tabindex="12">
		<%ELSE%>
		<input class="text" onclick="document.all.a_add.value='D'" type="submit" name="Action" value="Add Copy" tabindex="15">
		<input class="text" onclick="document.all.a_add.value='E'" type="submit" name="Action" value="Update" tabindex="13">
		<%End If%>
		</td>
	</tr>
		

</table>

<p>
&nbsp;</form>
<!--#include file="footer.asp"-->
<p>&nbsp;</p>



<%
conn.Close ' Close Connection
Set conn = Nothing
%>
<%

'-------------------------------------------------------------------------------
' Function LoadData
' - Load Data based on Key Value sKey
' - Variables setup: field variables

Function LoadData(sKey)
	Dim sKeyWrk, sSql, rs, sWhere, sGroupBy, sHaving, sOrderBy
	sKeyWrk = "" & AdjustSql(sKey) & ""
	sSql = "SELECT * FROM [Images]"
	sSql = sSql & " WHERE [Image] = " & sKeyWrk
	sGroupBy = ""
	sHaving = ""
	sOrderBy = ""
	If sGroupBy <> "" Then
		sSql = sSql & " GROUP BY " & sGroupBy
	End If	
	If sHaving <> "" Then
		sSql = sSql & " HAVING " & sHaving
	End If	
	If sOrderBy <> "" Then
		sSql = sSql & " ORDER BY " & sOrderBy
	End If	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sSql, conn
	If rs.Eof Then
		LoadData = False
	Else
		LoadData = True
		rs.MoveFirst

		' Get the field contents
		x_Image = rs("Image")
		x_Dir = rs("Dir")
		x_Code = rs("Code")
		x_Client = rs("Client")
		x_Range = rs("Range")
		x_Description = rs("Description")
		x_Color = rs("Color")
		x_Width = rs("Width")
		x_Height = rs("Height")
		x_Depth = rs("Depth")
		x_Price = rs("Price")
		x_Edge = rs("Edge")
		x_Info = Replace(rs("Info")&"","<br>",VbCrLf)
	End If
	rs.Close
	Set rs = Nothing
End Function
%>
<%

'-------------------------------------------------------------------------------
' Function AddData
' - Add Data
' - Variables used: field variables



Function AddData()
	Dim sSql, rs, sWhere, sGroupBy, sHaving, sOrderBy

	' Add New Record
	sSql = "SELECT * FROM [Images]"
	sSql = sSql & " WHERE 0 = 1"
	sGroupBy = ""
	sHaving = ""
	sOrderBy = ""
	If sGroupBy <> "" Then
		sSql = sSql & " GROUP BY " & sGroupBy
	End If	
	If sHaving <> "" Then
		sSql = sSql & " HAVING " & sHaving
	End If	
	If sOrderBy <> "" Then
		sSql = sSql & " ORDER BY " & sOrderBy
	End If	
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open sSql, conn, 1, 2
	rs.AddNew

	' Field Dir
	sTmp = x_Dir
	If Not IsNumeric(sTmp) Then
		sTmp = 0
	Else
		sTmp = CLng(sTmp)
	End If
	rs("Dir") = sTmp

	' Field Code
	sTmp = Trim(x_Code)
	If Trim(sTmp) = "" Then sTmp = Null
	rs("Code") = sTmp
	
	' Field Client
	sTmp = Trim(x_Client)
	If Trim(sTmp) = "" Then sTmp = Null
	rs("Client") = sTmp


	' Field Range
	sTmp = x_Range
	If Not IsNumeric(sTmp) Then
		sTmp = 0
	Else
		sTmp = CLng(sTmp)
	End If
	rs("Range") = sTmp

	' Field Description
	sTmp = x_Description
	If Not IsNumeric(sTmp) Then
		sTmp = 0
	Else
		sTmp = CLng(sTmp)
	End If
	rs("Description") = sTmp

	' Field Color
	sTmp = x_Color
	If Not IsNumeric(sTmp) Then
		sTmp = 0
	Else
		sTmp = CLng(sTmp)
	End If
	rs("Color") = sTmp

	' Field Width
	sTmp = x_Width
	If Not IsNumeric(sTmp) Then
		sTmp = Null
	Else
		sTmp = CLng(sTmp)
	End If
	rs("Width") = sTmp

	' Field Height
	sTmp = x_Height
	If Not IsNumeric(sTmp) Then
		sTmp = Null
	Else
		sTmp = CLng(sTmp)
	End If
	rs("Height") = sTmp

	' Field Depth
	sTmp = x_Depth
	If Not IsNumeric(sTmp) Then
		sTmp = Null
	Else
		sTmp = CLng(sTmp)
	End If
	rs("Depth") = sTmp

	' Field Price
	sTmp = x_Price
	If Not IsNumeric(sTmp) Then
		sTmp = Null
	Else
		sTmp = CLng(sTmp)
	End If
	rs("Price") = sTmp

	' Field Edge
	sTmp = Trim(x_Edge)
	If Trim(sTmp) = "" Then sTmp = Null
	rs("Edge") = sTmp

	' Field Info
	sTmp = Replace(Trim(x_Info),VbCrLf,"<br>")
	If Trim(sTmp) = "" Then sTmp = Null
	rs("Info") = sTmp
	rs("Sync_info") = "True"
	rs("Sync_img") = "True"
	rs.Update
	rs.Close
	Set rs = Nothing
	
	Set rs = conn.Execute("SELECT @@IDENTITY;")
		If Not rs.Eof Then
		image = rs(0)
		End IF
	rs.Close
	Set rs = Nothing
	
	AddData = True
End Function

' Function EditData
' - Edit Data based on Key Value sKey
' - Variables used: field variables

Function EditData(sKey)
	Dim sKeyWrk, sSql, rs, sWhere, sGroupBy, sHaving, sOrderBy

	' Open record
	sKeyWrk = "" & AdjustSql(sKey) & ""
	sSql = "SELECT * FROM [Images]"
	sSql = sSql & " WHERE [Image] = " & sKeyWrk
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open sSql, conn, 1, 2
	If rs.Eof Then
		EditData = False ' Update Failed
	Else

		sTmp = x_Dir
		If Not IsNumeric(sTmp) Then
			sTmp = 0
		Else
			sTmp = CLng(sTmp)
		End If
		rs("Dir") = sTmp

		' Field Code
		sTmp = Trim(x_Code)
		If Trim(sTmp) = "" Then sTmp = Null
		rs("Code") = sTmp
		
		' Field Client
		sTmp = Trim(x_Client)
		If Trim(sTmp) = "" Then sTmp = Null
		rs("Client") = sTmp


		' Field Range
		sTmp = x_Range
		If Not IsNumeric(sTmp) Then
			sTmp = 0
		Else
			sTmp = CLng(sTmp)
		End If
		rs("Range") = sTmp

		' Field Description
		sTmp = x_Description
		If Not IsNumeric(sTmp) Then
			sTmp = 0
		Else
			sTmp = CLng(sTmp)
		End If
		rs("Description") = sTmp

		' Field Color
		sTmp = x_Color
		If Not IsNumeric(sTmp) Then
			sTmp = 0
		Else
			sTmp = CLng(sTmp)
		End If
		rs("Color") = sTmp

		' Field Width
		sTmp = x_Width
		If Not IsNumeric(sTmp) Then
			sTmp = Null
		Else
			sTmp = CLng(sTmp)
		End If
		rs("Width") = sTmp

		' Field Height
		sTmp = x_Height
		If Not IsNumeric(sTmp) Then
			sTmp = Null
		Else
			sTmp = CLng(sTmp)
		End If
		rs("Height") = sTmp

		' Field Depth
		sTmp = x_Depth
		If Not IsNumeric(sTmp) Then
			sTmp = Null
		Else
			sTmp = CLng(sTmp)
		End If
		rs("Depth") = sTmp

		' Field Price
		sTmp = x_Price
		If Not IsNumeric(sTmp) Then
			sTmp = Null
		Else
			sTmp = CLng(sTmp)
		End If
		rs("Price") = sTmp

		' Field Edge
		sTmp = Trim(x_Edge)
		If Trim(sTmp) = "" Then sTmp = Null
		rs("Edge") = sTmp

		' Field Info
		sTmp = Replace(Trim(x_Info),VbCrLf,"<br>")

		If Trim(sTmp) = "" Then sTmp = Null
		rs("Info") = sTmp
		
		rs("Sync_info") = "True"
	
		rs.Update
		EditData = True ' Update Successful
	End If
	rs.Close
	Set rs = Nothing
End Function


%>