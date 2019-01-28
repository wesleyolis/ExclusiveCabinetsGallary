<!--#include file="ewconfig.asp"-->
<!--#include file="db.asp"-->
<!--#include file="Imagesinfo.asp"-->
<!--#include file="advsecu.asp"-->
<!--#include file="aspmkrfn.asp"-->
<!--#include file="ewupload.asp"-->
<%
Response.Expires = 0
Response.ExpiresAbsolute = #1/1/1980# ' Expired
Response.AddHeader "pragma", "no-cache"
Response.AddHeader "cache-control", "private, no-cache, no-store, must-revalidate"
%>
<%

' Initialize common variables
x_Image = Null: ox_Image = Null: z_Image = Null
x_Dir = Null: ox_Dir = Null: z_Dir = Null
x_Code = Null: ox_Code = Null: z_Code = Null
x_Client = Null: ox_Client = Null: z_Client = Null
x_Range = Null: ox_Range = Null: z_Range = Null
x_Description = Null: ox_Description = Null: z_Description = Null
x_Color = Null: ox_Color = Null: z_Color = Null
x_Width = Null: ox_Width = Null: z_Width = Null
x_Height = Null: ox_Height = Null: z_Height = Null
x_Depth = Null: ox_Depth = Null: z_Depth = Null
x_Price = Null: ox_Price = Null: z_Price = Null
x_Edge = Null: ox_Edge = Null: z_Edge = Null
x_Sync_info = Null: ox_Sync_info = Null: z_Sync_info = Null
x_Sync_img = Null: ox_Sync_img = Null: z_Sync_img = Null
x_Info = Null: ox_Info = Null: z_Info = Null
%>
<%
Response.Buffer = True

' Load key from QueryString
x_Image = Request.QueryString("Image")

' Get action
sAction = Request.Form("a_edit")
If sAction = "" Or IsNull(sAction) Then
	sAction = "I"	' Display with input box
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
	x_Sync_info = Request.Form("x_Sync_info")
	x_Sync_img = Request.Form("x_Sync_img")
	x_Info = Request.Form("x_Info")
End If

' Check if valid key
If x_Image = "" Or IsNull(x_Image) Then Response.Redirect "Imageslist.asp"

' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case sAction
	Case "I": ' Get a record to display
		If Not LoadData() Then ' Load Record based on key
			Session(ewSessionMessage) = "No records found"
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "Imageslist.asp"
		End If
	Case "U": ' Update
		If EditData() Then ' Update Record based on key
			Session(ewSessionMessage) = "Update Record Successful"
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "Imageslist.asp"
		End If
End Select
%>
<!--#include file="header.asp"-->
<script type="text/javascript">
<!--
EW_LookupFn = "ewlookup.asp"; // ewlookup file name
EW_AddOptFn = "ewaddopt.asp"; // ewaddopt.asp file name
EW_MultiPagePage = "Page"; // multi-page Page Text
EW_MultiPageOf = "of"; // multi-page Of Text
//-->
</script>
<script type="text/javascript" src="ew.js"></script>
<script type="text/javascript">
<!--
EW_dateSep = "/"; // set date separator	
//-->
</script>
<script type="text/javascript">
<!--
function EW_checkMyForm(EW_this) {
if (EW_this.x_Range && !EW_hasValue(EW_this.x_Range, "SELECT" )) {
	if (!EW_onError(EW_this, EW_this.x_Range, "SELECT", "Please enter required field - Range"))
		return false;
}
if (EW_this.x_Description && !EW_hasValue(EW_this.x_Description, "SELECT" )) {
	if (!EW_onError(EW_this, EW_this.x_Description, "SELECT", "Please enter required field - Description"))
		return false;
}
if (EW_this.x_Color && !EW_hasValue(EW_this.x_Color, "SELECT" )) {
	if (!EW_onError(EW_this, EW_this.x_Color, "SELECT", "Please enter required field - Color"))
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
<script type="text/javascript">
<!--
	var EW_DHTMLEditors = [];
//-->
</script>
<p><span class="aspmaker">Edit TABLE: Images<br><br><a href="Imageslist.asp">Back to List</a></span></p>
<%
If Session(ewSessionMessage) <> "" Then
%>
<p><span class="ewmsg"><%= Session(ewSessionMessage) %></span></p>
<%
	Session(ewSessionMessage) = "" ' Clear message
End If
%>
<form name="fImagesedit" id="fImagesedit" action="Imagesedit.asp" method="post" onSubmit="return EW_checkMyForm(this);">
<p>
<input type="hidden" name="a_edit" value="U">
<table class="ewTable">
	<tr id="r_Image">
		<td class="ewTableHeader"><span>Image</span></td>
		<td class="ewTableAltRow"><span id="cb_x_Image">
<% sTmp = x_Image %><% If Not IsNull(sTmp) Then %><a href="Imagesview.asp?Image=<%= sTmp %>"><img src="<%= x_Image %>" Border=0></a><% Else %><img src="<%= x_Image %>" Border=0><% End If %><input type="hidden" id="x_Image" name="x_Image" value="<%= x_Image %>">
</span></td>
	</tr>
	<tr id="r_Code">
		<td class="ewTableHeader"><span>Code</span></td>
		<td class="ewTableAltRow"><span id="cb_x_Code">
<input type="text" name="x_Code" id="x_Code" size="30" maxlength="16" value="<%= Server.HTMLEncode(x_Code&"") %>">
</span></td>
	</tr>
	<tr id="r_Client">
		<td class="ewTableHeader"><span>Client</span></td>
		<td class="ewTableAltRow"><span id="cb_x_Client">
<input type="text" name="x_Client" id="x_Client" size="30" maxlength="16" value="<%= Server.HTMLEncode(x_Client&"") %>">
</span></td>
	</tr>
	<tr id="r_Range">
		<td class="ewTableHeader"><span>Range<span class='ewmsg'>&nbsp;*</span></span></td>
		<td class="ewTableAltRow"><span id="cb_x_Range">
<%
lst_x_Range = "<select id='x_Range' name='x_Range'>"
lst_x_Range = lst_x_Range & "<option value=''>Please Select</option>"
sSqlWrk = "SELECT [Range], [Name] FROM [Ranges]"
Set rswrk = Server.CreateObject("ADODB.Recordset")
rswrk.Open sSqlWrk, conn, 1, 2
If Not rswrk.Eof Then
	datawrk = rswrk.GetRows
	rowswrk = UBound(datawrk, 2)
	For rowcntwrk = 0 To rowswrk
		lst_x_Range = lst_x_Range & "<option value='" & datawrk(0, rowcntwrk) & "'"
		If CStr(datawrk(0, rowcntwrk)&"") = CStr(x_Range&"") Then
			lst_x_Range = lst_x_Range & " selected"
		End If
		lst_x_Range = lst_x_Range & ">" & datawrk(1, rowcntwrk) & "</option>"
	Next
End If
rswrk.Close
Set rswrk = Nothing
lst_x_Range = lst_x_Range & "</select>"
Response.Write lst_x_Range
%>
</span></td>
	</tr>
	<tr id="r_Description">
		<td class="ewTableHeader"><span>Description<span class='ewmsg'>&nbsp;*</span></span></td>
		<td class="ewTableAltRow"><span id="cb_x_Description">
<%
lst_x_Description = "<select id='x_Description' name='x_Description'>"
lst_x_Description = lst_x_Description & "<option value=''>Please Select</option>"
sSqlWrk = "SELECT [Description], [Name], '' FROM [Descriptions]"
If x_Description = "" Or IsNull(x_Description) Then
	sSqlWrk = sSqlWrk & " WHERE 0=1"
Else
	sSqlWrk = sSqlWrk & " WHERE [Description] = " & AdjustSql(x_Description) & ""
End If
Set rswrk = Server.CreateObject("ADODB.Recordset")
rswrk.Open sSqlWrk, conn, 1, 2
If Not rswrk.Eof Then
	datawrk = rswrk.GetRows
	rowswrk = UBound(datawrk, 2)
	For rowcntwrk = 0 To rowswrk
		lst_x_Description = lst_x_Description & "<option value='" & datawrk(0, rowcntwrk) & "'"
		If CStr(datawrk(0, rowcntwrk)&"") = CStr(x_Description&"") Then
			lst_x_Description = lst_x_Description & " selected"
		End If
		lst_x_Description = lst_x_Description & ">" & datawrk(1, rowcntwrk) & "</option>"
	Next
End If
rswrk.Close
Set rswrk = Nothing
lst_x_Description = lst_x_Description & "</select>"
Response.Write lst_x_Description
sSqlWrk = "SELECT [Description], [Name], '' FROM [Descriptions]"
sSqlWrk = EW_Encode(TEAencrypt(sSqlWrk, EW_RANDOM_KEY))
%>
<input type="hidden" name="s_x_Description" value="<%= sSqlWrk %>">
&nbsp;<a href="javascript:void(0);" onclick="EW_ShowAddOption('x_Description');">Add Description</a>
</span><span>
<div id="ao_x_Description" style="display: none;">
<input type="hidden" id="ltn_x_Description" value="Descriptions">
<input type="hidden" id="lfn_x_Description" value="Description">
<input type="hidden" id="dfn_x_Description" value="Name">
<input type="hidden" id="lfm_x_Description" value="Please enter required field - Description">
<input type="hidden" id="dfm_x_Description" value="Please enter required field - Name">
<input type="hidden" id="lfq_x_Description" value="">
<input type="hidden" id="dfq_x_Description" value="'">
<table class="ewAddOption">
<tr><td><span>Name</span></td><td><input type="text" id="df_x_Description" size="30"></td></tr>
<tr><td colspan="2" align="right"><input type="button" value="ADD" onClick="EW_PostNewOption('x_Description')"><input type="button" value="CANCEL" onClick="EW_HideAddOption('x_Description')"></td></tr>
</table>
</div>
</span></td>
	</tr>
	<tr id="r_Color">
		<td class="ewTableHeader"><span>Color<span class='ewmsg'>&nbsp;*</span></span></td>
		<td class="ewTableAltRow"><span id="cb_x_Color">
<%
lst_x_Color = "<select id='x_Color' name='x_Color'>"
lst_x_Color = lst_x_Color & "<option value=''>Please Select</option>"
sSqlWrk = "SELECT [Color], [Name] FROM [Colors]"
Set rswrk = Server.CreateObject("ADODB.Recordset")
rswrk.Open sSqlWrk, conn, 1, 2
If Not rswrk.Eof Then
	datawrk = rswrk.GetRows
	rowswrk = UBound(datawrk, 2)
	For rowcntwrk = 0 To rowswrk
		lst_x_Color = lst_x_Color & "<option value='" & datawrk(0, rowcntwrk) & "'"
		If CStr(datawrk(0, rowcntwrk)&"") = CStr(x_Color&"") Then
			lst_x_Color = lst_x_Color & " selected"
		End If
		lst_x_Color = lst_x_Color & ">" & datawrk(1, rowcntwrk) & "</option>"
	Next
End If
rswrk.Close
Set rswrk = Nothing
lst_x_Color = lst_x_Color & "</select>"
Response.Write lst_x_Color
%>
</span></td>
	</tr>
	<tr id="r_Width">
		<td class="ewTableHeader"><span>Width</span></td>
		<td class="ewTableAltRow"><span id="cb_x_Width">
<input type="text" name="x_Width" id="x_Width" size="30" value="<%= Server.HTMLEncode(x_Width&"") %>">
</span></td>
	</tr>
	<tr id="r_Height">
		<td class="ewTableHeader"><span>Height</span></td>
		<td class="ewTableAltRow"><span id="cb_x_Height">
<input type="text" name="x_Height" id="x_Height" size="30" value="<%= Server.HTMLEncode(x_Height&"") %>">
</span></td>
	</tr>
	<tr id="r_Depth">
		<td class="ewTableHeader"><span>Depth</span></td>
		<td class="ewTableAltRow"><span id="cb_x_Depth">
<input type="text" name="x_Depth" id="x_Depth" size="30" value="<%= Server.HTMLEncode(x_Depth&"") %>">
</span></td>
	</tr>
	<tr id="r_Price">
		<td class="ewTableHeader"><span>Price</span></td>
		<td class="ewTableAltRow"><span id="cb_x_Price">
<input type="text" name="x_Price" id="x_Price" size="30" value="<%= Server.HTMLEncode(x_Price&"") %>">
</span></td>
	</tr>
	<tr id="r_Edge">
		<td class="ewTableHeader"><span>Edge</span></td>
		<td class="ewTableAltRow"><span id="cb_x_Edge">
<input type="text" name="x_Edge" id="x_Edge" size="30" maxlength="25" value="<%= Server.HTMLEncode(x_Edge&"") %>">
</span></td>
	</tr>
	<tr id="r_Sync_info">
		<td class="ewTableHeader"><span>Sync info</span></td>
		<td class="ewTableAltRow"><span id="cb_x_Sync_info">
<input type="radio" name="x_Sync_info"<% If x_Sync_info = True Then %> checked<% End If %> value="Yes">
<%= "Yes" %>
<input type="radio" name="x_Sync_info"<% If x_Sync_info = False Then %> checked<% End If %> value="No">
<%= "No" %>
</span></td>
	</tr>
	<tr id="r_Sync_img">
		<td class="ewTableHeader"><span>Sync img</span></td>
		<td class="ewTableAltRow"><span id="cb_x_Sync_img">
<input type="radio" name="x_Sync_img"<% If x_Sync_img = True Then %> checked<% End If %> value="Yes">
<%= "Yes" %>
<input type="radio" name="x_Sync_img"<% If x_Sync_img = False Then %> checked<% End If %> value="No">
<%= "No" %>
</span></td>
	</tr>
	<tr id="r_Info">
		<td class="ewTableHeader"><span>Info</span></td>
		<td class="ewTableAltRow"><span id="cb_x_Info">
<textarea cols="35" rows="4" id="x_Info" name="x_Info"><%= x_Info %></textarea>
</span></td>
	</tr>
</table>
<p>
<input type="submit" name="btnAction" id="btnAction" value="EDIT">
</form>
<script language="JavaScript">
<!--
var f = document.fImagesedit;
EW_ajaxupdatecombo(f.x_Description, f.x_Description.options?f.x_Description.options[f.x_Description.selectedIndex].value:f.x_Description.value);
//-->
</script>
<!--#include file="footer.asp"-->
<%
conn.Close ' Close Connection
Set conn = Nothing
%>
<%

'-------------------------------------------------------------------------------
' Function LoadData
' - Load Data based on Key Value
' - Variables setup: field variables

Function LoadData()
	Dim rs, sSql, sFilter
	sFilter = ewSqlKeyWhere
	If Not IsNumeric(x_Image) Then
		LoadData = False
		Exit Function
	End If
	sFilter = Replace(sFilter, "@Image", AdjustSql(x_Image)) ' Replace key value
	sSql = ewBuildSql(ewSqlSelect, ewSqlWhere, ewSqlGroupBy, ewSqlHaving, ewSqlOrderBy, sFilter, "")
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
		x_Sync_info = rs("Sync_info")
		x_Sync_img = rs("Sync_img")
		x_Info = rs("Info")
	End If
	rs.Close
	Set rs = Nothing
End Function
%>
<%

'-------------------------------------------------------------------------------
' Function EditData
' - Edit Data based on Key Value
' - Variables used: field variables

Function EditData()
	On Error Resume Next
	Dim rs, sSql, sFilter
	sFilter = ewSqlKeyWhere
	If Not IsNumeric(x_Image) Then
		EditData = False
		Exit Function
	End If
	sFilter = Replace(sFilter, "@Image", AdjustSql(x_Image)) ' Replace key value
	sSql = ewBuildSql(ewSqlSelect, ewSqlWhere, ewSqlGroupBy, ewSqlHaving, ewSqlOrderBy, sFilter, "")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open sSql, conn, 1, 2
	If Err.Number <> 0 Then
		Session(ewSessionMessage) = Err.Description
		rs.Close
		Set rs = Nothing
		EditData = False
		Exit Function
	End If

	' clone old and new rs object
	Dim rsold, rsnew
	Set rsold = rs.Clone(1)
	rsold.Requery()
	Set rsnew = rs.Clone(1)
	If rs.Eof Then
		EditData = False ' Update Failed
	Else

		' Field Image
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

		' Field Sync_info
		sTmp = x_Sync_info
		If sTmp = "Yes" Then
			rs("Sync_info") = True
		Else
			rs("Sync_info") = False
		End If

		' Field Sync_img
		sTmp = x_Sync_img
		If sTmp = "Yes" Then
			rs("Sync_img") = True
		Else
			rs("Sync_img") = False
		End If

		' Field Info
		sTmp = Trim(x_Info)
		If Trim(sTmp) = "" Then sTmp = Null
		rs("Info") = sTmp

		' Call updating event
		If Recordset_Updating(rs, rsnew) Then
			rs.Update
			If Err.Number <> 0 Then
				Session(ewSessionMessage) = Err.Description
				EditData = False
			Else
				EditData = True
			End If
		Else
			rs.CancelUpdate
			EditData = False
		End If
	End If

	' Call updated event
	If EditData Then
		Call Recordset_Updated(rsold, rsnew)
	End If
	rs.Close
	Set rs = Nothing
	rsold.Close
	Set rsold = Nothing
	rsnew.Close
	Set rsnew = Nothing
End Function

'-------------------------------------------------------------------------------
' Recordset updating event

Function Recordset_Updating(rsold, rsnew)
	On Error Resume Next

	' Please enter your customized codes here
	Recordset_Updating = True
End Function

'-------------------------------------------------------------------------------
' Recordset updated event

Sub Recordset_Updated(rsold, rsnew)
	On Error Resume Next
End Sub
%>
