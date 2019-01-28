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
bCopy = True
x_Image = Request.QueryString("Image")
If x_Image = "" Or IsNull(x_Image) Then
	bCopy = False
End If

' Get action
sAction = Request.Form("a_add")
If (sAction = "" Or IsNull(sAction)) Then
	If bCopy Then
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
	x_Sync_info = Request.Form("x_Sync_info")
	x_Sync_img = Request.Form("x_Sync_img")
	x_Info = Request.Form("x_Info")
End If

' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case sAction
	Case "C": ' Get a record to display
		If Not LoadData() Then ' Load Record based on key
			Session(ewSessionMessage) = "No records found"
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "Imageslist.asp"
		End If
	Case "A": ' Add
		If AddData() Then ' Add New Record
			Session(ewSessionMessage) = "Add New Record Successful"
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "Imageslist.asp"
		Else
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
<p><span class="aspmaker">Add to TABLE: Images<br><br><a href="Imageslist.asp">Back to List</a></span></p>
<%
If Session(ewSessionMessage) <> "" Then
%>
<p><span class="ewmsg"><%= Session(ewSessionMessage) %></span></p>
<%
	Session(ewSessionMessage) = "" ' Clear message
End If
%>
<form name="fImagesadd" id="fImagesadd" action="Imagesadd.asp" method="post" onSubmit="return EW_checkMyForm(this);">
<p>
<input type="hidden" name="a_add" value="A">
<%
If Session(ewSessionMessage) <> "" Then
%>
<p><span class="ewmsg"><%= Session(ewSessionMessage) %></span></p>
<%
	Session(ewSessionMessage) = "" ' Clear message
End If
%>
<table class="ewTable">
<% x_Dir = 0 ' Set default value %>
<input type="hidden" id="x_Dir" name="x_Dir" value="<%= x_Dir %>">
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
<% If IsNull(x_Range) or x_Range = "" Then x_Range = 0 ' Set default value %>
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
<% If IsNull(x_Description) or x_Description = "" Then x_Description = 0 ' Set default value %>
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
<% If IsNull(x_Color) or x_Color = "" Then x_Color = 0 ' Set default value %>
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
<% If IsNull(x_Width) or x_Width = "" Then x_Width = 0 ' Set default value %>
<input type="text" name="x_Width" id="x_Width" size="30" value="<%= Server.HTMLEncode(x_Width&"") %>">
</span></td>
	</tr>
	<tr id="r_Height">
		<td class="ewTableHeader"><span>Height</span></td>
		<td class="ewTableAltRow"><span id="cb_x_Height">
<% If IsNull(x_Height) or x_Height = "" Then x_Height = 0 ' Set default value %>
<input type="text" name="x_Height" id="x_Height" size="30" value="<%= Server.HTMLEncode(x_Height&"") %>">
</span></td>
	</tr>
	<tr id="r_Depth">
		<td class="ewTableHeader"><span>Depth</span></td>
		<td class="ewTableAltRow"><span id="cb_x_Depth">
<% If IsNull(x_Depth) or x_Depth = "" Then x_Depth = 0 ' Set default value %>
<input type="text" name="x_Depth" id="x_Depth" size="30" value="<%= Server.HTMLEncode(x_Depth&"") %>">
</span></td>
	</tr>
	<tr id="r_Price">
		<td class="ewTableHeader"><span>Price</span></td>
		<td class="ewTableAltRow"><span id="cb_x_Price">
<% If IsNull(x_Price) or x_Price = "" Then x_Price = 0 ' Set default value %>
<input type="text" name="x_Price" id="x_Price" size="30" value="<%= Server.HTMLEncode(x_Price&"") %>">
</span></td>
	</tr>
	<tr id="r_Edge">
		<td class="ewTableHeader"><span>Edge</span></td>
		<td class="ewTableAltRow"><span id="cb_x_Edge">
<input type="text" name="x_Edge" id="x_Edge" size="30" maxlength="25" value="<%= Server.HTMLEncode(x_Edge&"") %>">
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
<input type="submit" name="btnAction" id="btnAction" value="ADD">
</form>
<script language="JavaScript">
<!--
var f = document.fImagesadd;
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
' Function AddData
' - Add Data
' - Variables used: field variables

Function AddData()
	On Error Resume Next
	Dim rs, sSql, sFilter
	Dim bCheckKey, sSqlChk, sWhereChk
	sFilter = ewSqlKeyWhere

	' Check for duplicate key
	bCheckKey = True
	If x_Image = "" Or IsNull(x_Image) Then
		bCheckKey = False
	Else
		sFilter = Replace(sFilter, "@Image", AdjustSql(x_Image)) ' Replace key value
	End If
	If Not IsNumeric(x_Image) Then
		bCheckKey = False
	End If
	If bCheckKey Then
		sSqlChk = ewBuildSql(ewSqlSelect, ewSqlWhere, ewSqlGroupBy, ewSqlHaving, ewSqlOrderBy, sFilter, "")
		Set rsChk = conn.Execute(sSqlChk)
		If Err.Number <> 0 Then
			Session(ewSessionMessage) = Err.Description
			rsChk.Close
			Set rsChk = Nothing
			AddData = False
			Exit Function
		ElseIf Not rsChk.Eof Then
			Session(ewSessionMessage) = "Duplicate value for primary key"
			rsChk.Close
			Set rsChk = Nothing
			AddData = False
			Exit Function
		End If
		rsChk.Close
		Set rsChk = Nothing
	End If

	' Add New Record
	sFilter = "(0 = 1)"
	sSql = ewBuildSql(ewSqlSelect, ewSqlWhere, ewSqlGroupBy, ewSqlHaving, ewSqlOrderBy, sFilter, "")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open sSql, conn, 1, 2
	If Err.Number <> 0 Then
		Session(ewSessionMessage) = Err.Description
		rs.Close
		Set rs = Nothing
		AddData = False
		Exit Function
	End If

	' Clone new rs object
	Dim rsnew
	Set rsnew = rs.Clone(1)
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
	sTmp = Trim(x_Info)
	If Trim(sTmp) = "" Then sTmp = Null
	rs("Info") = sTmp

	' Call recordset inserting event
	If Recordset_Inserting(rs) Then
		rs.Update
		If Err.Number <> 0 Then
			Session(ewSessionMessage) = Err.Description
			AddData = False
		Else
			AddData = True
		End If
	Else
		rs.CancelUpdate
		AddData = False
	End If
	rs.Close
	Set rs = Nothing

	' Call recordset inserted event
	If AddData Then
		Call Recordset_Inserted(rsnew)
	End If
	rsnew.Close
	Set rsnew = Nothing
End Function

'-------------------------------------------------------------------------------
' Recordset inserting event

Function Recordset_Inserting(rsnew)
	On Error Resume Next

	' Please enter your customized codes here
	Recordset_Inserting = True
End Function

'-------------------------------------------------------------------------------
' Recordset inserted event

Sub Recordset_Inserted(rsnew)
	On Error Resume Next

	' Get key value
	Dim sKey
	sKey = ""
	If sKey <> "" Then sKey = sKey & ","
	sKey = sKey & rsnew.Fields("Image")
	x_Image = rsnew.Fields("Image")
End Sub
%>
