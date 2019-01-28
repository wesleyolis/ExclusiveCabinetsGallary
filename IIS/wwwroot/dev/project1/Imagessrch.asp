<!--#include file="ewconfig.asp"-->
<!--#include file="db.asp"-->
<!--#include file="Imagesinfo.asp"-->
<!--#include file="advsecu.asp"-->
<!--#include file="aspmkrfn.asp"-->
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

' Get action
sAction = Request.Form("a_search")
Select Case sAction
	Case "S": ' Get Search Criteria

	' Build search string for advanced search, remove blank field
	sSrchStr = ""

	' Field Image
	x_Image = Request.Form("x_Image")
	z_Image = Request.Form("z_Image")
	sSrchWrk = ""
	If x_Image <> "" Then
		sSrchWrk = sSrchWrk & "x_Image=" & Server.URLEncode(x_Image)
		sSrchWrk = sSrchWrk & "&z_Image=" & Server.URLEncode(z_Image)
	End If
	If sSrchWrk <> "" Then
		If sSrchStr <> "" Then sSrchStr = sSrchStr & "&"
		sSrchStr = sSrchStr & sSrchWrk
	End If

	' Field Dir
	x_Dir = Request.Form("x_Dir")
	z_Dir = Request.Form("z_Dir")
	sSrchWrk = ""
	If x_Dir <> "" Then
		sSrchWrk = sSrchWrk & "x_Dir=" & Server.URLEncode(x_Dir)
		sSrchWrk = sSrchWrk & "&z_Dir=" & Server.URLEncode(z_Dir)
	End If
	If sSrchWrk <> "" Then
		If sSrchStr <> "" Then sSrchStr = sSrchStr & "&"
		sSrchStr = sSrchStr & sSrchWrk
	End If

	' Field Code
	x_Code = Request.Form("x_Code")
	z_Code = Request.Form("z_Code")
	sSrchWrk = ""
	If x_Code <> "" Then
		sSrchWrk = sSrchWrk & "x_Code=" & Server.URLEncode(x_Code)
		sSrchWrk = sSrchWrk & "&z_Code=" & Server.URLEncode(z_Code)
	End If
	If sSrchWrk <> "" Then
		If sSrchStr <> "" Then sSrchStr = sSrchStr & "&"
		sSrchStr = sSrchStr & sSrchWrk
	End If

	' Field Client
	x_Client = Request.Form("x_Client")
	z_Client = Request.Form("z_Client")
	sSrchWrk = ""
	If x_Client <> "" Then
		sSrchWrk = sSrchWrk & "x_Client=" & Server.URLEncode(x_Client)
		sSrchWrk = sSrchWrk & "&z_Client=" & Server.URLEncode(z_Client)
	End If
	If sSrchWrk <> "" Then
		If sSrchStr <> "" Then sSrchStr = sSrchStr & "&"
		sSrchStr = sSrchStr & sSrchWrk
	End If

	' Field Range
	x_Range = Request.Form("x_Range")
	z_Range = Request.Form("z_Range")
	sSrchWrk = ""
	If x_Range <> "" Then
		sSrchWrk = sSrchWrk & "x_Range=" & Server.URLEncode(x_Range)
		sSrchWrk = sSrchWrk & "&z_Range=" & Server.URLEncode(z_Range)
	End If
	If sSrchWrk <> "" Then
		If sSrchStr <> "" Then sSrchStr = sSrchStr & "&"
		sSrchStr = sSrchStr & sSrchWrk
	End If

	' Field Description
	x_Description = Request.Form("x_Description")
	z_Description = Request.Form("z_Description")
	sSrchWrk = ""
	If x_Description <> "" Then
		sSrchWrk = sSrchWrk & "x_Description=" & Server.URLEncode(x_Description)
		sSrchWrk = sSrchWrk & "&z_Description=" & Server.URLEncode(z_Description)
	End If
	If sSrchWrk <> "" Then
		If sSrchStr <> "" Then sSrchStr = sSrchStr & "&"
		sSrchStr = sSrchStr & sSrchWrk
	End If

	' Field Color
	x_Color = Request.Form("x_Color")
	z_Color = Request.Form("z_Color")
	sSrchWrk = ""
	If x_Color <> "" Then
		sSrchWrk = sSrchWrk & "x_Color=" & Server.URLEncode(x_Color)
		sSrchWrk = sSrchWrk & "&z_Color=" & Server.URLEncode(z_Color)
	End If
	If sSrchWrk <> "" Then
		If sSrchStr <> "" Then sSrchStr = sSrchStr & "&"
		sSrchStr = sSrchStr & sSrchWrk
	End If

	' Field Width
	x_Width = Request.Form("x_Width")
	z_Width = Request.Form("z_Width")
	sSrchWrk = ""
	If x_Width <> "" Then
		sSrchWrk = sSrchWrk & "x_Width=" & Server.URLEncode(x_Width)
		sSrchWrk = sSrchWrk & "&z_Width=" & Server.URLEncode(z_Width)
	End If
	If sSrchWrk <> "" Then
		If sSrchStr <> "" Then sSrchStr = sSrchStr & "&"
		sSrchStr = sSrchStr & sSrchWrk
	End If

	' Field Height
	x_Height = Request.Form("x_Height")
	z_Height = Request.Form("z_Height")
	sSrchWrk = ""
	If x_Height <> "" Then
		sSrchWrk = sSrchWrk & "x_Height=" & Server.URLEncode(x_Height)
		sSrchWrk = sSrchWrk & "&z_Height=" & Server.URLEncode(z_Height)
	End If
	If sSrchWrk <> "" Then
		If sSrchStr <> "" Then sSrchStr = sSrchStr & "&"
		sSrchStr = sSrchStr & sSrchWrk
	End If

	' Field Depth
	x_Depth = Request.Form("x_Depth")
	z_Depth = Request.Form("z_Depth")
	sSrchWrk = ""
	If x_Depth <> "" Then
		sSrchWrk = sSrchWrk & "x_Depth=" & Server.URLEncode(x_Depth)
		sSrchWrk = sSrchWrk & "&z_Depth=" & Server.URLEncode(z_Depth)
	End If
	If sSrchWrk <> "" Then
		If sSrchStr <> "" Then sSrchStr = sSrchStr & "&"
		sSrchStr = sSrchStr & sSrchWrk
	End If

	' Field Price
	x_Price = Request.Form("x_Price")
	z_Price = Request.Form("z_Price")
	sSrchWrk = ""
	If x_Price <> "" Then
		sSrchWrk = sSrchWrk & "x_Price=" & Server.URLEncode(x_Price)
		sSrchWrk = sSrchWrk & "&z_Price=" & Server.URLEncode(z_Price)
	End If
	If sSrchWrk <> "" Then
		If sSrchStr <> "" Then sSrchStr = sSrchStr & "&"
		sSrchStr = sSrchStr & sSrchWrk
	End If

	' Field Edge
	x_Edge = Request.Form("x_Edge")
	z_Edge = Request.Form("z_Edge")
	sSrchWrk = ""
	If x_Edge <> "" Then
		sSrchWrk = sSrchWrk & "x_Edge=" & Server.URLEncode(x_Edge)
		sSrchWrk = sSrchWrk & "&z_Edge=" & Server.URLEncode(z_Edge)
	End If
	If sSrchWrk <> "" Then
		If sSrchStr <> "" Then sSrchStr = sSrchStr & "&"
		sSrchStr = sSrchStr & sSrchWrk
	End If

	' Field Sync_info
	x_Sync_info = Request.Form("x_Sync_info")
	z_Sync_info = Request.Form("z_Sync_info")
	sSrchWrk = ""
	If x_Sync_info <> "" Then
		sSrchWrk = sSrchWrk & "x_Sync_info=" & Server.URLEncode(x_Sync_info)
		sSrchWrk = sSrchWrk & "&z_Sync_info=" & Server.URLEncode(z_Sync_info)
	End If
	If sSrchWrk <> "" Then
		If sSrchStr <> "" Then sSrchStr = sSrchStr & "&"
		sSrchStr = sSrchStr & sSrchWrk
	End If

	' Field Sync_img
	x_Sync_img = Request.Form("x_Sync_img")
	z_Sync_img = Request.Form("z_Sync_img")
	sSrchWrk = ""
	If x_Sync_img <> "" Then
		sSrchWrk = sSrchWrk & "x_Sync_img=" & Server.URLEncode(x_Sync_img)
		sSrchWrk = sSrchWrk & "&z_Sync_img=" & Server.URLEncode(z_Sync_img)
	End If
	If sSrchWrk <> "" Then
		If sSrchStr <> "" Then sSrchStr = sSrchStr & "&"
		sSrchStr = sSrchStr & sSrchWrk
	End If

	' Field Info
	x_Info = Request.Form("x_Info")
	z_Info = Request.Form("z_Info")
	sSrchWrk = ""
	If x_Info <> "" Then
		sSrchWrk = sSrchWrk & "x_Info=" & Server.URLEncode(x_Info)
		sSrchWrk = sSrchWrk & "&z_Info=" & Server.URLEncode(z_Info)
	End If
	If sSrchWrk <> "" Then
		If sSrchStr <> "" Then sSrchStr = sSrchStr & "&"
		sSrchStr = sSrchStr & sSrchWrk
	End If
	If sSrchStr <> "" Then
		Response.Clear
		Response.Redirect "Imageslist.asp" & "?" & sSrchStr
	End If
	Case Else ' Restore search settings
		x_Image = Session(ewSessionTblAdvSrch & "_x_Image")
		x_Dir = Session(ewSessionTblAdvSrch & "_x_Dir")
		x_Code = Session(ewSessionTblAdvSrch & "_x_Code")
		x_Client = Session(ewSessionTblAdvSrch & "_x_Client")
		x_Range = Session(ewSessionTblAdvSrch & "_x_Range")
		x_Description = Session(ewSessionTblAdvSrch & "_x_Description")
		x_Color = Session(ewSessionTblAdvSrch & "_x_Color")
		x_Width = Session(ewSessionTblAdvSrch & "_x_Width")
		x_Height = Session(ewSessionTblAdvSrch & "_x_Height")
		x_Depth = Session(ewSessionTblAdvSrch & "_x_Depth")
		x_Price = Session(ewSessionTblAdvSrch & "_x_Price")
		x_Edge = Session(ewSessionTblAdvSrch & "_x_Edge")
		x_Sync_info = Session(ewSessionTblAdvSrch & "_x_Sync_info")
		x_Sync_img = Session(ewSessionTblAdvSrch & "_x_Sync_img")
		x_Info = Session(ewSessionTblAdvSrch & "_x_Info")
End Select

' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
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
if (EW_this.x_Image && !EW_checkinteger(EW_this.x_Image.value)) {
	if (!EW_onError(EW_this, EW_this.x_Image, "NO", "Incorrect integer - Image"))
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
<p><span class="aspmaker">Search TABLE: Images<br><br><a href="Imageslist.asp">Back to List</a></span></p>
<form name="fImagessearch" id="fImagessearch" action="Imagessrch.asp" method="post" onSubmit="return EW_checkMyForm(this);">
<p>
<input type="hidden" name="a_search" value="S">
<table class="ewTable">
	<tr>
		<td class="ewTableHeader"><span>Image</span></td>
		<td class="ewTableAltRow"><span class="aspmaker">=<input type="hidden" name="z_Image" value="=,,"></span></td>
		<td class="ewTableAltRow"><span class="aspmaker">
<input type="text" name="x_Image" value="<%= x_Image %>">
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Dir</span></td>
		<td class="ewTableAltRow"><span class="aspmaker">=<input type="hidden" name="z_Dir" value="=,,"></span></td>
		<td class="ewTableAltRow"><span class="aspmaker">
<input type="text" id="x_Dir" name="x_Dir" value="<%= x_Dir %>">
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Code</span></td>
		<td class="ewTableAltRow"><span class="aspmaker">contains<input type="hidden" name="z_Code" value="LIKE,'%,%'"></span></td>
		<td class="ewTableAltRow"><span class="aspmaker">
<input type="text" name="x_Code" id="x_Code" size="30" maxlength="16" value="<%= Server.HTMLEncode(x_Code&"") %>">
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Client</span></td>
		<td class="ewTableAltRow"><span class="aspmaker">contains<input type="hidden" name="z_Client" value="LIKE,'%,%'"></span></td>
		<td class="ewTableAltRow"><span class="aspmaker">
<input type="text" name="x_Client" id="x_Client" size="30" maxlength="16" value="<%= Server.HTMLEncode(x_Client&"") %>">
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Range</span></td>
		<td class="ewTableAltRow"><span class="aspmaker">=<input type="hidden" name="z_Range" value="=,,"></span></td>
		<td class="ewTableAltRow"><span class="aspmaker">
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
	<tr>
		<td class="ewTableHeader"><span>Description</span></td>
		<td class="ewTableAltRow"><span class="aspmaker">=<input type="hidden" name="z_Description" value="=,,"></span></td>
		<td class="ewTableAltRow"><span class="aspmaker">
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
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Color</span></td>
		<td class="ewTableAltRow"><span class="aspmaker">=<input type="hidden" name="z_Color" value="=,,"></span></td>
		<td class="ewTableAltRow"><span class="aspmaker">
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
	<tr>
		<td class="ewTableHeader"><span>Width</span></td>
		<td class="ewTableAltRow"><span class="aspmaker">=<input type="hidden" name="z_Width" value="=,,"></span></td>
		<td class="ewTableAltRow"><span class="aspmaker">
<input type="text" name="x_Width" id="x_Width" size="30" value="<%= Server.HTMLEncode(x_Width&"") %>">
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Height</span></td>
		<td class="ewTableAltRow"><span class="aspmaker">=<input type="hidden" name="z_Height" value="=,,"></span></td>
		<td class="ewTableAltRow"><span class="aspmaker">
<input type="text" name="x_Height" id="x_Height" size="30" value="<%= Server.HTMLEncode(x_Height&"") %>">
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Depth</span></td>
		<td class="ewTableAltRow"><span class="aspmaker">=<input type="hidden" name="z_Depth" value="=,,"></span></td>
		<td class="ewTableAltRow"><span class="aspmaker">
<input type="text" name="x_Depth" id="x_Depth" size="30" value="<%= Server.HTMLEncode(x_Depth&"") %>">
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Price</span></td>
		<td class="ewTableAltRow"><span class="aspmaker">=<input type="hidden" name="z_Price" value="=,,"></span></td>
		<td class="ewTableAltRow"><span class="aspmaker">
<input type="text" name="x_Price" id="x_Price" size="30" value="<%= Server.HTMLEncode(x_Price&"") %>">
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Edge</span></td>
		<td class="ewTableAltRow"><span class="aspmaker">contains<input type="hidden" name="z_Edge" value="LIKE,'%,%'"></span></td>
		<td class="ewTableAltRow"><span class="aspmaker">
<input type="text" name="x_Edge" id="x_Edge" size="30" maxlength="25" value="<%= Server.HTMLEncode(x_Edge&"") %>">
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Sync info</span></td>
		<td class="ewTableAltRow"><span class="aspmaker">=<input type="hidden" name="z_Sync_info" value="=,,"></span></td>
		<td class="ewTableAltRow"><span class="aspmaker">
<input type="radio" name="x_Sync_info" value="Yes">
<%= "Yes" %>
<input type="radio" name="x_Sync_info" value="No">
<%= "No" %>
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Sync img</span></td>
		<td class="ewTableAltRow"><span class="aspmaker">=<input type="hidden" name="z_Sync_img" value="=,,"></span></td>
		<td class="ewTableAltRow"><span class="aspmaker">
<input type="radio" name="x_Sync_img" value="Yes">
<%= "Yes" %>
<input type="radio" name="x_Sync_img" value="No">
<%= "No" %>
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Info</span></td>
		<td class="ewTableAltRow"><span class="aspmaker">contains<input type="hidden" name="z_Info" value="LIKE,'%,%'"></span></td>
		<td class="ewTableAltRow"><span class="aspmaker">
<textarea cols="35" rows="4" id="x_Info" name="x_Info"><%= x_Info %></textarea>
</span></td>
	</tr>
</table>
<p>
<input type="submit" name="Action" value="Search">
<input type="submit" name="Reset" value="Reset" onclick="clearForm(this.form);">
</form>
<script language="JavaScript">
<!--
var f = document.fImagessearch;
EW_ajaxupdatecombo(f.x_Description, f.x_Description.options?f.x_Description.options[f.x_Description.selectedIndex].value:f.x_Description.value);
//-->
</script>
<!--#include file="footer.asp"-->
<%
conn.Close ' Close Connection
Set conn = Nothing
%>
