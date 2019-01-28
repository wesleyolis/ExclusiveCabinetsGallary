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
x_Image = Request.QueryString("Image")
If x_Image = "" Or IsNull(x_Image) Then Response.Redirect "Imageslist.asp"

' Get action
sAction = Request.Form("a_view")
If sAction = "" Or IsNull(sAction) Then
	sAction = "I"	' Display with input box
End If

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
<p><span class="aspmaker">View TABLE: Images<br><br>
<a href="Imageslist.asp">Back to List</a>&nbsp;
<a href="<% If Not IsNull(x_Image) Then Response.Write "Imagesedit.asp?Image=" & Server.URLEncode(x_Image) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Edit</a>&nbsp;
<a href="<% If Not IsNull(x_Image) Then Response.Write "Imagesadd.asp?Image=" & Server.URLEncode(x_Image) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Copy</a>&nbsp;
<a href="<% If Not IsNull(x_Image) Then Response.Write "Imagesdelete.asp?Image=" & Server.URLEncode(x_Image) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Delete</a>&nbsp;
<a href="Image_Groupslist.asp?showmaster=1&Image=<%= Server.URLEncode(x_Image)%>">Image Groups Details</a>&nbsp;
</span></p>
<p>
<form>
<table class="ewTable">
	<tr>
		<td class="ewTableHeader"><span>Image</span></td>
		<td class="ewTableAltRow"><span>
<% sTmp = x_Image %><% If Not IsNull(sTmp) Then %><a href="getimage.asp?I=<%= sTmp %>"><img src="thumbs/<%= x_Image %>.jpg" Border=0></a><% Else %><img src="<%= x_Image %>" Border=0><% End If %>
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Code</span></td>
		<td class="ewTableAltRow"><span>
<% Response.Write x_Code %>
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Client</span></td>
		<td class="ewTableAltRow"><span>
<% Response.Write x_Client %>
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Range</span></td>
		<td class="ewTableAltRow"><span>
<%
If Not IsNull(x_Range) Then
	sSqlWrk = "SELECT [Name] FROM [Ranges]"
	sTmp = x_Range
	sSqlWrk = sSqlWrk & " WHERE [Range] = " & AdjustSql(sTmp) & ""
	Set rswrk = conn.Execute(sSqlWrk)
	If Not rswrk.Eof Then
		sTmp = rswrk("Name")
	End If
	rswrk.Close
	Set rswrk = Nothing
Else
	sTmp = Null
End If
ox_Range = x_Range ' Backup Original Value
x_Range = sTmp
%>
<% sTmp = x_Range %><% If Not IsNull(sTmp) Then %><a href="Imageslist.asp?x_Range<%= sTmp %>"><% Response.Write x_Range %></a><% Else %><% Response.Write x_Range %><% End If %>
<% x_Range = ox_Range ' Restore Original Value %>
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Description</span></td>
		<td class="ewTableAltRow"><span>
<%
If Not IsNull(x_Description) Then
	sSqlWrk = "SELECT [Name] FROM [Descriptions]"
	sTmp = x_Description
	sSqlWrk = sSqlWrk & " WHERE [Description] = " & AdjustSql(sTmp) & ""
	Set rswrk = conn.Execute(sSqlWrk)
	If Not rswrk.Eof Then
		sTmp = rswrk("Name")
	End If
	rswrk.Close
	Set rswrk = Nothing
Else
	sTmp = Null
End If
ox_Description = x_Description ' Backup Original Value
x_Description = sTmp
%>
<% sTmp = x_Description %><% If Not IsNull(sTmp) Then %><a href="Imageslist.asp?x_Description<%= sTmp %>"><% Response.Write x_Description %></a><% Else %><% Response.Write x_Description %><% End If %>
<% x_Description = ox_Description ' Restore Original Value %>
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Color</span></td>
		<td class="ewTableAltRow"><span>
<%
If Not IsNull(x_Color) Then
	sSqlWrk = "SELECT [Name] FROM [Colors]"
	sTmp = x_Color
	sSqlWrk = sSqlWrk & " WHERE [Color] = " & AdjustSql(sTmp) & ""
	Set rswrk = conn.Execute(sSqlWrk)
	If Not rswrk.Eof Then
		sTmp = rswrk("Name")
	End If
	rswrk.Close
	Set rswrk = Nothing
Else
	sTmp = Null
End If
ox_Color = x_Color ' Backup Original Value
x_Color = sTmp
%>
<% Response.Write x_Color %>
<% x_Color = ox_Color ' Restore Original Value %>
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Width</span></td>
		<td class="ewTableAltRow"><span>
<% Response.Write x_Width %>
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Height</span></td>
		<td class="ewTableAltRow"><span>
<% Response.Write x_Height %>
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Depth</span></td>
		<td class="ewTableAltRow"><span>
<% Response.Write x_Depth %>
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Price</span></td>
		<td class="ewTableAltRow"><span>
<% Response.Write x_Price %>
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Edge</span></td>
		<td class="ewTableAltRow"><span>
<% Response.Write x_Edge %>
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Info</span></td>
		<td class="ewTableAltRow"><span>
<%= Replace(x_Info&"", vbLf, "<br>") %>
</span></td>
	</tr>
</table>
</form>
<p>
<!--#include file="footer.asp"-->
<%
conn.Close ' Close Connection
Set conn = Nothing
%><%

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