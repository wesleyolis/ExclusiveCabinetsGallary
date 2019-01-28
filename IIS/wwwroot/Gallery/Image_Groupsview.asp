<!--#include file="ewconfig.asp"-->
<!--#include file="db.asp"-->
<!--#include file="Image_Groupsinfo.asp"-->
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
x_Group = Null: ox_Group = Null: z_Group = Null
%>
<%
Response.Buffer = True
x_Image = Request.QueryString("Image")
If x_Image = "" Or IsNull(x_Image) Then Response.Redirect "Image_Groupslist.asp"
x_Group = Request.QueryString("Group")
If x_Group = "" Or IsNull(x_Group) Then Response.Redirect "Image_Groupslist.asp"

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
			Response.Redirect "Image_Groupslist.asp"
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
<p><span class="aspmaker">View TABLE: Image Groups<br><br>
<a href="Image_Groupslist.asp">Back to List</a>&nbsp;
<a href="<% If Not IsNull(x_Image) AND Not IsNull(x_Group) Then Response.Write "Image_Groupsedit.asp?Image=" & Server.URLEncode(x_Image) & "&Group=" & Server.URLEncode(x_Group) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Edit</a>&nbsp;
<a href="<% If Not IsNull(x_Image) AND Not IsNull(x_Group) Then Response.Write "Image_Groupsadd.asp?Image=" & Server.URLEncode(x_Image) & "&Group=" & Server.URLEncode(x_Group) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Copy</a>&nbsp;
<a href="<% If Not IsNull(x_Image) AND Not IsNull(x_Group) Then Response.Write "Image_Groupsdelete.asp?Image=" & Server.URLEncode(x_Image) & "&Group=" & Server.URLEncode(x_Group) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Delete</a>&nbsp;
</span></p>
<p>
<form>
<table class="ewTable">
	<tr>
		<td class="ewTableHeader"><span>Group</span></td>
		<td class="ewTableAltRow"><span>
<%
If Not IsNull(x_Group) Then
	sSqlWrk = "SELECT [Name] FROM [Color_Groups]"
	sTmp = x_Group
	sSqlWrk = sSqlWrk & " WHERE [Index] = " & AdjustSql(sTmp) & ""
	sSqlWrk = sSqlWrk & " ORDER BY [Name] Asc"
	Set rswrk = conn.Execute(sSqlWrk)
	If Not rswrk.Eof Then
		sTmp = rswrk("Name")
	End If
	rswrk.Close
	Set rswrk = Nothing
Else
	sTmp = Null
End If
ox_Group = x_Group ' Backup Original Value
x_Group = sTmp
%>
<% Response.Write x_Group %>
<% x_Group = ox_Group ' Restore Original Value %>
</span></td>
	</tr>
</table>
</form>
<p>
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
	If Not IsNumeric(x_Group) Then
		LoadData = False
		Exit Function
	End If
	sFilter = Replace(sFilter, "@Group", AdjustSql(x_Group)) ' Replace key value
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
		x_Group = rs("Group")
	End If
	rs.Close
	Set rs = Nothing
End Function
%>
