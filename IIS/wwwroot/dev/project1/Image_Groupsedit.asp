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

' Load key from QueryString
x_Image = Request.QueryString("Image")
x_Group = Request.QueryString("Group")

' Get action
sAction = Request.Form("a_edit")
If sAction = "" Or IsNull(sAction) Then
	sAction = "I"	' Display with input box
Else

	' Get fields from form
	x_Image = Request.Form("x_Image")
	x_Group = Request.Form("x_Group")
End If

' Check if valid key
If x_Image = "" Or IsNull(x_Image) Then Response.Redirect "Image_Groupslist.asp"
If x_Group = "" Or IsNull(x_Group) Then Response.Redirect "Image_Groupslist.asp"

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
	Case "U": ' Update
		If EditData() Then ' Update Record based on key
			Session(ewSessionMessage) = "Update Record Successful"
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
<script type="text/javascript">
<!--
EW_dateSep = "/"; // set date separator	
//-->
</script>
<script type="text/javascript">
<!--
function EW_checkMyForm(EW_this) {
if (EW_this.x_Group && !EW_hasValue(EW_this.x_Group, "SELECT" )) {
	if (!EW_onError(EW_this, EW_this.x_Group, "SELECT", "Please enter required field - Group"))
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
<p><span class="aspmaker">Edit TABLE: Image Groups<br><br><a href="Image_Groupslist.asp">Back to List</a></span></p>
<%
If Session(ewSessionMessage) <> "" Then
%>
<p><span class="ewmsg"><%= Session(ewSessionMessage) %></span></p>
<%
	Session(ewSessionMessage) = "" ' Clear message
End If
%>
<form name="fImage_Groupsedit" id="fImage_Groupsedit" action="Image_Groupsedit.asp" method="post" onSubmit="return EW_checkMyForm(this);">
<p>
<input type="hidden" name="a_edit" value="U">
<table class="ewTable">
	<input type="hidden" id="x_Image" name="x_Image" value="<%= x_Image %>">
	<tr id="r_Group">
		<td class="ewTableHeader"><span>Group<span class='ewmsg'>&nbsp;*</span></span></td>
		<td class="ewTableAltRow"><span id="cb_x_Group">
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
<input type="hidden" id="x_Group" name="x_Group" value="<%= x_Group %>">
</span></td>
	</tr>
</table>
<p>
<input type="submit" name="btnAction" id="btnAction" value="EDIT">
</form>
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
	If Not IsNumeric(x_Group) Then
		EditData = False
		Exit Function
	End If
	sFilter = Replace(sFilter, "@Group", AdjustSql(x_Group)) ' Replace key value
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

		' Field Group
		sTmp = x_Group
		If Not IsNumeric(sTmp) Then
			sTmp = Null
		Else
			sTmp = CLng(sTmp)
		End If
		rs("Group") = sTmp

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
