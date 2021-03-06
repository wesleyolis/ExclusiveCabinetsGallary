<!--#include file="ewconfig.asp"-->
<!--#include file="db.asp"-->
<!--#include file="Color_Groupsinfo.asp"-->
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
x_Index = Null: ox_Index = Null: z_Index = Null
x_Name = Null: ox_Name = Null: z_Name = Null
x_Memo = Null: ox_Memo = Null: z_Memo = Null
%>
<%
Response.Buffer = True

' Load key from QueryString
x_Index = Request.QueryString("Index")

' Get action
sAction = Request.Form("a_edit")
If sAction = "" Or IsNull(sAction) Then
	sAction = "I"	' Display with input box
Else

	' Get fields from form
	x_Index = Request.Form("x_Index")
	x_Name = Request.Form("x_Name")
	x_Memo = Request.Form("x_Memo")
End If

' Check if valid key
If x_Index = "" Or IsNull(x_Index) Then Response.Redirect "Color_Groupslist.asp"

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
			Response.Redirect "Color_Groupslist.asp"
		End If
	Case "U": ' Update
		If EditData() Then ' Update Record based on key
			Session(ewSessionMessage) = "Update Record Successful"
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "Color_Groupslist.asp"
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
if (EW_this.x_Name && !EW_hasValue(EW_this.x_Name, "TEXT" )) {
	if (!EW_onError(EW_this, EW_this.x_Name, "TEXT", "Please enter required field - Name"))
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
<p><span class="aspmaker">Edit TABLE: Color Groups<br><br><a href="Color_Groupslist.asp">Back to List</a></span></p>
<%
If Session(ewSessionMessage) <> "" Then
%>
<p><span class="ewmsg"><%= Session(ewSessionMessage) %></span></p>
<%
	Session(ewSessionMessage) = "" ' Clear message
End If
%>
<form name="fColor_Groupsedit" id="fColor_Groupsedit" action="Color_Groupsedit.asp" method="post" onSubmit="return EW_checkMyForm(this);">
<p>
<input type="hidden" name="a_edit" value="U">
<table class="ewTable">
<input type="hidden" id="x_Index" name="x_Index" value="<%= x_Index %>">
	<tr id="r_Name">
		<td class="ewTableHeader"><span>Name<span class='ewmsg'>&nbsp;*</span></span></td>
		<td class="ewTableAltRow"><span id="cb_x_Name">
<input type="text" name="x_Name" id="x_Name" size="30" maxlength="25" value="<%= Server.HTMLEncode(x_Name&"") %>">
</span></td>
	</tr>
	<tr id="r_Memo">
		<td class="ewTableHeader"><span>Memo</span></td>
		<td class="ewTableAltRow"><span id="cb_x_Memo">
<textarea cols="35" rows="4" id="x_Memo" name="x_Memo"><%= x_Memo %></textarea>
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
	If Not IsNumeric(x_Index) Then
		LoadData = False
		Exit Function
	End If
	sFilter = Replace(sFilter, "@Index", AdjustSql(x_Index)) ' Replace key value
	sSql = ewBuildSql(ewSqlSelect, ewSqlWhere, ewSqlGroupBy, ewSqlHaving, ewSqlOrderBy, sFilter, "")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sSql, conn
	If rs.Eof Then
		LoadData = False
	Else
		LoadData = True
		rs.MoveFirst

		' Get the field contents
		x_Index = rs("Index")
		x_Name = rs("Name")
		x_Memo = rs("Memo")
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
	If Not IsNumeric(x_Index) Then
		EditData = False
		Exit Function
	End If
	sFilter = Replace(sFilter, "@Index", AdjustSql(x_Index)) ' Replace key value
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

		' Field Index
		' Field Name

		sTmp = Trim(x_Name)
		If Trim(sTmp) = "" Then sTmp = ""
		rs("Name") = sTmp

		' Field Memo
		sTmp = Trim(x_Memo)
		If Trim(sTmp) = "" Then sTmp = Null
		rs("Memo") = sTmp

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
