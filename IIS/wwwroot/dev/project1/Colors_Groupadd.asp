<!--#include file="ewconfig.asp"-->
<!--#include file="db.asp"-->
<!--#include file="Colors_Groupinfo.asp"-->
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
x_Grp = Null: ox_Grp = Null: z_Grp = Null
x_Colours = Null: ox_Colours = Null: z_Colours = Null
%>
<%
Response.Buffer = True

' Load key from QueryString
bCopy = True
x_Index = Request.QueryString("Index")
If x_Index = "" Or IsNull(x_Index) Then
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
	x_Index = Request.Form("x_Index")
	x_Grp = Request.Form("x_Grp")
	x_Colours = Request.Form("x_Colours")
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
			Response.Redirect "Colors_Grouplist.asp"
		End If
	Case "A": ' Add
		If AddData() Then ' Add New Record
			Session(ewSessionMessage) = "Add New Record Successful"
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "Colors_Grouplist.asp"
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
<script type ="text/javascript" src="ewast.js"></script>
<script type="text/javascript">
<!--
EW_dateSep = "/"; // set date separator	
//-->
</script>
<script type="text/javascript">
<!--
function EW_checkMyForm(EW_this) {
if (EW_this.x_Grp && !EW_checkinteger(EW_this.x_Grp.value)) {
	if (!EW_onError(EW_this, EW_this.x_Grp, "TEXT", "Incorrect integer - Grp"))
		return false; 
}
if (EW_this.x_Colours && !EW_hasValue(EW_this.x_Colours, "SELECT" )) {
	if (!EW_onError(EW_this, EW_this.x_Colours, "SELECT", "Please enter required field - Colours"))
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
<p><span class="aspmaker">Add to TABLE: Colors Group<br><br><a href="Colors_Grouplist.asp">Back to List</a></span></p>
<%
If Session(ewSessionMessage) <> "" Then
%>
<p><span class="ewmsg"><%= Session(ewSessionMessage) %></span></p>
<%
	Session(ewSessionMessage) = "" ' Clear message
End If
%>
<form name="fColors_Groupadd" id="fColors_Groupadd" action="Colors_Groupadd.asp" method="post" onSubmit="return EW_checkMyForm(this);">
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
	<tr id="r_Grp">
		<td class="ewTableHeader"><span>Grp</span></td>
		<td class="ewTableAltRow"><span id="cb_x_Grp">
<% If Session("Colors_Group_MasterKey_Grp") <> "" Then
x_Grp = Session("Colors_Group_MasterKey_Grp") %>
<% Response.Write x_Grp %>
<input type="hidden" id="x_Grp" name="x_Grp" value="<%= x_Grp %>">
<% Else %>
<div><input type="text" name="x_Grp" id="x_Grp" size="30" value="<%=x_Grp%>" onblur="EW_astHideDiv('as_x_Grp');" onkeydown="EW_astOnKeyDown('x_Grp', 'as_x_Grp', event);" onkeypress="return EW_astOnKeyPress(event);" onkeyup="EW_astOnKeyUp('x_Grp', 'as_x_Grp', event);" autocomplete="off"></div>
<div class='ewAstList' style='visibility:hidden' id='as_x_Grp'></div>
<input type="hidden" id="sv_x_Grp" name="sv_x_Grp" value="">
<%
	sSqlWrk = "SELECT DISTINCT [Name], '' FROM [Color_Groups] WHERE ([Name] LIKE '@FILTER_VALUE%')"
	sSqlWrk = sSqlWrk & " ORDER BY [Name] Asc"
	sSqlWrk = EW_Encode(TEAencrypt(sSqlWrk, EW_RANDOM_KEY))
%>
<input type="hidden" name="s_x_Grp" value="<%= sSqlWrk %>">
<% End If %>
</span></td>
	</tr>
	<tr id="r_Colours">
		<td class="ewTableHeader"><span>Colours<span class='ewmsg'>&nbsp;*</span></span></td>
		<td class="ewTableAltRow"><span id="cb_x_Colours">
<% If IsNull(x_Colours) or x_Colours = "" Then x_Colours = 0 ' Set default value %>
<%
lst_x_Colours = "<select id='x_Colours' name='x_Colours'>"
lst_x_Colours = lst_x_Colours & "<option value=''>Please Select</option>"
sSqlWrk = "SELECT [Color], [Name] FROM [Colors]"
sSqlWrk = sSqlWrk & " ORDER BY [Name] Asc"
Set rswrk = Server.CreateObject("ADODB.Recordset")
rswrk.Open sSqlWrk, conn, 1, 2
If Not rswrk.Eof Then
	datawrk = rswrk.GetRows
	rowswrk = UBound(datawrk, 2)
	For rowcntwrk = 0 To rowswrk
		lst_x_Colours = lst_x_Colours & "<option value='" & datawrk(0, rowcntwrk) & "'"
		If CStr(datawrk(0, rowcntwrk)&"") = CStr(x_Colours&"") Then
			lst_x_Colours = lst_x_Colours & " selected"
		End If
		lst_x_Colours = lst_x_Colours & ">" & datawrk(1, rowcntwrk) & "</option>"
	Next
End If
rswrk.Close
Set rswrk = Nothing
lst_x_Colours = lst_x_Colours & "</select>"
Response.Write lst_x_Colours
%>
</span></td>
	</tr>
</table>
<p>
<input type="submit" name="btnAction" id="btnAction" value="ADD">
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
		x_Grp = rs("Grp")
		x_Colours = rs("Colours")
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
	If x_Index = "" Or IsNull(x_Index) Then
		bCheckKey = False
	Else
		sFilter = Replace(sFilter, "@Index", AdjustSql(x_Index)) ' Replace key value
	End If
	If Not IsNumeric(x_Index) Then
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

	' Field Grp
	sTmp = x_Grp
	If Not IsNumeric(sTmp) Then
		sTmp = Null
	Else
		sTmp = CLng(sTmp)
	End If
	rs("Grp") = sTmp

	' Field Colours
	sTmp = x_Colours
	If Not IsNumeric(sTmp) Then
		sTmp = Null
	Else
		sTmp = CLng(sTmp)
	End If
	rs("Colours") = sTmp

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
	sKey = sKey & rsnew.Fields("Index")
	x_Index = rsnew.Fields("Index")
End Sub
%>
