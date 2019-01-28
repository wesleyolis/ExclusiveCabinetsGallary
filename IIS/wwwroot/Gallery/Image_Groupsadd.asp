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
bCopy = True
x_Image = Request.QueryString("Image")
If x_Image = "" Or IsNull(x_Image) Then
	bCopy = False
End If
x_Group = Request.QueryString("Group")
If x_Group = "" Or IsNull(x_Group) Then
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
	x_Group = Request.Form("x_Group")
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
			Response.Redirect "Image_Groupslist.asp"
		End If
	Case "A": ' Add
		If AddData() Then ' Add New Record
			Session(ewSessionMessage) = "Add New Record Successful"
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "Image_Groupslist.asp"
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
if (EW_this.x_Image && !EW_checkinteger(EW_this.x_Image.value)) {
	if (!EW_onError(EW_this, EW_this.x_Image, "TEXT", "Incorrect integer - Image"))
		return false; 
}
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
<p><span class="aspmaker">Add to TABLE: Image Groups<br><br><a href="Image_Groupslist.asp">Back to List</a></span></p>
<%
If Session(ewSessionMessage) <> "" Then
%>
<p><span class="ewmsg"><%= Session(ewSessionMessage) %></span></p>
<%
	Session(ewSessionMessage) = "" ' Clear message
End If
%>
<form name="fImage_Groupsadd" id="fImage_Groupsadd" action="Image_Groupsadd.asp" method="post" onSubmit="return EW_checkMyForm(this);">
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
	<tr id="r_Image">
		<td class="ewTableHeader"><span>Image</span></td>
		<td class="ewTableAltRow"><span id="cb_x_Image">
<% If Session("Image_Groups_MasterKey_Image") <> "" Then
x_Image = Session("Image_Groups_MasterKey_Image") %>
<img src="thumbs/<% Response.Write x_Image %>.jpg">

<input type="hidden" id="x_Image" name="x_Image" value="<%= x_Image %>">
<% Else %>
<input type="text" name="x_Image" id="x_Image" size="30" value="<%= Server.HTMLEncode(x_Image&"") %>">
<% End If %>
</span></td>
	</tr>
	<tr id="r_Group">
		<td class="ewTableHeader"><span>Group<span class='ewmsg'>&nbsp;*</span></span></td>
		<td class="ewTableAltRow"><span id="cb_x_Group">
<% If IsNull(x_Group) or x_Group = "" Then x_Group = 0 ' Set default value %>
<%
lst_x_Group = "<select id='x_Group' name='x_Group'>"
lst_x_Group = lst_x_Group & "<option value=''>Please Select</option>"
sSqlWrk = "SELECT [Index], [Name] FROM [Color_Groups]"
sSqlWrk = sSqlWrk & " ORDER BY [Name] Asc"
Set rswrk = Server.CreateObject("ADODB.Recordset")
rswrk.Open sSqlWrk, conn, 1, 2
If Not rswrk.Eof Then
	datawrk = rswrk.GetRows
	rowswrk = UBound(datawrk, 2)
	For rowcntwrk = 0 To rowswrk
		lst_x_Group = lst_x_Group & "<option value='" & datawrk(0, rowcntwrk) & "'"
		If CStr(datawrk(0, rowcntwrk)&"") = CStr(x_Group&"") Then
			lst_x_Group = lst_x_Group & " selected"
		End If
		lst_x_Group = lst_x_Group & ">" & datawrk(1, rowcntwrk) & "</option>"
	Next
End If
rswrk.Close
Set rswrk = Nothing
lst_x_Group = lst_x_Group & "</select>"
Response.Write lst_x_Group
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
	If x_Group = "" Or IsNull(x_Group) Then
		bCheckKey = False
	Else
		sFilter = Replace(sFilter, "@Group", AdjustSql(x_Group)) ' Replace key value
	End If
	If Not IsNumeric(x_Group) Then
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

	' Field Image
	sTmp = x_Image
	If Not IsNumeric(sTmp) Then
		sTmp = Null
	Else
		sTmp = CLng(sTmp)
	End If
	rs("Image") = sTmp

	' Field Group
	sTmp = x_Group
	If Not IsNumeric(sTmp) Then
		sTmp = Null
	Else
		sTmp = CLng(sTmp)
	End If
	rs("Group") = sTmp

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
	If sKey <> "" Then sKey = sKey & ","
	sKey = sKey & rsnew.Fields("Group")
End Sub
%>