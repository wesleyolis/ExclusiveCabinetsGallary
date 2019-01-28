<%@ CodePage = 1252 LCID = 7177 %>
<% Session.Timeout = 20 %>
<%
Response.expires = 0
Response.expiresabsolute = Now() - 1
Response.addHeader "pragma", "no-cache"
Response.addHeader "cache-control", "private"
Response.CacheControl = "no-cache"
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
x_Range = Null
x_Name = Null
%>
<!--#include file="db.asp"-->
<!--#include file="aspmkrfn.asp"-->
<%
Response.Buffer = True
sKey = Request.Querystring("key")
If sKey = "" Or IsNull(sKey) Then sKey = Request.Form("key")

' Get action
sAction = Request.Form("a_edit")
If sAction = "" Or IsNull(sAction) Then
	sAction = "I"	' Display with input box
Else

	' Get fields from form
	x_Range = Request.Form("x_Range")
	x_Name = Request.Form("x_Name")
End If
If sKey = "" Or IsNull(sKey) Then Response.Redirect "Rangeslist.asp"

' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case sAction
	Case "I": ' Get a record to display
		If Not LoadData(sKey) Then ' Load Record based on key
			Session("ewmsg") = "No Record Found for Key = " & sKey
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "Rangeslist.asp"
		End If
	Case "U": ' Update
		If EditData(sKey) Then ' Update Record based on key
			Session("ewmsg") = "Update Record Successful for Key = " & sKey
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "Rangeslist.asp"
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
if (EW_this.x_Name && !EW_hasValue(EW_this.x_Name, "TEXT" )) {
	if (!EW_onError(EW_this, EW_this.x_Name, "TEXT", "Please enter required field - Name"))
		return false;
}
return true;
}
//-->
</script>
<p><span class="aspmaker">Edit TABLE: Ranges<br><br><a href="Rangeslist.asp">Back to List</a></span></p>
<form name="Rangesedit" id="Rangesedit" action="Rangesedit.asp" method="post" onSubmit="return EW_checkMyForm(this);">
<p>
<input type="hidden" name="a_edit" value="U">
<input type="hidden" name="key" value="<%= sKey %>">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#3366CC"><span class="aspmaker" style="color: #FFFFFF;">Range</span></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_Range %><input type="hidden" name="x_Range" value="<%= x_Range %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="#3366CC"><span class="aspmaker" style="color: #FFFFFF;">Name</span></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_Name" id="x_Name" size="30" maxlength="20" value="<%= Server.HTMLEncode(x_Name&"") %>">
</span></td>
	</tr>
</table>
<p>
<input type="submit" name="Action" value="EDIT">
</form>
<!--#include file="footer.asp"-->
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
	sSql = "SELECT * FROM [Ranges]"
	sSql = sSql & " WHERE [Range] = " & sKeyWrk
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
		x_Range = rs("Range")
		x_Name = rs("Name")
	End If
	rs.Close
	Set rs = Nothing
End Function
%>
<%

'-------------------------------------------------------------------------------
' Function EditData
' - Edit Data based on Key Value sKey
' - Variables used: field variables

Function EditData(sKey)
	Dim sKeyWrk, sSql, rs, sWhere, sGroupBy, sHaving, sOrderBy

	' Open record
	sKeyWrk = "" & AdjustSql(sKey) & ""
	sSql = "SELECT * FROM [Ranges]"
	sSql = sSql & " WHERE [Range] = " & sKeyWrk
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
	If rs.Eof Then
		EditData = False ' Update Failed
	Else

		' Field Range
		' Field Name

		sTmp = Trim(x_Name)
		If Trim(sTmp) = "" Then sTmp = ""
		rs("Name") = sTmp
		rs.Update
		EditData = True ' Update Successful
	End If
	rs.Close
	Set rs = Nothing
End Function
%>
