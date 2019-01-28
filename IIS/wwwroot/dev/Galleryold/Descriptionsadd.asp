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
x_Description = Null
x_Name = Null
%>
<!--#include file="db.asp"-->
<!--#include file="aspmkrfn.asp"-->
<%
Response.Buffer = True

' Get action
sAction = Request.Form("a_add")
If (sAction = "" Or IsNull(sAction)) Then
	sKey = Request.Querystring("key")
	If sKey <> "" Then
		sAction = "C" ' Copy record
	Else
		sAction = "I" ' Display blank record
	End If
Else

	' Get fields from form
	x_Description = Request.Form("x_Description")
	x_Name = Request.Form("x_Name")
End If

' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case sAction
	Case "C": ' Get a record to display
		If Not LoadData(sKey) Then ' Load Record based on key
			Session("ewmsg") = "No Record Found for Key = " & sKey
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "Descriptionslist.asp"
		End If
	Case "A": ' Add
		If AddData() Then ' Add New Record
			Session("ewmsg") = "Add New Record Successful"
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "Descriptionslist.asp"
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
<p><span class="aspmaker">Add to TABLE: Descriptions<br><br><a href="Descriptionslist.asp">Back to List</a></span></p>
<form name="Descriptionsadd" id="Descriptionsadd" action="Descriptionsadd.asp" method="post" onSubmit="return EW_checkMyForm(this);">
<p>
<input type="hidden" name="a_add" value="A">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#3366CC"><span class="aspmaker" style="color: #FFFFFF;">Name</span></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_Name" id="x_Name" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_Name&"") %>">
</span></td>
	</tr>
</table>
<p>
<input type="submit" name="Action" value="ADD">
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
	sSql = "SELECT * FROM [Descriptions]"
	sSql = sSql & " WHERE [Description] = " & sKeyWrk
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
		x_Description = rs("Description")
		x_Name = rs("Name")
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
	Dim sSql, rs, sWhere, sGroupBy, sHaving, sOrderBy

	' Add New Record
	sSql = "SELECT * FROM [Descriptions]"
	sSql = sSql & " WHERE 0 = 1"
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
	rs.AddNew

	' Field Name
	sTmp = Trim(x_Name)
	If Trim(sTmp) = "" Then sTmp = ""
	sSqlChk = "SELECT * FROM [Descriptions] WHERE [Name] = '" & AdjustSql(sTmp) & "'"
	Set rsChk = conn.Execute(sSqlChk)
	If Not rsChk.Eof Then
		Response.Write "Duplicate value for index or primary key -- Name, value = " & sTmp & "<br>"
		Response.Write "Press [Previous Page] key to continue!"
		Response.End
	End If
	rsChk.Close
	Set rsChk = Nothing
	rs("Name") = sTmp
	rs.Update
	rs.Close
	Set rs = Nothing
	AddData = True
End Function
%>
