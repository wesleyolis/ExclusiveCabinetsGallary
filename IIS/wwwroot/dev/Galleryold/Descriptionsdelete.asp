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

' Load Key Parameters
sKey = Request.querystring("key")
If sKey = "" Or IsNull(sKey) Then
	sKey = Request.Form("key_d")
End If
arRecKey = Split(sKey&"", ",")

' Multiple delete records
If UBound(arRecKey) = -1 Then Response.Redirect "Descriptionslist.asp"
For Each sRecKey In arRecKey

	' Remove spaces
	sRecKey = Trim(sRecKey)

	' Build the SQL
	sDbWhere = sDbWhere & "([Description]=" & AdjustSql(sRecKey) & ") OR "
Next
If Right(sDbWhere, 4) = " OR " Then sDbWhere = Left(sDbWhere, Len(sDbWhere)-4)

' Get action
sAction = Request.Form("a_delete")
If sAction = "" Or IsNull(sAction) Then
	sAction = "I"	' Display with input box
End If

' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str
Select Case sAction
	Case "I": ' Display
		If LoadRecordCount(sDbWhere) <= 0 Then
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "Descriptionslist.asp"
		End If
	Case "D": ' Delete
		If DeleteData(sDbWhere) Then
			Session("ewmsg") = "Delete Successful For Key = " & sKey
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "Descriptionslist.asp"
		End If
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Delete from TABLE: Descriptions<br><br><a href="Descriptionslist.asp">Back to List</a></span></p>
<form action="Descriptionsdelete.asp" method="post">
<p>
<input type="hidden" name="a_delete" value="D">
<input type="hidden" name="key_d" value="<%= sKey %>">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr bgcolor="#3366CC">
		<td valign="top"><span class="aspmaker" style="color: #FFFFFF;">Description</span></td>
		<td valign="top"><span class="aspmaker" style="color: #FFFFFF;">Name</span></td>
	</tr>
<%
nRecCount = 0
For Each sRecKey In arRecKey
	sRecKey = Trim(sRecKey)
	nRecCount = nRecCount + 1

	' Set row color
	sItemRowClass = " bgcolor=""#FFFFFF"""

	' Display alternate color for rows
	If nRecCount Mod 2 <> 0 Then
		sItemRowClass = " bgcolor=""#FFFFFF"""
	End If
	If LoadData(sRecKey) Then
%>
	<tr<%=sItemRowClass%>>
		<td><span class="aspmaker">
<% Response.Write x_Description %>
</span></td>
		<td><span class="aspmaker">
<% Response.Write x_Name %>
</span></td>
	</tr>
<%
	End If
Next
%>
</table>
<p>
<input type="submit" name="Action" value="CONFIRM DELETE">
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
' Function LoadRecordCount
' - Load Record Count based on input sql criteria sqlKey

Function LoadRecordCount(sqlKey)
	Dim sSql, rs, sWhere, sGroupBy, sHaving, sOrderBy
	sSql = "SELECT * FROM [Descriptions]"
	sSql = sSql & " WHERE " & sqlKey
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
	LoadRecordCount = rs.RecordCount
	rs.Close
	Set rs = Nothing
End Function
%>
<%

'-------------------------------------------------------------------------------
' Function DeleteData
' - Delete Records based on input sql criteria sqlKey

Function DeleteData(sqlKey)
	Dim sSql, rs, sWhere, sGroupBy, sHaving, sOrderBy
	sSql = "SELECT * FROM [Descriptions]"
	sSql = sSql & " WHERE " & sqlKey
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
	Do While Not rs.Eof
		rs.Delete
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
	DeleteData = True
End Function
%>
