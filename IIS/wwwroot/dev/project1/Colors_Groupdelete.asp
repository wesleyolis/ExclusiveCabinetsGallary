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

' Load Key Parameters
sKey = "": bSingleDelete = True
x_Index = Request.QueryString("Index")
If x_Index <> "" Then
	If Not IsNumeric(x_Index) Then
		Response.Redirect "Colors_Grouplist.asp" ' prevent sql injection
	End If
	If sKey <> "" Then sKey = sKey & ","
	sKey = sKey & x_Index
Else
	bSingleDelete = False
End If
If Not bSingleDelete Then
	sKey = Request.Form("key_d")
End If
If sKey = "" Or IsNull(sKey) Then Response.Redirect "Colors_Grouplist.asp"
arRecKey = Split(sKey&"", ",")
i = 0
Do While i <= UBound(arRecKey)
	sDbWhere = sDbWhere & "("

	' Remove spaces
	sRecKey = Trim(arRecKey(i+0))
	If Not IsNumeric(sRecKey) Then
		Response.Redirect "Colors_Grouplist.asp" ' prevent sql injection
	End If

	' Build the SQL
	sDbWhere = sDbWhere & "[Index]=" & AdjustSql(sRecKey) & " AND "
	If Right(sDbWhere, 5) = " AND " Then sDbWhere = Left(sDbWhere, Len(sDbWhere)-5) & ") OR "
	i = i + 1
Loop
If Right(sDbWhere, 4) = " OR " Then sDbWhere = Left(sDbWhere, Len(sDbWhere)-4)

' Get action
sAction = Request.Form("a_delete")
If sAction = "" Or IsNull(sAction) Then
	sAction = "I"	' Display record
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
			Response.Redirect "Colors_Grouplist.asp"
		End If
	Case "D": ' Delete
		If DeleteData(sDbWhere) Then
			Session(ewSessionMessage) = "Delete Successful"
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "Colors_Grouplist.asp"
		End If
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Delete from TABLE: Colors Group<br><br><a href="Colors_Grouplist.asp">Back to List</a></span></p>
<%
If Session(ewSessionMessage) <> "" Then
%>
<p><span class="ewmsg"><%= Session(ewSessionMessage) %></span></p>
<%
	Session(ewSessionMessage) = "" ' Clear message
End If
%>
<form action="Colors_Groupdelete.asp" method="post">
<p>
<input type="hidden" name="a_delete" value="D">
<input type="hidden" name="key_d" value="<%= sKey %>">
<table class="ewTable">
	<tr class="ewTableHeader">
		<td valign="top"><span>Colours</span></td>
	</tr>
<%
nRecCount = 0
i = 0
Do While i <= UBound(arRecKey)
	nRecCount = nRecCount + 1

	' Set row color
	sItemRowClass = " class=""ewTableRow"""

	' Display alternate color for rows
	If nRecCount Mod 2 <> 0 Then
		sItemRowClass = " class=""ewTableAltRow"""
	End If
	sRecKey = Trim(arRecKey(i+0))
	x_Index = sRecKey
	If LoadData() Then
%>
	<tr<%=sItemRowClass%>>
		<td><span>
<%
If Not IsNull(x_Colours) Then
	sSqlWrk = "SELECT [Name] FROM [Colors]"
	sTmp = x_Colours
	sSqlWrk = sSqlWrk & " WHERE [Color] = " & AdjustSql(sTmp) & ""
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
ox_Colours = x_Colours ' Backup Original Value
x_Colours = sTmp
%>
<% Response.Write x_Colours %>
<% x_Colours = ox_Colours ' Restore Original Value %>
</span></td>
	</tr>
<%
	End If
	i = i + 1
Loop
%>
</table>
<p>
<input type="submit" name="Action" value="CONFIRM DELETE">
</form>
<%
conn.Close ' Close Connection
Set conn = Nothing
%>
<!--#include file="footer.asp"-->
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
' Function LoadRecordCount
' - Load Record Count based on input sql criteria sqlKey

Function LoadRecordCount(sqlKey)
	On Error Resume Next
	Dim rs, sSql, sFilter
	sFilter = sqlKey
	sSql = ewBuildSql(ewSqlSelect, ewSqlWhere, ewSqlGroupBy, ewSqlHaving, ewSqlOrderBy, sFilter, "")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open sSql, conn, 1, 2
	LoadRecordCount = rs.RecordCount
	rs.Close
	Set rs = Nothing
	If Err.Number <> 0 Then
		Session(ewSessionMessage) = Err.Description
	End If
End Function
%>
<%

'-------------------------------------------------------------------------------
' Function DeleteData
' - Delete Records based on input sql criteria sqlKey

Function DeleteData(sqlKey)
	On Error Resume Next
	Dim rs, sSql, sFilter
	sFilter = sqlKey
	sSql = ewBuildSql(ewSqlSelect, ewSqlWhere, ewSqlGroupBy, ewSqlHaving, ewSqlOrderBy, sFilter, "")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open sSql, conn, 1, 2
	If Err.Number <> 0 Then
		Session(ewSessionMessage) = Err.Description
		rs.Close
		Set rs = Nothing
		DeleteData = False
		Exit Function
	End If

	' Clone old rs object
	Dim rsold
	Set rsold = rs.Clone(1)
	rsold.Requery

	' Call recordset deleting event
	DeleteData = Recordset_Deleting(rs)
	If DeleteData Then
		Do While Not rs.Eof
			rs.Delete
			If Err.Number <> 0 Then
				Session(ewSessionMessage) = Err.Description
				DeleteData = False
				Exit Do
			End If
			rs.MoveNext
		Loop
	End If
	rs.Close
	Set rs = Nothing

	' Call recordset deleted event
	If DeleteData Then
		Call Recordset_Deleted(rsold)
	End If
	rsold.Close
	Set rsold = Nothing
End Function

'-------------------------------------------------------------------------------
' Recordset deleting event

Function Recordset_Deleting(rsold)
	On Error Resume Next

	' Please enter your customized codes here
	Recordset_Deleting = True
End Function

'-------------------------------------------------------------------------------
' Recordset deleted event

Sub Recordset_Deleted(rsold)
	On Error Resume Next
End Sub
%>
