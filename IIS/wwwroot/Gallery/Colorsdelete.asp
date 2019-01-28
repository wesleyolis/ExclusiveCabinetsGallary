<!--#include file="ewconfig.asp"-->
<!--#include file="db.asp"-->
<!--#include file="Colorsinfo.asp"-->
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
x_Color = Null: ox_Color = Null: z_Color = Null
x_Name = Null: ox_Name = Null: z_Name = Null
x_Sync = Null: ox_Sync = Null: z_Sync = Null
%>
<%
Response.Buffer = True

' Load Key Parameters
sKey = "": bSingleDelete = True
x_Color = Request.QueryString("Color")
If x_Color <> "" Then
	If Not IsNumeric(x_Color) Then
		Response.Redirect "Colorslist.asp" ' prevent sql injection
	End If
	If sKey <> "" Then sKey = sKey & ","
	sKey = sKey & x_Color
Else
	bSingleDelete = False
End If
If Not bSingleDelete Then
	sKey = Request.Form("key_d")
End If
If sKey = "" Or IsNull(sKey) Then Response.Redirect "Colorslist.asp"
arRecKey = Split(sKey&"", ",")
i = 0
Do While i <= UBound(arRecKey)
	sDbWhere = sDbWhere & "("

	' Remove spaces
	sRecKey = Trim(arRecKey(i+0))
	If Not IsNumeric(sRecKey) Then
		Response.Redirect "Colorslist.asp" ' prevent sql injection
	End If

	' Build the SQL
	sDbWhere = sDbWhere & "[Color]=" & AdjustSql(sRecKey) & " AND "
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
			Response.Redirect "Colorslist.asp"
		End If
	Case "D": ' Delete
		If DeleteData(sDbWhere) Then
			Session(ewSessionMessage) = "Delete Successful"
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "Colorslist.asp"
		End If
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Delete from TABLE: Colors<br><br><a href="Colorslist.asp">Back to List</a></span></p>
<%
If Session(ewSessionMessage) <> "" Then
%>
<p><span class="ewmsg"><%= Session(ewSessionMessage) %></span></p>
<%
	Session(ewSessionMessage) = "" ' Clear message
End If
%>
<form action="Colorsdelete.asp" method="post">
<p>
<input type="hidden" name="a_delete" value="D">
<input type="hidden" name="key_d" value="<%= sKey %>">
<table class="ewTable">
	<tr class="ewTableHeader">
		<td valign="top"><span>Color</span></td>
		<td valign="top"><span>Name</span></td>
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
	x_Color = sRecKey
	If LoadData() Then
%>
	<tr<%=sItemRowClass%>>
		<td><span>
<img src="thumbs/Color_<%= x_Color %>.jpg" Border=0>
</span></td>
		<td><span>
<% Response.Write x_Name %>
</span></td>
	</tr>
<%
	End If
	i = i + 1
Loop
%>
</table>
<p>Delete Original Copy <input type="checkbox" name="o" value="True"></p>
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
	If Not IsNumeric(x_Color) Then
		LoadData = False
		Exit Function
	End If
	sFilter = Replace(sFilter, "@Color", AdjustSql(x_Color)) ' Replace key value
	sSql = ewBuildSql(ewSqlSelect, ewSqlWhere, ewSqlGroupBy, ewSqlHaving, ewSqlOrderBy, sFilter, "")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open sSql, conn
	If rs.Eof Then
		LoadData = False
	Else
		LoadData = True
		rs.MoveFirst

		' Get the field contents
		x_Color = rs("Color")
		x_Name = rs("Name")
		x_Sync = rs("Sync")
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
	
	'Delte image too
	orig = Null
	Set objFSO = Server.CreateObject("Scripting.FileSystemObject")


	path = Server.MapPath("thumbs")
		If (objFSO.FileExists(path & "/Color_" & sKey & ".jpg"))=True Then
		objFSO.DeleteFile path & "/Color_" & sKey & ".jpg",TRUE
		
		End IF
		
		
		path = Server.MapPath("orig")
		If (objFSO.FileExists(path & "/Color_" & sKey & ".jpg"))=True Then
			If Request.Form("o") <> "True" Then
			path2 = Server.MapPath("old_orig")
			objFSO.CopyFile path & "/Color_" & sKey & ".jpg", path2 & "/Color_" & sKey & ".jpg",TRUE
			End IF
			objFSO.DeleteFile path & "/Color_" & sKey & ".jpg",TRUE
		End IF

	

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