<!--#include file="ewconfig.asp"-->
<!--#include file="db.asp"-->
<!--#include file="Imagesinfo.asp"-->
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
x_Dir = Null: ox_Dir = Null: z_Dir = Null
x_Code = Null: ox_Code = Null: z_Code = Null
x_Client = Null: ox_Client = Null: z_Client = Null
x_Range = Null: ox_Range = Null: z_Range = Null
x_Description = Null: ox_Description = Null: z_Description = Null
x_Color = Null: ox_Color = Null: z_Color = Null
x_Width = Null: ox_Width = Null: z_Width = Null
x_Height = Null: ox_Height = Null: z_Height = Null
x_Depth = Null: ox_Depth = Null: z_Depth = Null
x_Price = Null: ox_Price = Null: z_Price = Null
x_Edge = Null: ox_Edge = Null: z_Edge = Null
x_Sync_info = Null: ox_Sync_info = Null: z_Sync_info = Null
x_Sync_img = Null: ox_Sync_img = Null: z_Sync_img = Null
x_Info = Null: ox_Info = Null: z_Info = Null
%>
<%
Response.Buffer = True

' Load Key Parameters
sKey = "": bSingleDelete = True
x_Image = Request.QueryString("Image")
If x_Image <> "" Then
	If Not IsNumeric(x_Image) Then
		Response.Redirect "Imageslist.asp" ' prevent sql injection
	End If
	If sKey <> "" Then sKey = sKey & ","
	sKey = sKey & x_Image
Else
	bSingleDelete = False
End If
If Not bSingleDelete Then
	sKey = Request.Form("key_d")
End If
If sKey = "" Or IsNull(sKey) Then Response.Redirect "Imageslist.asp"
arRecKey = Split(sKey&"", ",")
i = 0
Do While i <= UBound(arRecKey)
	sDbWhere = sDbWhere & "("

	' Remove spaces
	sRecKey = Trim(arRecKey(i+0))
	If Not IsNumeric(sRecKey) Then
		Response.Redirect "Imageslist.asp" ' prevent sql injection
	End If

	' Build the SQL
	sDbWhere = sDbWhere & "[Image]=" & AdjustSql(sRecKey) & " AND "
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
			Response.Redirect "Imageslist.asp"
		End If
	Case "D": ' Delete
		If DeleteData(sDbWhere) Then
			Session(ewSessionMessage) = "Delete Successful"
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "Imageslist.asp"
		End If
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">Delete from TABLE: Images<br><br><a href="Imageslist.asp">Back to List</a></span></p>
<%
If Session(ewSessionMessage) <> "" Then
%>
<p><span class="ewmsg"><%= Session(ewSessionMessage) %></span></p>
<%
	Session(ewSessionMessage) = "" ' Clear message
End If
%>
<form action="Imagesdelete.asp" method="post">
<p>
<input type="hidden" name="a_delete" value="D">
<input type="hidden" name="key_d" value="<%= sKey %>">
<table class="ewTable">
	<tr class="ewTableHeader">
		<td valign="top"><span>Image</span></td>
		<td valign="top"><span>Code</span></td>
		<td valign="top"><span>Client</span></td>
		<td valign="top"><span>Range</span></td>
		<td valign="top"><span>Description</span></td>
		<td valign="top"><span>Color</span></td>
		<td valign="top"><span>Width</span></td>
		<td valign="top"><span>Height</span></td>
		<td valign="top"><span>Depth</span></td>
		<td valign="top"><span>Price</span></td>
		<td valign="top"><span>Edge</span></td>
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
	x_Image = sRecKey
	If LoadData() Then
%>
	<tr<%=sItemRowClass%>>
		<td><span>
<% sTmp = x_Image %><% If Not IsNull(sTmp) Then %><a href="Imagesview.asp?Image=<%= sTmp %>"><img src="<%= x_Image %>" Border=0></a><% Else %><img src="<%= x_Image %>" Border=0><% End If %>
</span></td>
		<td><span>
<% Response.Write x_Code %>
</span></td>
		<td><span>
<% Response.Write x_Client %>
</span></td>
		<td><span>
<%
If Not IsNull(x_Range) Then
	sSqlWrk = "SELECT [Name] FROM [Ranges]"
	sTmp = x_Range
	sSqlWrk = sSqlWrk & " WHERE [Range] = " & AdjustSql(sTmp) & ""
	Set rswrk = conn.Execute(sSqlWrk)
	If Not rswrk.Eof Then
		sTmp = rswrk("Name")
	End If
	rswrk.Close
	Set rswrk = Nothing
Else
	sTmp = Null
End If
ox_Range = x_Range ' Backup Original Value
x_Range = sTmp
%>
<% sTmp = x_Range %><% If Not IsNull(sTmp) Then %><a href="Imageslist.asp?x_Range<%= sTmp %>"><% Response.Write x_Range %></a><% Else %><% Response.Write x_Range %><% End If %>
<% x_Range = ox_Range ' Restore Original Value %>
</span></td>
		<td><span>
<%
If Not IsNull(x_Description) Then
	sSqlWrk = "SELECT [Name] FROM [Descriptions]"
	sTmp = x_Description
	sSqlWrk = sSqlWrk & " WHERE [Description] = " & AdjustSql(sTmp) & ""
	Set rswrk = conn.Execute(sSqlWrk)
	If Not rswrk.Eof Then
		sTmp = rswrk("Name")
	End If
	rswrk.Close
	Set rswrk = Nothing
Else
	sTmp = Null
End If
ox_Description = x_Description ' Backup Original Value
x_Description = sTmp
%>
<% sTmp = x_Description %><% If Not IsNull(sTmp) Then %><a href="Imageslist.asp?x_Description<%= sTmp %>"><% Response.Write x_Description %></a><% Else %><% Response.Write x_Description %><% End If %>
<% x_Description = ox_Description ' Restore Original Value %>
</span></td>
		<td><span>
<%
If Not IsNull(x_Color) Then
	sSqlWrk = "SELECT [Name] FROM [Colors]"
	sTmp = x_Color
	sSqlWrk = sSqlWrk & " WHERE [Color] = " & AdjustSql(sTmp) & ""
	Set rswrk = conn.Execute(sSqlWrk)
	If Not rswrk.Eof Then
		sTmp = rswrk("Name")
	End If
	rswrk.Close
	Set rswrk = Nothing
Else
	sTmp = Null
End If
ox_Color = x_Color ' Backup Original Value
x_Color = sTmp
%>
<% Response.Write x_Color %>
<% x_Color = ox_Color ' Restore Original Value %>
</span></td>
		<td><span>
<% Response.Write x_Width %>
</span></td>
		<td><span>
<% Response.Write x_Height %>
</span></td>
		<td><span>
<% Response.Write x_Depth %>
</span></td>
		<td><span>
<% Response.Write x_Price %>
</span></td>
		<td><span>
<% Response.Write x_Edge %>
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
	If Not IsNumeric(x_Image) Then
		LoadData = False
		Exit Function
	End If
	sFilter = Replace(sFilter, "@Image", AdjustSql(x_Image)) ' Replace key value
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
		x_Dir = rs("Dir")
		x_Code = rs("Code")
		x_Client = rs("Client")
		x_Range = rs("Range")
		x_Description = rs("Description")
		x_Color = rs("Color")
		x_Width = rs("Width")
		x_Height = rs("Height")
		x_Depth = rs("Depth")
		x_Price = rs("Price")
		x_Edge = rs("Edge")
		x_Sync_info = rs("Sync_info")
		x_Sync_img = rs("Sync_img")
		x_Info = rs("Info")
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
