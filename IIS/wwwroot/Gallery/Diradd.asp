<%@ CodePage = 1252 LCID = 7177 %>
<% Session.Timeout = 20 %>
<%
Response.expires = 0
Response.expiresabsolute = Now() - 2
Response.addHeader "pragma",  "no-cache"
Response.addHeader "cache-control", "no-store"
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
x_Dir = Null
x_UDir = Null
x_Name = Null
x_Description = Null
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
		x_Dir  = sKey  
		x_UDir  = Request.Querystring("d") 
	End If
Else

	' Get fields from form
	x_Dir = Request.Form("x_Dir")
	x_UDir = Request.Form("x_UDir")
	x_Name = Request.Form("x_Name")
	x_Description = Request.Form("x_Description")
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
			Response.Redirect "Dirbrowser.asp?d=" & x_UDir
		End If
	Case "A": ' Add
		If AddData() Then ' Add New Record
			Session("ewmsg") = "Add New Record Successful"
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "Dirbrowser.asp?d=" & x_UDir
		End If
	Case "U": ' Update
		If EditData(x_Dir) Then ' Update Record based on key
			Session("ewmsg") = "Update Record Successful for Key = " & sKey
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "Dirbrowser.asp?d=" & x_Dir
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
if (EW_this.x_UDir && !EW_hasValue(EW_this.x_UDir, "TEXT" )) {
	if (!EW_onError(EW_this, EW_this.x_UDir, "TEXT", "Please enter required field - UDir"))
		return false;
}
if (EW_this.x_UDir && !EW_checkinteger(EW_this.x_UDir.value)) {
	if (!EW_onError(EW_this, EW_this.x_UDir, "TEXT", "Incorrect integer - UDir"))
		return false; 
}
if (EW_this.x_Name && !EW_hasValue(EW_this.x_Name, "TEXT" )) {
	if (!EW_onError(EW_this, EW_this.x_Name, "TEXT", "Please enter required field - Name"))
		return false;
}
return true;
}
//-->
</script>
<p><span class="aspmaker">Add to TABLE: Dir<br><br><a href="Dirlist.asp">Back to List</a></span></p>
<form name="Diradd" id="Diradd" action="Diradd.asp" method="post" onSubmit="return EW_checkMyForm(this);">
<p>
<input type="hidden" name="a_add" value="A">
<input type="hidden" name="x_Dir" id="x_Dir" value="<%=x_Dir%>">
<input type="hidden" name="x_Udir" id="x_UDir" value="<%=x_UDir%>">
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#3366CC"><span class="aspmaker" style="color: #FFFFFF;">Name</span></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<input type="text" name="x_Name" id="x_Name" size="30" maxlength="50" value="<%= Server.HTMLEncode(x_Name&"") %>">
</span></td>
	</tr>
	<tr>
		<td bgcolor="#3366CC"><span class="aspmaker" style="color: #FFFFFF;">Description</span></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<textarea cols="35" rows="4" id="x_Description" name="x_Description"><%= x_Description %></textarea>
</span></td>
	</tr>
</table>
<p>
<%if sAction <> "I" Then%>
<input onclick="document.all.a_add.value='A'" type="submit" name="Action" value="ADD Copy">
<input onclick="document.all.a_add.value='U'" type="submit" name="Update" value="Update">
<%Else%>
<input onclick="document.all.a_add.value='A'" type="submit" name="Action" value="ADD">
<%End IF%></form>
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
	sSql = "SELECT * FROM [Dir]"
	sSql = sSql & " WHERE [Dir] = " & sKeyWrk
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
		x_Dir = rs("Dir")
		x_UDir = rs("UDir")
		x_Name = rs("Name")
		x_Description = rs("Description")
		If  x_Description <> "" Then
		x_Description = Replace(x_Description,"<br>",VbCrLf)
		End IF
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
	sSql = "SELECT * FROM [Dir]"
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

	' Field UDir
	sTmp = x_UDir
	If Not IsNumeric(sTmp) Then
		sTmp = 0
	Else
		sTmp = CLng(sTmp)
	End If
	rs("UDir") = sTmp

	' Field Name
	sTmp = Trim(x_Name)
	If Trim(sTmp) = "" Then sTmp = ""
	rs("Name") = sTmp

	' Field Description
	Trim(x_Description)
	If Trim(sTmp) = "" Then 
	sTmp = Null
	Else
	sTmp = Replace(sTmp,VbCrLf,"<br>")
	End If
	
	rs("Description") = sTmp
	rs.Update
	rs.Close
	Set rs = Nothing
	AddData = True
End Function



Function EditData(sKey)
	Dim sKeyWrk, sSql, rs, sWhere, sGroupBy, sHaving, sOrderBy

	' Open record
	sKeyWrk = "" & AdjustSql(sKey) & ""
	sSql = "SELECT * FROM [Dir]"
	sSql = sSql & " WHERE [Dir] = " & sKeyWrk
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

		' Field Dir
		' Field UDir

		sTmp = x_UDir
		If Not IsNumeric(sTmp) Then
			sTmp = 0
		Else
			sTmp = CLng(sTmp)
		End If
		rs("UDir") = sTmp

		' Field Name
		sTmp = Trim(x_Name)
		If Trim(sTmp) = "" Then sTmp = ""
		rs("Name") = sTmp

		' Field Description
		sTmp = Trim(x_Description)
		If Trim(sTmp) = "" Then sTmp = Null
		rs("Description") = sTmp
		rs.Update
		EditData = True ' Update Successful
	End If
	rs.Close
	Set rs = Nothing
End Function

%>