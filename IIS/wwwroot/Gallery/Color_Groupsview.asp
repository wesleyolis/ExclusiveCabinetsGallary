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
x_Index = Request.QueryString("Index")
If x_Index = "" Or IsNull(x_Index) Then Response.Redirect "Color_Groupslist.asp"

' Get action
sAction = Request.Form("a_view")
If sAction = "" Or IsNull(sAction) Then
	sAction = "I"	' Display with input box
End If

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
<p><span class="aspmaker">View TABLE: Color Groups<br><br>
<a href="Color_Groupslist.asp">Back to List</a>&nbsp;
<a href="<% If Not IsNull(x_Index) Then Response.Write "Color_Groupsedit.asp?Index=" & Server.URLEncode(x_Index) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Edit</a>&nbsp;
<a href="<% If Not IsNull(x_Index) Then Response.Write "Color_Groupsdelete.asp?Index=" & Server.URLEncode(x_Index) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Delete</a>&nbsp;
<a href="Colors_Grouplist.asp?showmaster=1&Grp=<%= Server.URLEncode(x_Index)%>">Colors Group Details</a>&nbsp;
</span></p>
<p>
<form>
<table class="ewTable">
	<tr>
		<td class="ewTableHeader"><span>Name</span></td>
		<td class="ewTableAltRow"><span>
<% Response.Write x_Name %>
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Memo</span></td>
		<td class="ewTableAltRow"><span>
<%= Replace(x_Memo&"", vbLf, "<br>") %>
</span></td>
	</tr>
</table>
</form>
<p>
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
