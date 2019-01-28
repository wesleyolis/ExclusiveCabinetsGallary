<!--#include file="ewconfig.asp"-->
<!--#include file="db.asp"-->
<!--#include file="Rangesinfo.asp"-->
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
x_Range = Null: ox_Range = Null: z_Range = Null
x_Name = Null: ox_Name = Null: z_Name = Null
x_Sync = Null: ox_Sync = Null: z_Sync = Null
%>
<%
Response.Buffer = True
x_Range = Request.QueryString("Range")
If x_Range = "" Or IsNull(x_Range) Then Response.Redirect "Rangeslist.asp"

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
			Response.Redirect "Rangeslist.asp"
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
<p><span class="aspmaker">View TABLE: Ranges<br><br>
<a href="Rangeslist.asp">Back to List</a>&nbsp;
<a href="<% If Not IsNull(x_Range) Then Response.Write "Rangesedit.asp?Range=" & Server.URLEncode(x_Range) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Edit</a>&nbsp;
<a href="<% If Not IsNull(x_Range) Then Response.Write "Rangesadd.asp?Range=" & Server.URLEncode(x_Range) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Copy</a>&nbsp;
<a href="<% If Not IsNull(x_Range) Then Response.Write "Rangesdelete.asp?Range=" & Server.URLEncode(x_Range) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Delete</a>&nbsp;
</span></p>
<p>
<form>
<table class="ewTable">
	<tr>
		<td class="ewTableHeader"><span>Range</span></td>
		<td class="ewTableAltRow"><span>
<% Response.Write x_Range %>
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Name</span></td>
		<td class="ewTableAltRow"><span>
<% Response.Write x_Name %>
</span></td>
	</tr>
	<tr>
		<td class="ewTableHeader"><span>Sync</span></td>
		<td class="ewTableAltRow"><span>
<%
If x_Sync = True Then
	sTmp = "Yes"
Else
	sTmp = "No"
End If
ox_Sync = x_Sync ' Backup Original Value
x_Sync = sTmp
%>
<% Response.Write x_Sync %>
<% x_Sync = ox_Sync ' Restore Original Value %>
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
	If Not IsNumeric(x_Range) Then
		LoadData = False
		Exit Function
	End If
	sFilter = Replace(sFilter, "@Range", AdjustSql(x_Range)) ' Replace key value
	sSql = ewBuildSql(ewSqlSelect, ewSqlWhere, ewSqlGroupBy, ewSqlHaving, ewSqlOrderBy, sFilter, "")
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
		x_Sync = rs("Sync")
	End If
	rs.Close
	Set rs = Nothing
End Function
%>
