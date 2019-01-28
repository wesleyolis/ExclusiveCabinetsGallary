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
If sKey = "" Or IsNull(sKey) Then Response.Redirect "Rangeslist.asp"

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
		If Not LoadData(sKey) Then ' Load Record based on key
			Session("ewmsg") = "No Record Found for Key = " & sKey
			conn.Close ' Close Connection
			Set conn = Nothing
			Response.Clear
			Response.Redirect "Rangeslist.asp"
		End If
End Select
%>
<!--#include file="header.asp"-->
<p><span class="aspmaker">View TABLE: Ranges<br><br>
<a href="Rangeslist.asp">Back to List</a>&nbsp;
<a href="<%= "Rangesedit.asp?key=" & Server.URLEncode(sKey) %>">Edit</a>&nbsp;
<a href="<%= "Rangesadd.asp?key=" & Server.URLEncode(sKey) %>">Copy</a>&nbsp;
<a href="<%= "Rangesdelete.asp?key=" & Server.URLEncode(sKey) %>">Delete</a>&nbsp;
</span></p>
<p>
<form>
<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td bgcolor="#3366CC"><span class="aspmaker" style="color: #FFFFFF;">Range</span></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_Range %>
</span></td>
	</tr>
	<tr>
		<td bgcolor="#3366CC"><span class="aspmaker" style="color: #FFFFFF;">Name</span></td>
		<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_Name %>
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
