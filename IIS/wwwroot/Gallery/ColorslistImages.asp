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
x_Color = Null
x_Name = Null
%>
<!--#include file="db.asp"-->
<!--#include file="aspmkrfn.asp"-->
<% 
nStartRec = 0
nStopRec = 0
nTotalRecs = 0
nRecCount = 0
nRecActual = 0
sKeyMaster = ""
sDbWhereMaster = ""
sSrchAdvanced = ""
sSrchBasic = ""
sSrchWhere = ""
sDbWhere = ""
sDefaultOrderBy = ""
sDefaultFilter = ""
sWhere = ""
sGroupBy = ""
sHaving = ""
sOrderBy = ""
sSqlMaster = ""
nDisplayRecs = 20
nRecRange = 10

' Set up records per page dynamically
SetUpDisplayRecs()

' Multi Column
nRecPerRow = 1

' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str

' Handle Reset Command
ResetCmd()

' Build WHERE condition
sDbWhere = ""
If sDbWhereMaster <> "" Then
	sDbWhere = sDbWhere & "(" & sDbWhereMaster & ") AND "
End If
If sSrchWhere <> "" Then
	sDbWhere = sDbWhere & "(" & sSrchWhere & ") AND "
End If
If Len(sDbWhere) > 5 Then
	sDbWhere = Mid(sDbWhere, 1, Len(sDbWhere)-5) ' Trim rightmost AND
End If

' Build SQL
sSql = "SELECT * FROM [Colors]"

' Load Default Filter
sDefaultFilter = ""
sGroupBy = ""
sHaving = ""

' Load Default Order
sDefaultOrderBy = ""
sWhere = ""
If sDefaultFilter <> "" Then
	sWhere = sWhere & "(" & sDefaultFilter & ") AND "
End If
If sDbWhere <> "" Then
	sWhere = sWhere & "(" & sDbWhere & ") AND "
End If
If Right(sWhere, 5) = " AND " Then sWhere = Left(sWhere, Len(sWhere)-5)
If sWhere <> "" Then
	sSql = sSql & " WHERE " & sWhere
End If
If sGroupBy <> "" Then
	sSql = sSql & " GROUP BY " & sGroupBy
End If	
If sHaving <> "" Then
	sSql = sSql & " HAVING " & sHaving
End If	

' Set Up Sorting Order
sOrderBy = ""
SetUpSortOrder()
If sOrderBy <> "" Then
	sSql = sSql & " ORDER BY " & sOrderBy
End If	

'Session("ewmsg") = sSql ' Uncomment to show SQL for debugging
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
function EW_selectKey(elem) {
	var f = elem.form;	
	if (!f.key_d) return;
	if (f.key_d[0]) {
		for (var i=0; i<f.key_d.length; i++)
			f.key_d[i].checked = elem.checked;	
	} else {
		f.key_d.checked = elem.checked;	
	}
}
//-->
</script>
<%

' Set up Record Set
Set rs = Server.CreateObject("ADODB.Recordset")
rs.CursorLocation = 3
rs.Open sSql, conn, 1, 2
nTotalRecs = rs.RecordCount
If nDisplayRecs <= 0 Then ' Display All Records
	nDisplayRecs = nTotalRecs
End If
nStartRec = 1
SetUpStartRec() ' Set Up Start Record Position
%>
<p><span class="aspmaker"><font size="4"><b>Admin Edit</b></font></span><b><span class="aspmaker"><font size="4"> 
Colours</font></span></b></p>
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td><span class="aspmaker"><a href="Colorsadd.asp">Add</a></span></td>
	</tr>
</table>
<p>
<%
If Session("ewmsg") <> "" Then
%>
<p><span class="aspmaker" style="color: Red;"><%= Session("ewmsg") %></span></p>
<%
	Session("ewmsg") = "" ' Clear message
End If
%>
<form action="Colorslist.asp" name="ewpagerform" id="ewpagerform">
<table bgcolor="" border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
	<tr>
		<td nowrap>
<span class="aspmaker">
<%

' Display page numbers
If nTotalRecs > 0 Then
	rsEof = (nTotalRecs < (nStartRec + nDisplayRecs))
	If CLng(nTotalRecs) > CLng(nDisplayRecs) Then

		' Find out if there should be Backward or Forward Buttons on the TABLE.
		If 	nStartRec = 1 Then
			isPrev = False
		Else
			isPrev = True
			PrevStart = nStartRec - nDisplayRecs
			If PrevStart < 1 Then PrevStart = 1 %>	
		<a href="Colorslist.asp?start=<%=PrevStart%>"><b>Prev</b></a>
		<%
		End If
		If (isPrev Or (Not rsEof)) Then
			x = 1
			y = 1
			dx1 = ((nStartRec-1)\(nDisplayRecs*nRecRange))*nDisplayRecs*nRecRange+1
			dy1 = ((nStartRec-1)\(nDisplayRecs*nRecRange))*nRecRange+1
			If (dx1+nDisplayRecs*nRecRange-1) > nTotalRecs Then
				dx2 = (nTotalRecs\nDisplayRecs)*nDisplayRecs+1
				dy2 = (nTotalRecs\nDisplayRecs)+1
			Else
				dx2 = dx1+nDisplayRecs*nRecRange-1
				dy2 = dy1+nRecRange-1
			End If
			While x <= nTotalRecs
				If x >= dx1 And x <= dx2 Then
					If CLng(nStartRec) = CLng(x) Then %>
		<b><%=y%></b>
					<%	Else %>
		<a href="Colorslist.asp?start=<%=x%>"><b><%=y%></b></a>
					<%	End If
					x = x + nDisplayRecs
					y = y + 1
				ElseIf x >= (dx1-nDisplayRecs*nRecRange) And x <= (dx2+nDisplayRecs*nRecRange) Then
					If x+nRecRange*nDisplayRecs < nTotalRecs Then %>
		<a href="Colorslist.asp?start=<%=x%>"><b><%=y%>-<%=y+nRecRange-1%></b></a>
					<% Else
						ny=(nTotalRecs-1)\nDisplayRecs+1
							If ny = y Then %>
		<a href="Colorslist.asp?start=<%=x%>"><b><%=y%></b></a>
							<% Else %>
		<a href="Colorslist.asp?start=<%=x%>"><b><%=y%>-<%=ny%></b></a>
							<%	End If
					End If
					x=x+nRecRange*nDisplayRecs
					y=y+nRecRange
				Else
					x=x+nRecRange*nDisplayRecs
					y=y+nRecRange
				End If
			Wend
		End If

		' Next link
		If NOT rsEof Then
			NextStart = nStartRec + nDisplayRecs
			isMore = True %>
		<a href="Colorslist.asp?start=<%=NextStart%>"><b>Next</b></a>
		<% Else
			isMore = False
		End If %>
		<br>	
<%	End If
	If CLng(nStartRec) > CLng(nTotalRecs) Then nStartRec = nTotalRecs
	nStopRec = nStartRec + nDisplayRecs - 1
	nRecCount = nTotalRecs - 1
	If rsEof Then nRecCount = nTotalRecs
	If nStopRec > nRecCount Then nStopRec = nRecCount %>
	Records <%= nStartRec %> to <%= nStopRec %> of <%= nTotalRecs %>
<% Else %>
	No records found
<% End If %>
</span>
		</td>
<% If nTotalRecs > 0 Then %>
		<td nowrap>&nbsp;&nbsp;&nbsp;&nbsp;</td>
		<td align="right" valign="top" nowrap><span class="aspmaker">Records Per Page&nbsp;
<select name="RecPerPage" onChange="this.form.submit();" class="aspmaker">
<option value="20"<% If nDisplayRecs = 20 Then response.write " selected" %>>20</option>
<option value="30"<% If nDisplayRecs = 30 Then response.write " selected" %>>30</option>
<option value="50"<% If nDisplayRecs = 50 Then response.write " selected" %>>50</option>
</select>
		</span></td>
<% End If %>
	</tr>
</table>
</form>	
<form method="post">


<table border="0" cellpadding="2" cellspacing="0" width="308" id="table1">
	<tr>
		<td height="28" width="129">&nbsp;</td>
		<td width="22" align="center" height="28">				
<span class="aspmaker">
		De<b>l</b></span></td>
		<td  bgcolor="#3366CC" width="145" height="28"><span class="aspmaker" style="color: #FFFFFF;"><a href="Colorslist.asp?order=<%= Server.URLEncode("Name") %>" style="color: #FFFFFF;">Name<% If Session("Colors_x_Name_Sort") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session("Colors_Name_Sort") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
			</span></td>
	</tr>
	



<%

' Avoid starting record > total records
If CLng(nStartRec) > CLng(nTotalRecs) Then
	nStartRec = nTotalRecs
End If

' Set the last record to display
nStopRec = nStartRec + nDisplayRecs - 1

' Move to first record directly for performance reason
nRecCount = nStartRec - 1
If Not rs.Eof Then
	rs.MoveFirst
	rs.Move nStartRec - 1
End If
nRecActual = 0
Do While (Not rs.Eof) And (nRecCount < nStopRec)
	nRecCount = nRecCount + 1
	If CLng(nRecCount) >= CLng(nStartRec) Then 
		nRecActual = nRecActual + 1

	' Set row color
	sItemRowClass = " bgcolor=""#FFFFFF"""

	' Display alternate color for rows
	If nRecCount Mod 2 <> 0 Then
		sItemRowClass = " bgcolor=""#FFFFFF"""
	End If

		' Load Key for record
		sKey = rs("Color")
		x_Color = rs("Color")
		x_Name = rs("Name")
%>
<% If (nRecActual Mod nRecPerRow = 1) OR (nRecPerRow < 2) Then %>
<tr>  
<% End If %>  
<td width="129" <td valign="top" <%=sItemRowClass%>><span class="aspmaker">
<a href="<% If Not IsNull(sKey) Then Response.Write "Colorsedit.asp?key=" & Server.URLEncode(sKey) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Edit</a>&nbsp;&nbsp;
<a href="<% If Not IsNull(sKey) Then Response.Write "Colorsadd.asp?key=" & Server.URLEncode(sKey) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Copy</a>
&nbsp;&nbsp;<a href="<% If Not IsNull(sKey) Then Response.Write "Imagescolor.asp?key=" & Server.URLEncode(sKey) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Image</a>
</span></td>
<td align="center"><input type="checkbox" name="key_d" value="<%= sKey %>" class="aspmaker"></td>
<td><span class="aspmaker"><% Response.Write x_Name %></span></td>
<% If (nRecActual Mod nRecPerRow = 0) Or (nRecPerRow < 2) Then %>  
</tr>
<% End If %>
<%
	End If
	rs.MoveNext
Loop
%>
<% If (nRecActual Mod nRecPerRow) <> 0 Then
	For i = 1 to (nRecPerRow - nRecActual Mod nRecPerRow) %>  
	<td>&nbsp;</td>
	<% Next %>
	</tr>  
<% End If %>  
</table>
<% If nRecActual > 0 Then %>
<p><input type="button" name="btndelete" value="DELETE SELECTED" onClick="this.form.action='Colorsdelete.asp';this.form.encoding='application/x-www-form-urlencoded';this.form.submit();"></p>
<% End If %>
</form>
<%

' Close recordset and connection
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>
<!--#include file="footer.asp"-->
<%

'-------------------------------------------------------------------------------
' Function SetUpDisplayRecs
' - Set up Number of Records displayed per page based on Form element RecPerPage
' - Variables setup: nDisplayRecs

Sub SetUpDisplayRecs()
	Dim sWrk
	sWrk = Request.QueryString("RecPerPage")
	If sWrk <> "" Then
		If IsNumeric(sWrk) Then
			nDisplayRecs = CInt(sWrk)
		Else
			If UCase(sWrk) = "ALL" Then ' Display All Records
				nDisplayRecs = -1
			Else
				nDisplayRecs = 20 ' Non-numeric, Load Default
			End If
		End If
		Session("Colors_RecPerPage") = nDisplayRecs ' Save to Session

		' Reset Start Position (Reset Command)
		nStartRec = 1
		Session("Colors_REC") = nStartRec
	Else
		If Session("Colors_RecPerPage") <> "" Then
			nDisplayRecs = Session("Colors_RecPerPage") ' Restore from Session
		Else
			nDisplayRecs = 20 ' Load Default
		End If
	End If
End Sub

'-------------------------------------------------------------------------------
' Function SetUpSortOrder
' - Set up Sort parameters based on Sort Links clicked
' - Variables setup: sOrderBy, Session("Table_OrderBy"), Session("Table_Field_Sort")

Sub SetUpSortOrder()
	Dim sOrder, sSortField, sLastSort, sThisSort
	Dim bCtrl

	' Check for an Order parameter
	If Request.QueryString("order").Count > 0 Then
		sOrder = Request.QueryString("order")

		' Field Name
		If sOrder = "Name" Then
			sSortField = "[Name]"
			sLastSort = Session("Colors_x_Name_Sort")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session("Colors_x_Name_Sort") = sThisSort
		Else
			If Session("Colors_x_Name_Sort") <> "" Then Session("Colors_x_Name_Sort") = ""
		End If
		Session("Colors_OrderBy") = sSortField & " " & sThisSort
		Session("Colors_REC") = 1
	End If
	sOrderBy = Session("Colors_OrderBy")
	If sOrderBy = "" Then
		sOrderBy = sDefaultOrderBy
		Session("Colors_OrderBy") = sOrderBy
	End If
End Sub

'-------------------------------------------------------------------------------
' Function SetUpStartRec
' - Set up Starting Record parameters based on Pager Navigation
' - Variables setup: nStartRec

Sub SetUpStartRec()
	Dim nPageNo

	' Check for a START parameter
	If Request.QueryString("start").Count > 0 Then
		nStartRec = Request.QueryString("start")
		Session("Colors_REC") = nStartRec
	ElseIf Request.QueryString("pageno").Count > 0 Then
		nPageNo = Request.QueryString("pageno")
		If IsNumeric(nPageNo) Then
			nStartRec = (nPageNo-1)*nDisplayRecs+1
			If nStartRec <= 0 Then
				nStartRec = 1
			ElseIf nStartRec >= ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1 Then
				nStartRec = ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1
			End If
			Session("Colors_REC") = nStartRec
		Else
			nStartRec = Session("Colors_REC")
			If Not IsNumeric(nStartRec) Or nStartRec = "" Then			
				nStartRec = 1 ' Reset start record counter
				Session("Colors_REC") = nStartRec
			End If
		End If
	Else
		nStartRec = Session("Colors_REC")
		If Not IsNumeric(nStartRec) Or nStartRec = "" Then		
			nStartRec = 1 'Reset start record counter
			Session("Colors_REC") = nStartRec
		End If
	End If
End Sub

'-------------------------------------------------------------------------------
' Function ResetCmd
' - Clear list page parameters
' - RESET: reset search parameters
' - RESETALL: reset search & master/detail parameters
' - RESETSORT: reset sort parameters

Sub ResetCmd()
	Dim sCmd

	' Get Reset Cmd
	If Request.QueryString("cmd").Count > 0 Then
		sCmd = Request.QueryString("cmd")

		' Reset Search Criteria
		If UCase(sCmd) = "RESET" Then
			sSrchWhere = ""
			Session("Colors_searchwhere") = sSrchWhere

		' Reset Search Criteria & Session Keys
		ElseIf UCase(sCmd) = "RESETALL" Then
			sSrchWhere = ""
			Session("Colors_searchwhere") = sSrchWhere

		' Reset Sort Criteria
		ElseIf UCase(sCmd) = "RESETSORT" Then
			sOrderBy = ""
			Session("Colors_OrderBy") = sOrderBy
			If Session("Colors_x_Name_Sort") <> "" Then Session("Colors_x_Name_Sort") = ""
		End If

		' Reset Start Position (Reset Command)
		nStartRec = 1
		Session("Colors_REC") = nStartRec
	End If
End Sub
%>