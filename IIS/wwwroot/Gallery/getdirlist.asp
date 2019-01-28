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
x_Dir = Null
x_Name = Null
x_Description = Null
x_UDir = Null
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

' Get Search Criteria for Basic Search
SetUpBasicSearch()

' Build Search Criteria
If sSrchAdvanced <> "" Then
	sSrchWhere = sSrchAdvanced ' Advanced Search
ElseIf sSrchBasic <> "" Then
	sSrchWhere = sSrchBasic ' Basic Search
End If

' Save Search Criteria
If sSrchWhere <> "" Then
	Session("getdir_searchwhere") = sSrchWhere

	' Reset start record counter (new search)
	nStartRec = 1
	Session("getdir_REC") = nStartRec
Else
	sSrchWhere = Session("getdir_searchwhere")
End If

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
sSql = "SELECT * FROM [getdir]"

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
<p><span class="aspmaker">VIEW: getdir
</span></p>
<form action="getdirlist.asp">
<table border="0" cellspacing="0" cellpadding="0">
	<tr>
		<td><span class="aspmaker">
			<input type="text" name="psearch" size="20">
			<input type="Submit" name="Submit" value="Search &nbsp;(*)">&nbsp;&nbsp;
			<a href="getdirlist.asp?cmd=reset">Show all</a>&nbsp;&nbsp;
		</span></td>
	</tr>
	<tr><td><span class="aspmaker"><input type="radio" name="psearchtype" value="" checked>Exact phrase&nbsp;&nbsp;<input type="radio" name="psearchtype" value="AND">All words&nbsp;&nbsp;<input type="radio" name="psearchtype" value="OR">Any word</span></td></tr>	
</table>
</form>
<%
If Session("ewmsg") <> "" Then
%>
<p><span class="aspmaker" style="color: Red;"><%= Session("ewmsg") %></span></p>
<%
	Session("ewmsg") = "" ' Clear message
End If
%>
<form action="getdirlist.asp" name="ewpagerform" id="ewpagerform">
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
		<a href="getdirlist.asp?start=<%=PrevStart%>"><b>Prev</b></a>
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
		<a href="getdirlist.asp?start=<%=x%>"><b><%=y%></b></a>
					<%	End If
					x = x + nDisplayRecs
					y = y + 1
				ElseIf x >= (dx1-nDisplayRecs*nRecRange) And x <= (dx2+nDisplayRecs*nRecRange) Then
					If x+nRecRange*nDisplayRecs < nTotalRecs Then %>
		<a href="getdirlist.asp?start=<%=x%>"><b><%=y%>-<%=y+nRecRange-1%></b></a>
					<% Else
						ny=(nTotalRecs-1)\nDisplayRecs+1
							If ny = y Then %>
		<a href="getdirlist.asp?start=<%=x%>"><b><%=y%></b></a>
							<% Else %>
		<a href="getdirlist.asp?start=<%=x%>"><b><%=y%>-<%=ny%></b></a>
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
		<a href="getdirlist.asp?start=<%=NextStart%>"><b>Next</b></a>
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
<table border="0" cellspacing="5" cellpadding="5">
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
		x_Dir = rs("Dir")
		x_Name = rs("Name")
		x_Description = rs("Description")
		x_UDir = rs("UDir")
%>
<% If (nRecActual Mod nRecPerRow = 1) OR (nRecPerRow < 2) Then %>
<tr>  
<% End If %>  
	<td valign="top"<%=sItemRowClass%>>
	<table border="0" cellspacing="1" cellpadding="4" bgcolor="#CCCCCC">
		<tr>
			<td bgcolor="#3366CC"><span class="aspmaker" style="color: #FFFFFF;">
	<a href="getdirlist.asp?order=<%= Server.URLEncode("Dir") %>" style="color: #FFFFFF;">Dir<% If Session("getdir_x_Dir_Sort") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session("getdir_Dir_Sort") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
			</span></td>
			<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_Dir %>
</span></td>
		</tr>
		<tr>
			<td bgcolor="#3366CC"><span class="aspmaker" style="color: #FFFFFF;">
	<a href="getdirlist.asp?order=<%= Server.URLEncode("Name") %>" style="color: #FFFFFF;">Name&nbsp;(*)<% If Session("getdir_x_Name_Sort") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session("getdir_Name_Sort") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
			</span></td>
			<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_Name %>
</span></td>
		</tr>
		<tr>
			<td bgcolor="#3366CC"><span class="aspmaker" style="color: #FFFFFF;">
	<a href="getdirlist.asp?order=<%= Server.URLEncode("UDir") %>" style="color: #FFFFFF;">UDir<% If Session("getdir_x_UDir_Sort") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session("getdir_UDir_Sort") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
			</span></td>
			<td bgcolor="#FFFFFF"><span class="aspmaker">
<% Response.Write x_UDir %>
</span></td>
		</tr>
	</table>				
<span class="aspmaker">
</span>
	</td>
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
		Session("getdir_RecPerPage") = nDisplayRecs ' Save to Session

		' Reset Start Position (Reset Command)
		nStartRec = 1
		Session("getdir_REC") = nStartRec
	Else
		If Session("getdir_RecPerPage") <> "" Then
			nDisplayRecs = Session("getdir_RecPerPage") ' Restore from Session
		Else
			nDisplayRecs = 20 ' Load Default
		End If
	End If
End Sub

'-------------------------------------------------------------------------------
' Function BasicSearchSQL
' - Build WHERE clause for a keyword

Function BasicSearchSQL(Keyword)
	Dim sKeyword
	sKeyword = AdjustSql(Keyword)
	BasicSearchSQL = ""
	BasicSearchSQL = BasicSearchSQL & "[Name] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[Description] LIKE '%" & sKeyword & "%' OR "
	If Right(BasicSearchSQL, 4) = " OR " Then BasicSearchSQL = Left(BasicSearchSQL, Len(BasicSearchSQL)-4)
End Function

'-------------------------------------------------------------------------------
' Function SetUpBasicSearch
' - Set up Basic Search parameter based on form elements pSearch & pSearchType
' - Variables setup: sSrchBasic

Sub SetUpBasicSearch()
	Dim sSearch, sSearchType, arKeyword, sKeyword
	sSearch = Request.QueryString("psearch")
	sSearchType = Request.QueryString("psearchType")
	If sSearch <> "" Then
		If sSearchType <> "" Then
			While InStr(sSearch, "  ") > 0
				sSearch = Replace(sSearch, "  ", " ")
			Wend
			arKeyword = Split(Trim(sSearch), " ")
			For Each sKeyword In arKeyword
				sSrchBasic = sSrchBasic & "(" & BasicSearchSQL(sKeyword) & ") " & sSearchType & " "
			Next
		Else
			sSrchBasic = BasicSearchSQL(sSearch)
		End If
	End If
	If Right(sSrchBasic, 4) = " OR " Then sSrchBasic = Left(sSrchBasic, Len(sSrchBasic)-4)
	If Right(sSrchBasic, 5) = " AND " Then sSrchBasic = Left(sSrchBasic, Len(sSrchBasic)-5)
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

		' Field Dir
		If sOrder = "Dir" Then
			sSortField = "[Dir]"
			sLastSort = Session("getdir_x_Dir_Sort")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session("getdir_x_Dir_Sort") = sThisSort
		Else
			If Session("getdir_x_Dir_Sort") <> "" Then Session("getdir_x_Dir_Sort") = ""
		End If

		' Field Name
		If sOrder = "Name" Then
			sSortField = "[Name]"
			sLastSort = Session("getdir_x_Name_Sort")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session("getdir_x_Name_Sort") = sThisSort
		Else
			If Session("getdir_x_Name_Sort") <> "" Then Session("getdir_x_Name_Sort") = ""
		End If

		' Field UDir
		If sOrder = "UDir" Then
			sSortField = "[UDir]"
			sLastSort = Session("getdir_x_UDir_Sort")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session("getdir_x_UDir_Sort") = sThisSort
		Else
			If Session("getdir_x_UDir_Sort") <> "" Then Session("getdir_x_UDir_Sort") = ""
		End If
		Session("getdir_OrderBy") = sSortField & " " & sThisSort
		Session("getdir_REC") = 1
	End If
	sOrderBy = Session("getdir_OrderBy")
	If sOrderBy = "" Then
		sOrderBy = sDefaultOrderBy
		Session("getdir_OrderBy") = sOrderBy
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
		Session("getdir_REC") = nStartRec
	ElseIf Request.QueryString("pageno").Count > 0 Then
		nPageNo = Request.QueryString("pageno")
		If IsNumeric(nPageNo) Then
			nStartRec = (nPageNo-1)*nDisplayRecs+1
			If nStartRec <= 0 Then
				nStartRec = 1
			ElseIf nStartRec >= ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1 Then
				nStartRec = ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1
			End If
			Session("getdir_REC") = nStartRec
		Else
			nStartRec = Session("getdir_REC")
			If Not IsNumeric(nStartRec) Or nStartRec = "" Then			
				nStartRec = 1 ' Reset start record counter
				Session("getdir_REC") = nStartRec
			End If
		End If
	Else
		nStartRec = Session("getdir_REC")
		If Not IsNumeric(nStartRec) Or nStartRec = "" Then		
			nStartRec = 1 'Reset start record counter
			Session("getdir_REC") = nStartRec
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
			Session("getdir_searchwhere") = sSrchWhere

		' Reset Search Criteria & Session Keys
		ElseIf UCase(sCmd) = "RESETALL" Then
			sSrchWhere = ""
			Session("getdir_searchwhere") = sSrchWhere

		' Reset Sort Criteria
		ElseIf UCase(sCmd) = "RESETSORT" Then
			sOrderBy = ""
			Session("getdir_OrderBy") = sOrderBy
			If Session("getdir_x_Dir_Sort") <> "" Then Session("getdir_x_Dir_Sort") = ""
			If Session("getdir_x_Name_Sort") <> "" Then Session("getdir_x_Name_Sort") = ""
			If Session("getdir_x_UDir_Sort") <> "" Then Session("getdir_x_UDir_Sort") = ""
		End If

		' Reset Start Position (Reset Command)
		nStartRec = 1
		Session("getdir_REC") = nStartRec
	End If
End Sub
%>
