<!--#include file="ewconfig.asp"-->
<!--#include file="db.asp"-->
<!--#include file="Image_Groupsinfo.asp"-->
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
x_Group = Null: ox_Group = Null: z_Group = Null
%>
<%
sExport = Request.QueryString("export") ' Load Export Request
If sExport = "html" Then

	' Printer Friendly
End If
If sExport = "excel" Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "Content-Disposition", "attachment; filename=" & ewTblVar & ".xls"
End If
If sExport = "csv" Then
	Response.ContentType = "application/csv"
	Response.AddHeader "Content-Disposition:", "attachment; filename=" & ewTblVar & ".csv"
End If
%>
<%
nStartRec = 0
nStopRec = 0
nTotalRecs = 0
nRecCount = 0
nRecActual = 0
sDbWhereMaster = ""
sDbWhereDetail = ""
sSrchAdvanced = ""
psearch = ""
psearchtype = ""
sSrchBasic = ""
sSrchWhere = ""
sDbWhere = ""
sOrderBy = ""
sSqlMaster = ""
bMasterRecordExist = False
x_Dir = Null: ox_Dir = Null
x_Code = Null: ox_Code = Null
x_Client = Null: ox_Client = Null
x_Range = Null: ox_Range = Null
x_Description = Null: ox_Description = Null
x_Color = Null: ox_Color = Null
x_Width = Null: ox_Width = Null
x_Height = Null: ox_Height = Null
x_Depth = Null: ox_Depth = Null
x_Price = Null: ox_Price = Null
x_Edge = Null: ox_Edge = Null
x_Sync_info = Null: ox_Sync_info = Null
x_Sync_img = Null: ox_Sync_img = Null
x_Info = Null: ox_Info = Null
nDisplayRecs = 25
nRecRange = 10

' Set up records per page dynamically
SetUpDisplayRecs()

' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str

' Handle Reset Command
ResetCmd()

' Set Up Master Detail Parameters
SetUpMasterDetail()

' Build Filter condition
sDbWhere = ""
If sDbWhereDetail <> "" Then
	If sDbWhere <> "" Then sDbWhere = sDbWhere & " AND "
	sDbWhere = sDbWhere & "(" & sDbWhereDetail & ")"
End If
If sSrchWhere <> "" Then
	If sDbWhere <> "" Then sDbWhere = sDbWhere & " AND "
	sDbWhere = sDbWhere & "(" & sSrchWhere & ")"
End If

' Set Up Sorting Order
sOrderBy = ""
SetUpSortOrder()

' Set up SQL
sSql = ewBuildSql(ewSqlSelect, ewSqlWhere, ewSqlGroupBy, ewSqlHaving, ewSqlOrderBy, sDbWhere, sOrderBy)

'Response.Write sSql ' Uncomment to show SQL for debugging
' Build Master Record SQL

If sDbWhereMaster <> "" Then
	sSqlMaster = ewBuildSql(ewSqlMasterSelect, ewSqlMasterWhere, ewSqlMasterGroupBy, ewSqlMasterHaving, ewSqlMasterOrderBy, sDbWhereMaster, "")
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open sSqlMaster, conn
	bMasterRecordExist = (Not rs.Eof)
	If Not bMasterRecordExist Then
		Session(ewSessionTblMasterWhere) = ""
		Session(ewSessionTblDetailWhere) = ""
		Session(ewSessionMessage) = "No records found"
		conn.Close ' Close Connection
		Set conn = Nothing
		Response.Redirect "Imageslist.asp"
	End If
End If

' Export Data only
If sExport = "xml" Or sExport = "csv" Then
	Call ExportData(sExport, sSql)
	conn.Close ' Close Connection
	Set conn = Nothing
	Response.End
End If
%>
<% If sExport <> "word" And sExport <> "excel" Then %>
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
<script type="text/javascript">
<!--
EW_dateSep = "/"; // set date separator	
//-->
</script>
<script type="text/javascript">
<!--
var firstrowoffset = 1; // first data row start at
var tablename = 'ewlistmain'; // table name
var usecss = true; // use css
//var usecss = false; // use css
var rowclass = 'ewTableRow'; // row class
var rowaltclass = 'ewTableAltRow'; // row alternate class
var rowmoverclass = 'ewTableHighlightRow'; // row mouse over class
var rowselectedclass = 'ewTableSelectRow'; // row selected class
var roweditclass = 'ewTableEditRow'; // row edit class
var rowcolor = '#FFFFFF'; // row color
var rowaltcolor = '#F5F5F5'; // row alternate color
var rowmovercolor = '#FFCCFF'; // row mouse over color
var rowselectedcolor = '#CCFFFF'; // row selected color
var roweditcolor = '#FFFF99'; // row edit color
//-->
</script>
<script type="text/javascript">
<!--
	var EW_DHTMLEditors = [];
//-->
</script>
<% End If %>
<%
If sDbWhereMaster <> "" Then
	If bMasterRecordExist Then %>
<p><span class="aspmaker">Master Record: Images
<% If sExport = "" Then %>
<br><a href="Imageslist.asp">Back to Master Page</a></span>
<% End If %>
</p>
<table class="ewTable">
	<tr class="ewTableHeader">
		<td valign="top" style="white-space: nowrap;"><span>Image</span></td>
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
	<tr class="ewTableSelectRow">
<%
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
%>
		<td style="white-space: nowrap;"><span>
<% sTmp = x_Image %><% If Not IsNull(sTmp) Then %><a href="Imagesview.asp?Image=<%= sTmp %>"><img src="thumbs/<%= x_Image %>.jpg" Border=0></a><% Else %><img src="thumbs/<%= x_Image %>.jpg" Border=0><% End If %>
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
</table>
<br>
<%
	End If
	rs.Close
	Set rs = Nothing
End If
%>
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
<p><span class="aspmaker">TABLE: Image Groups
<% If sExport = "" Then %>
&nbsp;&nbsp;<a href="Image_Groupslist.asp?export=html">Printer Friendly</a>
&nbsp;&nbsp;<a href="Image_Groupslist.asp?export=excel">Export to Excel</a>
&nbsp;&nbsp;<a href="Image_Groupslist.asp?export=csv">Export to CSV</a>
<% End If %>
</span></p>
<% If sExport = "" Then %>
<table class="ewListAdd">
	<tr>
		<td><span class="aspmaker"><a href="Image_Groupsadd.asp">Add</a></span></td>
	</tr>
</table>
<p>
<% End If %>
<%
If Session(ewSessionMessage) <> "" Then
%>
<p><span class="ewmsg"><%= Session(ewSessionMessage) %></span></p>
<%
	Session(ewSessionMessage) = "" ' Clear message
End If
%>
<% If sExport = "" Then %>
<form action="Image_Groupslist.asp" name="ewpagerform" id="ewpagerform">
<table border="0" cellspacing="0" cellpadding="0">
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
		<a href="Image_Groupslist.asp?start=<%=PrevStart%>"><b>Prev</b></a>
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
		<a href="Image_Groupslist.asp?start=<%=x%>"><b><%=y%></b></a>
					<%	End If
					x = x + nDisplayRecs
					y = y + 1
				ElseIf x >= (dx1-nDisplayRecs*nRecRange) And x <= (dx2+nDisplayRecs*nRecRange) Then
					If x+nRecRange*nDisplayRecs < nTotalRecs Then %>
		<a href="Image_Groupslist.asp?start=<%=x%>"><b><%=y%>-<%=y+nRecRange-1%></b></a>
					<% Else
						ny=(nTotalRecs-1)\nDisplayRecs+1
							If ny = y Then %>
		<a href="Image_Groupslist.asp?start=<%=x%>"><b><%=y%></b></a>
							<% Else %>
		<a href="Image_Groupslist.asp?start=<%=x%>"><b><%=y%>-<%=ny%></b></a>
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
		<a href="Image_Groupslist.asp?start=<%=NextStart%>"><b>Next</b></a>
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
	<% If sSrchWhere = "0=101" Then %>
	<% Else %>
	No records found
	<% End If %>
<% End If %>
</span>
		</td>
<% If nTotalRecs > 0 Then %>
		<td nowrap>&nbsp;&nbsp;&nbsp;&nbsp;</td>
		<td align="right" valign="top" nowrap><span class="aspmaker">Records Per Page&nbsp;
<select name="<%= ewTblRecPerPage %>" onChange="this.form.submit();" class="aspmaker">
<option value="25"<% If nDisplayRecs = 25 Then response.write " selected" %>>25</option>
<option value="50"<% If nDisplayRecs = 50 Then response.write " selected" %>>50</option>
<option value="100"<% If nDisplayRecs = 100 Then response.write " selected" %>>100</option>
<option value="200"<% If nDisplayRecs = 200 Then response.write " selected" %>>200</option>
<option value="500"<% If nDisplayRecs = 500 Then response.write " selected" %>>500</option>
<option value="ALL"<% If Session(ewSessionTblRecPerPage) = -1 Then response.write " selected" %>>All Records</option>
</select>
		</span></td>
<% End If %>
	</tr>
</table>
</form>
<% End If %>
<% If nTotalRecs > 0 Then %>
<form method="post">
<table id="ewlistmain" class="ewTable">
	<!-- Table header -->
	<tr class="ewTableHeader">
		<td valign="top"><span>
<% If sExport <> "" Then %>
Group
<% Else %>
	<a href="Image_Groupslist.asp?order=<%= Server.URLEncode("Group") %>">Group<% If Session(ewSessionTblSort & "_x_Group") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session(ewSessionTblSort & "_x_Group") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</span></td>
<% If sExport = "" Then %>
<td>&nbsp;</td>
<% End If %>
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
	sItemRowClass = " class=""ewTableRow"""
	sListTrJs = " onmouseover='ew_mouseover(this);' onmouseout='ew_mouseout(this);' onclick='ew_click(this);'"

	' Display alternate color for rows
	If nRecCount Mod 2 <> 1 Then
		sItemRowClass = " class=""ewTableAltRow"""
	End If
	x_Image = rs("Image")
	x_Group = rs("Group")
%>
	<!-- Table body -->
	<tr<%=sItemRowClass%><%=sListTrJs%>>
		<!-- Group -->
		<td><span>
<%
If Not IsNull(x_Group) Then
	sSqlWrk = "SELECT [Name], [Index] FROM [Color_Groups]"
	sTmp = x_Group
	sTmp2 = ""
	sSqlWrk = sSqlWrk & " WHERE [Index] = " & AdjustSql(sTmp) & ""
	sSqlWrk = sSqlWrk & " ORDER BY [Name] Asc"
	Set rswrk = conn.Execute(sSqlWrk)
	If Not rswrk.Eof Then
		sTmp = rswrk("Name")
	End If
	If Not rswrk.Eof Then
		sTmp2 = rswrk("Index")
	End If
	rswrk.Close
	Set rswrk = Nothing
Else
	sTmp = Null
End If
ox_Group = x_Group ' Backup Original Value
x_Group = sTmp
%>
<a href="Colors_Grouplist.asp?showmaster=1&Grp=<%=sTmp2%>"><% Response.Write x_Group %></a>
<% x_Group = ox_Group ' Restore Original Value %>
</span></td>
<% If sExport = "" Then %>
<td><span class="aspmaker"><a href="<% If Not IsNull(x_Image) AND Not IsNull(x_Group) Then Response.Write "Image_Groupsdelete.asp?Image=" & Server.URLEncode(x_Image) & "&Group=" & Server.URLEncode(x_Group) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Delete</a></span></td>
<% End If %>
	</tr>
<%
	End If
	rs.MoveNext
Loop
%>
</table>
</form>
<% End If %>
<%

' Close recordset and connection
rs.Close
Set rs = Nothing
conn.Close
Set conn = Nothing
%>
<% If sExport = "" Then %>
<% End If %>
<% If sExport <> "word" And sExport <> "excel" Then %>
<!--#include file="footer.asp"-->
<% End If %>
<%

'-------------------------------------------------------------------------------
' Function SetUpDisplayRecs
' - Set up Number of Records displayed per page based on Form element RecPerPage
' - Variables setup: nDisplayRecs

Sub SetUpDisplayRecs()
	Dim sWrk
	sWrk = Request.QueryString(ewTblRecPerPage)
	If sWrk <> "" Then
		If IsNumeric(sWrk) Then
			nDisplayRecs = CInt(sWrk)
		Else
			If LCase(sWrk) = "all" Then ' Display All Records
				nDisplayRecs = -1
			Else
				nDisplayRecs = 25 ' Non-numeric, Load Default
			End If
		End If
		Session(ewSessionTblRecPerPage) = nDisplayRecs ' Save to Session

		' Reset Start Position (Reset Command)
		nStartRec = 1
		Session(ewSessionTblStartRec) = nStartRec
	Else
		If Session(ewSessionTblRecPerPage) <> "" Then
			nDisplayRecs = Session(ewSessionTblRecPerPage) ' Restore from Session
		Else
			nDisplayRecs = 25 ' Load Default
		End If
	End If
End Sub

'-------------------------------------------------------------------------------
' Function SetUpMasterDetail
' - Set up Master Detail criteria based on querystring parameter showmaster
' - Variables setup: sDbWhereMaster, Session(TblMasterWhere), sDbWhereDetail, Session(TblDetailWhere)

Sub SetUpMasterDetail()

	' Get the keys for master table
	If Request.QueryString(ewTblShowMaster).Count > 0 Then

		' Reset start record counter (new master key)
		nStartRec = 1
		Session(ewSessionTblStartRec) = nStartRec
		sDbWhereMaster = ewSqlMasterFilter
		sDbWhereDetail = ewSqlDetailFilter
		x_Image = Request.QueryString("Image") ' Load Parameter from QueryString
		If IsNumeric(x_Image) Then
		sDbWhereMaster = Replace(sDbWhereMaster, "@Image", AdjustSql(x_Image)) ' Replace key value
		sDbWhereDetail = Replace(sDbWhereDetail, "@Image", AdjustSql(x_Image)) ' Replace key value
		Session(ewSessionTblMasterKey & "_Image") = x_Image ' Save Master Key Value
		Else
		sDbWhereMaster = "0=1":sDbWhereDetail = "0=1"
		End If
		Session(ewSessionTblMasterWhere) = sDbWhereMaster
		Session(ewSessionTblDetailWhere) = sDbWhereDetail
	Else
		sDbWhereMaster = Session(ewSessionTblMasterWhere)
		sDbWhereDetail = Session(ewSessionTblDetailWhere)
	End If
End Sub

'-------------------------------------------------------------------------------
' Function ResetSearch
' - Clear all search parameters
'

Sub ResetSearch()

	' Clear search where
	sSrchWhere = ""
	Session(ewSessionTblSearchWhere) = sSrchWhere

	' Clear advanced search parameters
	Session(ewSessionTblAdvSrch & "_x_Image") = ""
	Session(ewSessionTblAdvSrch & "_x_Group") = ""
	Session(ewSessionTblBasicSrch) = ""
	Session(ewSessionTblBasicSrchType) = ""
End Sub

'-------------------------------------------------------------------------------
' Function RestoreSearch
' - Restore all search parameters
'

Sub RestoreSearch()

	' Restore advanced search settings
	x_Image = Session(ewSessionTblAdvSrch & "_x_Image")
	x_Group = Session(ewSessionTblAdvSrch & "_x_Group")
	psearch = Session(ewSessionTblBasicSrch)
	psearchtype = Session(ewSessionTblBasicSrchType)
End Sub

'-------------------------------------------------------------------------------
' Function SetUpSortOrder
' - Set up Sort parameters based on Sort Links clicked
' - Variables setup: sOrderBy, Session(TblOrderBy), Session(Tbl_Field_Sort)

Sub SetUpSortOrder()
	Dim sOrder, sSortField, sLastSort, sThisSort
	Dim bCtrl

	' Check for an Order parameter
	If Request.QueryString("order").Count > 0 Then
		sOrder = Request.QueryString("order")

		' Field [Group]
		If sOrder = "Group" Then
			sSortField = "[Group]"
			sLastSort = Session(ewSessionTblSort & "_x_Group")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session(ewSessionTblSort & "_x_Group") = sThisSort
		Else
			If Session(ewSessionTblSort & "_x_Group") <> "" Then Session(ewSessionTblSort & "_x_Group") = ""
		End If
		Session(ewSessionTblOrderBy) = sSortField & " " & sThisSort
		Session(ewSessionTblStartRec) = 1
	End If
	sOrderBy = Session(ewSessionTblOrderBy)
	If sOrderBy = "" Then
		sOrderBy = ewSqlOrderBy
		Session(ewSessionTblOrderBy) = sOrderBy
		If sOrderBy <> "" Then
			Dim arOrderBy, i
			arOrderBy = Split(ewSqlOrderBySessions, ",")
			For i = 0 to UBound(arOrderBy)\2
				Session(ewSessionTblSort & "_" & arOrderBy(i*2)) = arOrderBy(i*2+1)
			Next
		End If
	End If
End Sub

'-------------------------------------------------------------------------------
' Function SetUpStartRec
' - Set up Starting Record parameters based on Pager Navigation
' - Variables setup: nStartRec

Sub SetUpStartRec()
	Dim nPageNo

	' Check for a START parameter
	If Request.QueryString(ewTblStartRec).Count > 0 Then
		nStartRec = Request.QueryString(ewTblStartRec)
		Session(ewSessionTblStartRec) = nStartRec
	ElseIf Request.QueryString("pageno").Count > 0 Then
		nPageNo = Request.QueryString("pageno")
		If IsNumeric(nPageNo) Then
			nStartRec = (nPageNo-1)*nDisplayRecs+1
			If nStartRec <= 0 Then
				nStartRec = 1
			ElseIf nStartRec >= ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1 Then
				nStartRec = ((nTotalRecs-1)\nDisplayRecs)*nDisplayRecs+1
			End If
			Session(ewSessionTblStartRec) = nStartRec
		Else
			nStartRec = Session(ewSessionTblStartRec)
			If Not IsNumeric(nStartRec) Or nStartRec = "" Then
				nStartRec = 1 ' Reset start record counter
				Session(ewSessionTblStartRec) = nStartRec
			End If
		End If
	Else
		nStartRec = Session(ewSessionTblStartRec)
		If Not IsNumeric(nStartRec) Or nStartRec = "" Then
			nStartRec = 1 'Reset start record counter
			Session(ewSessionTblStartRec) = nStartRec
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
		If LCase(sCmd) = "reset" Then
			Call ResetSearch()

		' Reset Search Criteria & Session Keys
		ElseIf LCase(sCmd) = "resetall" Then
			Call ResetSearch()
			Session(ewSessionTblMasterWhere) = "" ' Clear master criteria
			sDbWhereMaster = ""
			Session(ewSessionTblDetailWhere) = "" ' Clear detail criteria
			sDbWhereDetail = ""
		Session(ewSessionTblMasterKey & "_Image") = "" ' Clear Master Key Value

		' Reset Sort Criteria
		ElseIf LCase(sCmd) = "resetsort" Then
			sOrderBy = ""
			Session(ewSessionTblOrderBy) = sOrderBy
			If Session(ewSessionTblSort & "_x_Group") <> "" Then Session(ewSessionTblSort & "_x_Group") = ""
		End If

		' Reset Start Position (Reset Command)
		nStartRec = 1
		Session(ewSessionTblStartRec) = nStartRec
	End If
End Sub

'-------------------------------------------------------------------------------
' Function ExportData
' - Export Data in Xml or Csv format

Sub ExportData(sExport, sSql)
	Dim oXmlDoc, oXmlTbl, oXmlRec, oXmlFld
	Dim sCsvStr
	Dim rs

	' Set up Record Set
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation = 3
	rs.Open sSql, conn, 1, 2
	nTotalRecs = rs.RecordCount
	nStartRec = 1
	SetUpStartRec() ' Set Up Start Record Position
	If sExport = "csv" Then
		sCsvStr = sCsvStr & """Group""" & ","
		sCsvStr = Left(sCsvStr, Len(sCsvStr)-1) ' Remove last comma
		sCsvStr = sCsvStr & vbCrLf
	End If

	' Avoid starting record > total records
	If CLng(nStartRec) > CLng(nTotalRecs) Then
		nStartRec = nTotalRecs
	End If

	' Set the last record to display
	If nDisplayRecs < 0 Then
		nStopRec = nTotalRecs
	Else
		nStopRec = nStartRec + nDisplayRecs - 1
	End If

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
			x_Group = rs("Group")
			If sExport = "csv" Then

				' Field Group
				sCsvStr = sCsvStr & """" & Replace(x_Group&"","""","""""") & """" & ","
				sCsvStr = Left(sCsvStr, Len(sCsvStr)-1) ' Remove last comma
				sCsvStr = sCsvStr & vbCrLf
			End If
		End If
		rs.MoveNext
	Loop

	' Close recordset and connection
	rs.Close
	Set rs = Nothing
	If sExport = "csv" Then
		Response.Write sCsvStr
	End If
End Sub
%>