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
nDisplayRecs = 25
nRecRange = 10

' Set up records per page dynamically
SetUpDisplayRecs()

' Open connection to the database
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open xDb_Conn_Str

' Handle Reset Command
ResetCmd()

' Get Search Criteria for Advanced Search
SetUpAdvancedSearch()

' Get Search Criteria for Basic Search
SetUpBasicSearch()

' Build Search Criteria
If sSrchAdvanced <> "" Then
	If sSrchWhere <> "" Then sSrchWhere = sSrchWhere & " AND "
	sSrchWhere = sSrchWhere & "(" & sSrchAdvanced & ")"
End If
If sSrchBasic <> "" Then
	If sSrchWhere <> "" Then sSrchWhere = sSrchWhere & " AND "
	sSrchWhere = sSrchWhere & "(" & sSrchBasic & ")"
End If

' Save Search Criteria
If sSrchWhere <> "" Then
	Session(ewSessionTblSearchWhere) = sSrchWhere
	nStartRec = 1 ' reset start record counter
	Session(ewSessionTblStartRec) = nStartRec
Else
	sSrchWhere = Session(ewSessionTblSearchWhere)
	Call RestoreSearch()
End If

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
	ew_clickall(elem);
}
function EW_selected(elem) {
	var f = elem.form;	
	if (!f.key_d) return false;
	if (f.key_d[0]) {
		for (var i=0; i<f.key_d.length; i++)
			if (f.key_d[i].checked) return true;
	} else {
		return f.key_d.checked;
	}
	return false;
}
//-->
</script>
<% End If %>
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
<p><span class="aspmaker">TABLE: Images
<% If sExport = "" Then %>
&nbsp;&nbsp;<a href="Imageslist.asp?export=html">Printer Friendly</a>
&nbsp;&nbsp;<a href="Imageslist.asp?export=excel">Export to Excel</a>
&nbsp;&nbsp;<a href="Imageslist.asp?export=csv">Export to CSV</a>
<% End If %>
</span></p>
<% If sExport = "" Then %>
<form id="fImageslistsrch" name="fImageslistsrch" action="Imageslist.asp" >
<table class="ewBasicSearch">
	<tr>
		<td><span class="aspmaker">
			<input type="text" name="<%=ewTblBasicSrch%>" size="20" value="<%=psearch%>">
			<input type="Submit" name="Submit" value="Search&nbsp;(*)">&nbsp;<input type="Button" name="Reset" value="Reset" onclick="clearForm(this.form);this.form.<%=ewTblBasicSrchType%>[0].checked = true;">&nbsp;
			<a href="Imageslist.asp?cmd=reset">Show all</a>&nbsp;
			<a href="Imagessrch.asp">Advanced Search</a>
		</span></td>
	</tr>
	<tr>
	<td><span class="aspmaker"><input type="radio" name="<%=ewTblBasicSrchType%>" value="" <% If psearchtype = "" Then %>checked<% End If %>>Exact phrase&nbsp;&nbsp;<input type="radio" name="<%=ewTblBasicSrchType%>" value="AND" <% If psearchtype = "AND" Then %>checked<% End If %>>All words&nbsp;&nbsp;<input type="radio" name="<%=ewTblBasicSrchType%>" value="OR" <% If psearchtype = "OR" Then %>checked<% End If %>>Any word</span></td>
	</tr>
</table>
</form>
<% End If %>
<% If sExport = "" Then %>
<table class="ewListAdd">
	<tr>
		<td><span class="aspmaker"><a href="Imagesadd.asp">Add</a></span></td>
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
<form action="Imageslist.asp" name="ewpagerform" id="ewpagerform">
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
		<a href="Imageslist.asp?start=<%=PrevStart%>"><b>Prev</b></a>
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
		<a href="Imageslist.asp?start=<%=x%>"><b><%=y%></b></a>
					<%	End If
					x = x + nDisplayRecs
					y = y + 1
				ElseIf x >= (dx1-nDisplayRecs*nRecRange) And x <= (dx2+nDisplayRecs*nRecRange) Then
					If x+nRecRange*nDisplayRecs < nTotalRecs Then %>
		<a href="Imageslist.asp?start=<%=x%>"><b><%=y%>-<%=y+nRecRange-1%></b></a>
					<% Else
						ny=(nTotalRecs-1)\nDisplayRecs+1
							If ny = y Then %>
		<a href="Imageslist.asp?start=<%=x%>"><b><%=y%></b></a>
							<% Else %>
		<a href="Imageslist.asp?start=<%=x%>"><b><%=y%>-<%=ny%></b></a>
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
		<a href="Imageslist.asp?start=<%=NextStart%>"><b>Next</b></a>
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
		<td valign="top" style="white-space: nowrap;"><span>
<% If sExport <> "" Then %>
Image
<% Else %>
	<a href="Imageslist.asp?order=<%= Server.URLEncode("Image") %>">Image<% If Session(ewSessionTblSort & "_x_Image") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session(ewSessionTblSort & "_x_Image") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</span></td>
		<td valign="top"><span>
<% If sExport <> "" Then %>
Code
<% Else %>
	<a href="Imageslist.asp?order=<%= Server.URLEncode("Code") %>">Code&nbsp;(*)<% If Session(ewSessionTblSort & "_x_Code") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session(ewSessionTblSort & "_x_Code") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</span></td>
		<td valign="top"><span>
<% If sExport <> "" Then %>
Client
<% Else %>
	<a href="Imageslist.asp?order=<%= Server.URLEncode("Client") %>">Client&nbsp;(*)<% If Session(ewSessionTblSort & "_x_Client") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session(ewSessionTblSort & "_x_Client") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</span></td>
		<td valign="top"><span>
<% If sExport <> "" Then %>
Range
<% Else %>
	<a href="Imageslist.asp?order=<%= Server.URLEncode("Range") %>">Range<% If Session(ewSessionTblSort & "_x_Range") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session(ewSessionTblSort & "_x_Range") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</span></td>
		<td valign="top"><span>
<% If sExport <> "" Then %>
Description
<% Else %>
	<a href="Imageslist.asp?order=<%= Server.URLEncode("Description") %>">Description<% If Session(ewSessionTblSort & "_x_Description") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session(ewSessionTblSort & "_x_Description") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</span></td>
		<td valign="top"><span>
<% If sExport <> "" Then %>
Color
<% Else %>
	<a href="Imageslist.asp?order=<%= Server.URLEncode("Color") %>">Color<% If Session(ewSessionTblSort & "_x_Color") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session(ewSessionTblSort & "_x_Color") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</span></td>
		<td valign="top"><span>
<% If sExport <> "" Then %>
Width
<% Else %>
	<a href="Imageslist.asp?order=<%= Server.URLEncode("Width") %>">Width<% If Session(ewSessionTblSort & "_x_Width") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session(ewSessionTblSort & "_x_Width") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</span></td>
		<td valign="top"><span>
<% If sExport <> "" Then %>
Height
<% Else %>
	<a href="Imageslist.asp?order=<%= Server.URLEncode("Height") %>">Height<% If Session(ewSessionTblSort & "_x_Height") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session(ewSessionTblSort & "_x_Height") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</span></td>
		<td valign="top"><span>
<% If sExport <> "" Then %>
Depth
<% Else %>
	<a href="Imageslist.asp?order=<%= Server.URLEncode("Depth") %>">Depth<% If Session(ewSessionTblSort & "_x_Depth") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session(ewSessionTblSort & "_x_Depth") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</span></td>
		<td valign="top"><span>
<% If sExport <> "" Then %>
Price
<% Else %>
	<a href="Imageslist.asp?order=<%= Server.URLEncode("Price") %>">Price<% If Session(ewSessionTblSort & "_x_Price") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session(ewSessionTblSort & "_x_Price") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</span></td>
		<td valign="top"><span>
<% If sExport <> "" Then %>
Edge
<% Else %>
	<a href="Imageslist.asp?order=<%= Server.URLEncode("Edge") %>">Edge&nbsp;(*)<% If Session(ewSessionTblSort & "_x_Edge") = "ASC" Then %><img src="images/sortup.gif" width="10" height="9" border="0"><% ElseIf Session(ewSessionTblSort & "_x_Edge") = "DESC" Then %><img src="images/sortdown.gif" width="10" height="9" border="0"><% End If %></a>
<% End If %>
		</span></td>
<% If sExport = "" Then %>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td>&nbsp;</td>
<td><input type="checkbox" class="aspmaker" onClick="EW_selectKey(this);"></td>
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
	<!-- Table body -->
	<tr<%=sItemRowClass%><%=sListTrJs%>>
		<!-- Image -->
		<td style="white-space: nowrap;"><span>
<% sTmp = x_Image %><% If Not IsNull(sTmp) Then %><a href="Imagesview.asp?Image=<%= sTmp %>"><img src="thumbs/<%= x_Image %>.jpg" Border=0></a><% Else %><img src="thumbs/<%= x_Image %>.jpg" Border=0><% End If %>
</span></td>
		<!-- Code -->
		<td><span>
<% Response.Write x_Code %>
</span></td>
		<!-- Client -->
		<td><span>
<% Response.Write x_Client %>
</span></td>
		<!-- Range -->
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
<% sTmp = x_Range %><% If Not IsNull(sTmp) Then %><a href="Imageslist.asp?x_Range=<%=x_Range %>"><% Response.Write x_Range %></a><% Else %><% Response.Write x_Range %><% End If %>
<% x_Range = ox_Range ' Restore Original Value %>
</span></td>
		<!-- Description -->
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
<% sTmp = x_Description %><% If Not IsNull(sTmp) Then %><a href="Imageslist.asp?x_Description=<%= x_Description %>"><% Response.Write x_Description %></a><% Else %><% Response.Write x_Description %><% End If %>
<% x_Description = ox_Description ' Restore Original Value %>
</span></td>
		<!-- Color -->
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
		<!-- Width -->
		<td><span>
<% Response.Write x_Width %>
</span></td>
		<!-- Height -->
		<td><span>
<% Response.Write x_Height %>
</span></td>
		<!-- Depth -->
		<td><span>
<% Response.Write x_Depth %>
</span></td>
		<!-- Price -->
		<td><span>
<% Response.Write x_Price %>
</span></td>
		<!-- Edge -->
		<td><span>
<% Response.Write x_Edge %>
</span></td>
<% If sExport = "" Then %>
<td><span class="aspmaker"><a href="<% If Not IsNull(x_Image) Then Response.Write "Imagesview.asp?Image=" & Server.URLEncode(x_Image) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">View</a></span></td>
<td><span class="aspmaker"><a href="<% If Not IsNull(x_Image) Then Response.Write "Imagesedit.asp?Image=" & Server.URLEncode(x_Image) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Edit</a></span></td>
<td><span class="aspmaker"><a href="<% If Not IsNull(x_Image) Then Response.Write "Imagesadd.asp?Image=" & Server.URLEncode(x_Image) Else Response.Write "javascript:alert('Invalid Record! Key is null');" End If %>">Copy</a></span></td>
<td><span class="aspmaker"><input type="checkbox" name="key_d" value="<%= x_Image %>" class="aspmaker" onclick='ew_clickmultidelete(this);'>Delete</span></td>
<td><span class="aspmaker"><a href="Image_Groupslist.asp?showmaster=1&Image=<%= Server.URLEncode(x_Image)%>">Image Groups Details</a></span></td>
<% End If %>
	</tr>
<%
	End If
	rs.MoveNext
Loop
%>
</table>
<% If sExport = "" Then %>
<% If nRecActual > 0 Then %>
<p>
<input type="button" name="btndelete" value="DELETE SELECTED" onClick="if (!EW_selected(this)) alert('No records selected'); else {this.form.action='Imagesdelete.asp';this.form.encoding='application/x-www-form-urlencoded';this.form.submit();}">
<p>
<% End If %>
<% End If %>
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
' Function SetUpAdvancedSearch
' - Set up Advanced Search parameter based on querystring parameters from Advanced Search Page
' - Variables setup: sSrchAdvanced

Sub SetUpAdvancedSearch()
	Dim arrFldOpr, arrFldOpr2, sSrchStr

	' Field Image
	sSrchStr = ""
	x_Image = Request.QueryString("x_Image")
	z_Image = Request.QueryString("z_Image")
	arrFldOpr = Split(z_Image, ",")
	If x_Image <> "" And IsNumeric(x_Image) And IsValidOpr(arrFldOpr) Then
		sSrchStr = sSrchStr & "[Image] " & arrFldOpr(0) & " " & _
			arrFldOpr(1) & AdjustSql(x_Image) & arrFldOpr(2)
	End If
	If sSrchStr <> "" Then
		If sSrchAdvanced <> "" Then sSrchAdvanced = sSrchAdvanced & " AND "
		sSrchAdvanced = sSrchAdvanced & "(" & sSrchStr & ")"
	End If

	' Field Dir
	sSrchStr = ""
	x_Dir = Request.QueryString("x_Dir")
	z_Dir = Request.QueryString("z_Dir")
	arrFldOpr = Split(z_Dir, ",")
	If x_Dir <> "" And IsNumeric(x_Dir) And IsValidOpr(arrFldOpr) Then
		sSrchStr = sSrchStr & "[Dir] " & arrFldOpr(0) & " " & _
			arrFldOpr(1) & AdjustSql(x_Dir) & arrFldOpr(2)
	End If
	If sSrchStr <> "" Then
		If sSrchAdvanced <> "" Then sSrchAdvanced = sSrchAdvanced & " AND "
		sSrchAdvanced = sSrchAdvanced & "(" & sSrchStr & ")"
	End If

	' Field Code
	sSrchStr = ""
	x_Code = Request.QueryString("x_Code")
	z_Code = Request.QueryString("z_Code")
	arrFldOpr = Split(z_Code, ",")
	If x_Code <> "" And IsValidOpr(arrFldOpr) Then
		sSrchStr = sSrchStr & "[Code] " & arrFldOpr(0) & " " & _
			arrFldOpr(1) & AdjustSql(x_Code) & arrFldOpr(2)
	End If
	If sSrchStr <> "" Then
		If sSrchAdvanced <> "" Then sSrchAdvanced = sSrchAdvanced & " AND "
		sSrchAdvanced = sSrchAdvanced & "(" & sSrchStr & ")"
	End If

	' Field Client
	sSrchStr = ""
	x_Client = Request.QueryString("x_Client")
	z_Client = Request.QueryString("z_Client")
	arrFldOpr = Split(z_Client, ",")
	If x_Client <> "" And IsValidOpr(arrFldOpr) Then
		sSrchStr = sSrchStr & "[Client] " & arrFldOpr(0) & " " & _
			arrFldOpr(1) & AdjustSql(x_Client) & arrFldOpr(2)
	End If
	If sSrchStr <> "" Then
		If sSrchAdvanced <> "" Then sSrchAdvanced = sSrchAdvanced & " AND "
		sSrchAdvanced = sSrchAdvanced & "(" & sSrchStr & ")"
	End If

	' Field Range
	sSrchStr = ""
	x_Range = Request.QueryString("x_Range")
	z_Range = Request.QueryString("z_Range")
	arrFldOpr = Split(z_Range, ",")
	If x_Range <> "" And IsNumeric(x_Range) And IsValidOpr(arrFldOpr) Then
		sSrchStr = sSrchStr & "[Range] " & arrFldOpr(0) & " " & _
			arrFldOpr(1) & AdjustSql(x_Range) & arrFldOpr(2)
	End If
	If sSrchStr <> "" Then
		If sSrchAdvanced <> "" Then sSrchAdvanced = sSrchAdvanced & " AND "
		sSrchAdvanced = sSrchAdvanced & "(" & sSrchStr & ")"
	End If

	' Field Description
	sSrchStr = ""
	x_Description = Request.QueryString("x_Description")
	z_Description = Request.QueryString("z_Description")
	arrFldOpr = Split(z_Description, ",")
	If x_Description <> "" And IsNumeric(x_Description) And IsValidOpr(arrFldOpr) Then
		sSrchStr = sSrchStr & "[Description] " & arrFldOpr(0) & " " & _
			arrFldOpr(1) & AdjustSql(x_Description) & arrFldOpr(2)
	End If
	If sSrchStr <> "" Then
		If sSrchAdvanced <> "" Then sSrchAdvanced = sSrchAdvanced & " AND "
		sSrchAdvanced = sSrchAdvanced & "(" & sSrchStr & ")"
	End If

	' Field Color
	sSrchStr = ""
	x_Color = Request.QueryString("x_Color")
	z_Color = Request.QueryString("z_Color")
	arrFldOpr = Split(z_Color, ",")
	If x_Color <> "" And IsNumeric(x_Color) And IsValidOpr(arrFldOpr) Then
		sSrchStr = sSrchStr & "[Color] " & arrFldOpr(0) & " " & _
			arrFldOpr(1) & AdjustSql(x_Color) & arrFldOpr(2)
	End If
	If sSrchStr <> "" Then
		If sSrchAdvanced <> "" Then sSrchAdvanced = sSrchAdvanced & " AND "
		sSrchAdvanced = sSrchAdvanced & "(" & sSrchStr & ")"
	End If

	' Field Width
	sSrchStr = ""
	x_Width = Request.QueryString("x_Width")
	z_Width = Request.QueryString("z_Width")
	arrFldOpr = Split(z_Width, ",")
	If x_Width <> "" And IsNumeric(x_Width) And IsValidOpr(arrFldOpr) Then
		sSrchStr = sSrchStr & "[Width] " & arrFldOpr(0) & " " & _
			arrFldOpr(1) & AdjustSql(x_Width) & arrFldOpr(2)
	End If
	If sSrchStr <> "" Then
		If sSrchAdvanced <> "" Then sSrchAdvanced = sSrchAdvanced & " AND "
		sSrchAdvanced = sSrchAdvanced & "(" & sSrchStr & ")"
	End If

	' Field Height
	sSrchStr = ""
	x_Height = Request.QueryString("x_Height")
	z_Height = Request.QueryString("z_Height")
	arrFldOpr = Split(z_Height, ",")
	If x_Height <> "" And IsNumeric(x_Height) And IsValidOpr(arrFldOpr) Then
		sSrchStr = sSrchStr & "[Height] " & arrFldOpr(0) & " " & _
			arrFldOpr(1) & AdjustSql(x_Height) & arrFldOpr(2)
	End If
	If sSrchStr <> "" Then
		If sSrchAdvanced <> "" Then sSrchAdvanced = sSrchAdvanced & " AND "
		sSrchAdvanced = sSrchAdvanced & "(" & sSrchStr & ")"
	End If

	' Field Depth
	sSrchStr = ""
	x_Depth = Request.QueryString("x_Depth")
	z_Depth = Request.QueryString("z_Depth")
	arrFldOpr = Split(z_Depth, ",")
	If x_Depth <> "" And IsNumeric(x_Depth) And IsValidOpr(arrFldOpr) Then
		sSrchStr = sSrchStr & "[Depth] " & arrFldOpr(0) & " " & _
			arrFldOpr(1) & AdjustSql(x_Depth) & arrFldOpr(2)
	End If
	If sSrchStr <> "" Then
		If sSrchAdvanced <> "" Then sSrchAdvanced = sSrchAdvanced & " AND "
		sSrchAdvanced = sSrchAdvanced & "(" & sSrchStr & ")"
	End If

	' Field Price
	sSrchStr = ""
	x_Price = Request.QueryString("x_Price")
	z_Price = Request.QueryString("z_Price")
	arrFldOpr = Split(z_Price, ",")
	If x_Price <> "" And IsNumeric(x_Price) And IsValidOpr(arrFldOpr) Then
		sSrchStr = sSrchStr & "[Price] " & arrFldOpr(0) & " " & _
			arrFldOpr(1) & AdjustSql(x_Price) & arrFldOpr(2)
	End If
	If sSrchStr <> "" Then
		If sSrchAdvanced <> "" Then sSrchAdvanced = sSrchAdvanced & " AND "
		sSrchAdvanced = sSrchAdvanced & "(" & sSrchStr & ")"
	End If

	' Field Edge
	sSrchStr = ""
	x_Edge = Request.QueryString("x_Edge")
	z_Edge = Request.QueryString("z_Edge")
	arrFldOpr = Split(z_Edge, ",")
	If x_Edge <> "" And IsValidOpr(arrFldOpr) Then
		sSrchStr = sSrchStr & "[Edge] " & arrFldOpr(0) & " " & _
			arrFldOpr(1) & AdjustSql(x_Edge) & arrFldOpr(2)
	End If
	If sSrchStr <> "" Then
		If sSrchAdvanced <> "" Then sSrchAdvanced = sSrchAdvanced & " AND "
		sSrchAdvanced = sSrchAdvanced & "(" & sSrchStr & ")"
	End If

	' Field Sync_info
	sSrchStr = ""
	x_Sync_info = Request.QueryString("x_Sync_info")
	z_Sync_info = Request.QueryString("z_Sync_info")
	arrFldOpr = Split(z_Sync_info, ",")
	If x_Sync_info <> "" And IsValidOpr(arrFldOpr) Then
		sSrchStr = sSrchStr & "[Sync_info] " & arrFldOpr(0) & " " & _
			arrFldOpr(1) & AdjustSql(x_Sync_info) & arrFldOpr(2)
	End If
	If sSrchStr <> "" Then
		If sSrchAdvanced <> "" Then sSrchAdvanced = sSrchAdvanced & " AND "
		sSrchAdvanced = sSrchAdvanced & "(" & sSrchStr & ")"
	End If

	' Field Sync_img
	sSrchStr = ""
	x_Sync_img = Request.QueryString("x_Sync_img")
	z_Sync_img = Request.QueryString("z_Sync_img")
	arrFldOpr = Split(z_Sync_img, ",")
	If x_Sync_img <> "" And IsValidOpr(arrFldOpr) Then
		sSrchStr = sSrchStr & "[Sync_img] " & arrFldOpr(0) & " " & _
			arrFldOpr(1) & AdjustSql(x_Sync_img) & arrFldOpr(2)
	End If
	If sSrchStr <> "" Then
		If sSrchAdvanced <> "" Then sSrchAdvanced = sSrchAdvanced & " AND "
		sSrchAdvanced = sSrchAdvanced & "(" & sSrchStr & ")"
	End If

	' Field Info
	sSrchStr = ""
	x_Info = Request.QueryString("x_Info")
	z_Info = Request.QueryString("z_Info")
	arrFldOpr = Split(z_Info, ",")
	If x_Info <> "" And IsValidOpr(arrFldOpr) Then
		sSrchStr = sSrchStr & "[Info] " & arrFldOpr(0) & " " & _
			arrFldOpr(1) & AdjustSql(x_Info) & arrFldOpr(2)
	End If
	If sSrchStr <> "" Then
		If sSrchAdvanced <> "" Then sSrchAdvanced = sSrchAdvanced & " AND "
		sSrchAdvanced = sSrchAdvanced & "(" & sSrchStr & ")"
	End If
	If sSrchAdvanced <> "" Then ' save settings
		Session(ewSessionTblAdvSrch & "_x_Image") = x_Image
		Session(ewSessionTblAdvSrch & "_x_Dir") = x_Dir
		Session(ewSessionTblAdvSrch & "_x_Code") = x_Code
		Session(ewSessionTblAdvSrch & "_x_Client") = x_Client
		Session(ewSessionTblAdvSrch & "_x_Range") = x_Range
		Session(ewSessionTblAdvSrch & "_x_Description") = x_Description
		Session(ewSessionTblAdvSrch & "_x_Color") = x_Color
		Session(ewSessionTblAdvSrch & "_x_Width") = x_Width
		Session(ewSessionTblAdvSrch & "_x_Height") = x_Height
		Session(ewSessionTblAdvSrch & "_x_Depth") = x_Depth
		Session(ewSessionTblAdvSrch & "_x_Price") = x_Price
		Session(ewSessionTblAdvSrch & "_x_Edge") = x_Edge
		Session(ewSessionTblAdvSrch & "_x_Sync_info") = x_Sync_info
		Session(ewSessionTblAdvSrch & "_x_Sync_img") = x_Sync_img
		Session(ewSessionTblAdvSrch & "_x_Info") = x_Info
	End If
End Sub

' Function to check if the search operators are valid
Function IsValidOpr(arOpr)
	Dim Opr
	IsValidOpr = IsArray(arOpr)
	If IsValidOpr Then IsValidOpr = (UBound(arOpr) >= 2)
	If IsValidOpr Then
		For Each Opr In arOpr
			Opr = UCase(Trim(Opr))
				If Not (Opr = "=" Or Opr = "<" Or Opr = "<=" Or _
				Opr = ">" Or Opr = ">=" Or Opr = "<>" Or _
				Opr = "LIKE" Or Opr = "NOT LIKE" Or Opr = "BETWEEN" Or _
				Opr = "'" Or Opr = "'%" Or Opr = "%'" Or Opr = "#" Or Opr = "") Then
					IsValidOpr = False
					Exit For
			End If
		Next
	End If
End Function

'-------------------------------------------------------------------------------
' Function BasicSearchSQL
' - Build WHERE clause for a keyword

Function BasicSearchSQL(Keyword)
	Dim sKeyword
	sKeyword = AdjustSql(Keyword)
	BasicSearchSQL = ""
	BasicSearchSQL = BasicSearchSQL & "[Code] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[Client] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[Edge] LIKE '%" & sKeyword & "%' OR "
	BasicSearchSQL = BasicSearchSQL & "[Info] LIKE '%" & sKeyword & "%' OR "
	If Right(BasicSearchSQL, 4) = " OR " Then BasicSearchSQL = Left(BasicSearchSQL, Len(BasicSearchSQL)-4)
End Function

'-------------------------------------------------------------------------------
' Function SetUpBasicSearch
' - Set up Basic Search parameter based on form elements pSearch & pSearchType
' - Variables setup: sSrchBasic

Sub SetUpBasicSearch()
	Dim arKeyword, sKeyword
	psearch = Request.QueryString(ewTblBasicSrch)
	psearchtype = Request.QueryString(ewTblBasicSrchType)
	If psearch <> "" Then
		If psearchtype <> "" Then
			While InStr(psearch, "  ") > 0
				sSearch = Replace(psearch, "  ", " ")
			Wend
			arKeyword = Split(Trim(psearch), " ")
			For Each sKeyword In arKeyword
				sSrchBasic = sSrchBasic & "(" & BasicSearchSQL(sKeyword) & ") " & psearchtype & " "
			Next
		Else
			sSrchBasic = BasicSearchSQL(psearch)
		End If
	End If
	If Right(sSrchBasic, 4) = " OR " Then sSrchBasic = Left(sSrchBasic, Len(sSrchBasic)-4)
	If Right(sSrchBasic, 5) = " AND " Then sSrchBasic = Left(sSrchBasic, Len(sSrchBasic)-5)
	If psearch <> "" then
		Session(ewSessionTblBasicSrch) = psearch
		Session(ewSessionTblBasicSrchType) = psearchtype
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
	Session(ewSessionTblAdvSrch & "_x_Dir") = ""
	Session(ewSessionTblAdvSrch & "_x_Code") = ""
	Session(ewSessionTblAdvSrch & "_x_Client") = ""
	Session(ewSessionTblAdvSrch & "_x_Range") = ""
	Session(ewSessionTblAdvSrch & "_x_Description") = ""
	Session(ewSessionTblAdvSrch & "_x_Color") = ""
	Session(ewSessionTblAdvSrch & "_x_Width") = ""
	Session(ewSessionTblAdvSrch & "_x_Height") = ""
	Session(ewSessionTblAdvSrch & "_x_Depth") = ""
	Session(ewSessionTblAdvSrch & "_x_Price") = ""
	Session(ewSessionTblAdvSrch & "_x_Edge") = ""
	Session(ewSessionTblAdvSrch & "_x_Sync_info") = ""
	Session(ewSessionTblAdvSrch & "_x_Sync_img") = ""
	Session(ewSessionTblAdvSrch & "_x_Info") = ""
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
	x_Dir = Session(ewSessionTblAdvSrch & "_x_Dir")
	x_Code = Session(ewSessionTblAdvSrch & "_x_Code")
	x_Client = Session(ewSessionTblAdvSrch & "_x_Client")
	x_Range = Session(ewSessionTblAdvSrch & "_x_Range")
	x_Description = Session(ewSessionTblAdvSrch & "_x_Description")
	x_Color = Session(ewSessionTblAdvSrch & "_x_Color")
	x_Width = Session(ewSessionTblAdvSrch & "_x_Width")
	x_Height = Session(ewSessionTblAdvSrch & "_x_Height")
	x_Depth = Session(ewSessionTblAdvSrch & "_x_Depth")
	x_Price = Session(ewSessionTblAdvSrch & "_x_Price")
	x_Edge = Session(ewSessionTblAdvSrch & "_x_Edge")
	x_Sync_info = Session(ewSessionTblAdvSrch & "_x_Sync_info")
	x_Sync_img = Session(ewSessionTblAdvSrch & "_x_Sync_img")
	x_Info = Session(ewSessionTblAdvSrch & "_x_Info")
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

		' Field [Image]
		If sOrder = "Image" Then
			sSortField = "[Image]"
			sLastSort = Session(ewSessionTblSort & "_x_Image")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session(ewSessionTblSort & "_x_Image") = sThisSort
		Else
			If Session(ewSessionTblSort & "_x_Image") <> "" Then Session(ewSessionTblSort & "_x_Image") = ""
		End If

		' Field [Code]
		If sOrder = "Code" Then
			sSortField = "[Code]"
			sLastSort = Session(ewSessionTblSort & "_x_Code")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session(ewSessionTblSort & "_x_Code") = sThisSort
		Else
			If Session(ewSessionTblSort & "_x_Code") <> "" Then Session(ewSessionTblSort & "_x_Code") = ""
		End If

		' Field [Client]
		If sOrder = "Client" Then
			sSortField = "[Client]"
			sLastSort = Session(ewSessionTblSort & "_x_Client")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session(ewSessionTblSort & "_x_Client") = sThisSort
		Else
			If Session(ewSessionTblSort & "_x_Client") <> "" Then Session(ewSessionTblSort & "_x_Client") = ""
		End If

		' Field [Range]
		If sOrder = "Range" Then
			sSortField = "[Range]"
			sLastSort = Session(ewSessionTblSort & "_x_Range")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session(ewSessionTblSort & "_x_Range") = sThisSort
		Else
			If Session(ewSessionTblSort & "_x_Range") <> "" Then Session(ewSessionTblSort & "_x_Range") = ""
		End If

		' Field [Description]
		If sOrder = "Description" Then
			sSortField = "[Description]"
			sLastSort = Session(ewSessionTblSort & "_x_Description")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session(ewSessionTblSort & "_x_Description") = sThisSort
		Else
			If Session(ewSessionTblSort & "_x_Description") <> "" Then Session(ewSessionTblSort & "_x_Description") = ""
		End If

		' Field [Color]
		If sOrder = "Color" Then
			sSortField = "[Color]"
			sLastSort = Session(ewSessionTblSort & "_x_Color")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session(ewSessionTblSort & "_x_Color") = sThisSort
		Else
			If Session(ewSessionTblSort & "_x_Color") <> "" Then Session(ewSessionTblSort & "_x_Color") = ""
		End If

		' Field [Width]
		If sOrder = "Width" Then
			sSortField = "[Width]"
			sLastSort = Session(ewSessionTblSort & "_x_Width")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session(ewSessionTblSort & "_x_Width") = sThisSort
		Else
			If Session(ewSessionTblSort & "_x_Width") <> "" Then Session(ewSessionTblSort & "_x_Width") = ""
		End If

		' Field [Height]
		If sOrder = "Height" Then
			sSortField = "[Height]"
			sLastSort = Session(ewSessionTblSort & "_x_Height")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session(ewSessionTblSort & "_x_Height") = sThisSort
		Else
			If Session(ewSessionTblSort & "_x_Height") <> "" Then Session(ewSessionTblSort & "_x_Height") = ""
		End If

		' Field [Depth]
		If sOrder = "Depth" Then
			sSortField = "[Depth]"
			sLastSort = Session(ewSessionTblSort & "_x_Depth")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session(ewSessionTblSort & "_x_Depth") = sThisSort
		Else
			If Session(ewSessionTblSort & "_x_Depth") <> "" Then Session(ewSessionTblSort & "_x_Depth") = ""
		End If

		' Field [Price]
		If sOrder = "Price" Then
			sSortField = "[Price]"
			sLastSort = Session(ewSessionTblSort & "_x_Price")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session(ewSessionTblSort & "_x_Price") = sThisSort
		Else
			If Session(ewSessionTblSort & "_x_Price") <> "" Then Session(ewSessionTblSort & "_x_Price") = ""
		End If

		' Field [Edge]
		If sOrder = "Edge" Then
			sSortField = "[Edge]"
			sLastSort = Session(ewSessionTblSort & "_x_Edge")
			If sLastSort = "ASC" Then sThisSort = "DESC" Else sThisSort = "ASC"
			Session(ewSessionTblSort & "_x_Edge") = sThisSort
		Else
			If Session(ewSessionTblSort & "_x_Edge") <> "" Then Session(ewSessionTblSort & "_x_Edge") = ""
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

		' Reset Sort Criteria
		ElseIf LCase(sCmd) = "resetsort" Then
			sOrderBy = ""
			Session(ewSessionTblOrderBy) = sOrderBy
			If Session(ewSessionTblSort & "_x_Image") <> "" Then Session(ewSessionTblSort & "_x_Image") = ""
			If Session(ewSessionTblSort & "_x_Code") <> "" Then Session(ewSessionTblSort & "_x_Code") = ""
			If Session(ewSessionTblSort & "_x_Client") <> "" Then Session(ewSessionTblSort & "_x_Client") = ""
			If Session(ewSessionTblSort & "_x_Range") <> "" Then Session(ewSessionTblSort & "_x_Range") = ""
			If Session(ewSessionTblSort & "_x_Description") <> "" Then Session(ewSessionTblSort & "_x_Description") = ""
			If Session(ewSessionTblSort & "_x_Color") <> "" Then Session(ewSessionTblSort & "_x_Color") = ""
			If Session(ewSessionTblSort & "_x_Width") <> "" Then Session(ewSessionTblSort & "_x_Width") = ""
			If Session(ewSessionTblSort & "_x_Height") <> "" Then Session(ewSessionTblSort & "_x_Height") = ""
			If Session(ewSessionTblSort & "_x_Depth") <> "" Then Session(ewSessionTblSort & "_x_Depth") = ""
			If Session(ewSessionTblSort & "_x_Price") <> "" Then Session(ewSessionTblSort & "_x_Price") = ""
			If Session(ewSessionTblSort & "_x_Edge") <> "" Then Session(ewSessionTblSort & "_x_Edge") = ""
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
		sCsvStr = sCsvStr & """Image""" & ","
		sCsvStr = sCsvStr & """Code""" & ","
		sCsvStr = sCsvStr & """Client""" & ","
		sCsvStr = sCsvStr & """Range""" & ","
		sCsvStr = sCsvStr & """Description""" & ","
		sCsvStr = sCsvStr & """Color""" & ","
		sCsvStr = sCsvStr & """Width""" & ","
		sCsvStr = sCsvStr & """Height""" & ","
		sCsvStr = sCsvStr & """Depth""" & ","
		sCsvStr = sCsvStr & """Price""" & ","
		sCsvStr = sCsvStr & """Edge""" & ","
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
			x_Image = rs("Image")
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
			If sExport = "csv" Then

				' Field Image
				sCsvStr = sCsvStr & """" & Replace(x_Image&"","""","""""") & """" & ","

				' Field Code
				sCsvStr = sCsvStr & """" & Replace(x_Code&"","""","""""") & """" & ","

				' Field Client
				sCsvStr = sCsvStr & """" & Replace(x_Client&"","""","""""") & """" & ","

				' Field Range
				sCsvStr = sCsvStr & """" & Replace(x_Range&"","""","""""") & """" & ","

				' Field Description
				sCsvStr = sCsvStr & """" & Replace(x_Description&"","""","""""") & """" & ","

				' Field Color
				sCsvStr = sCsvStr & """" & Replace(x_Color&"","""","""""") & """" & ","

				' Field Width
				sCsvStr = sCsvStr & """" & Replace(x_Width&"","""","""""") & """" & ","

				' Field Height
				sCsvStr = sCsvStr & """" & Replace(x_Height&"","""","""""") & """" & ","

				' Field Depth
				sCsvStr = sCsvStr & """" & Replace(x_Depth&"","""","""""") & """" & ","

				' Field Price
				sCsvStr = sCsvStr & """" & Replace(x_Price&"","""","""""") & """" & ","

				' Field Edge
				sCsvStr = sCsvStr & """" & Replace(x_Edge&"","""","""""") & """" & ","
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