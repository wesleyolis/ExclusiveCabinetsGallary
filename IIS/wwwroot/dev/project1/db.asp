<%
xDb_Conn_Str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("Gallary.mdb") & ";"

' Function to Adjust SQL
Function AdjustSql(str)
	Dim sWrk
	sWrk = Trim(str & "")
	sWrk = Replace(sWrk, "'", "''") ' Adjust for Single Quote
	sWrk = Replace(sWrk, "[", "[[]") ' Adjust for Open Square Bracket
	AdjustSql = sWrk
End Function

' Function to Build SQL
Function ewBuildSql(sSelect, sWhere, sGroupBy, sHaving, sOrderBy, sFilter, sSort)
	Dim sSql, sDbWhere, sDbOrderBy
	sDbWhere = sWhere
	If sDbWhere <> "" Then
		sDbWhere = "(" & sDbWhere & ")"
	End If
	If sFilter <> "" Then
		If sDbWhere <> "" Then sDbWhere = sDbWhere & " AND "
		sDbWhere = sDbWhere & "(" & sFilter & ")"
	End If	
	sDbOrderBy = sOrderBy
	If sSort <> "" Then
		sDbOrderBy = sSort
	End If
	sSql = sSelect
	If sDbWhere <> "" Then
		sSql = sSql & " WHERE " & sDbWhere
	End If
	If sGroupBy <> "" Then
		sSql = sSql & " GROUP BY " & sGroupBy
	End If
	If sHaving <> "" Then
		sSql = sSql & " HAVING " & sHaving
	End If
	If sDbOrderBy <> "" Then
		sSql = sSql & " ORDER BY " & sDbOrderBy
	End If
	ewBuildSql = sSql
End Function
%>
