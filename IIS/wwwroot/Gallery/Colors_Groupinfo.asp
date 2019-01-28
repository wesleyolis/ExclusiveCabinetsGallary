<%

' ASPMaker 5 configuration
' Table Level Constants

Const ewTblVar = "Colors_Group"
Const ewTblRecPerPage = "RecPerPage"
Const ewSessionTblRecPerPage = "Colors_Group_RecPerPage"
Const ewTblStartRec = "start"
Const ewSessionTblStartRec = "Colors_Group_start"
Const ewTblShowMaster = "showmaster"
Const ewSessionTblMasterKey = "Colors_Group_MasterKey"
Const ewSessionTblMasterWhere = "Colors_Group_MasterWhere"
Const ewSessionTblDetailWhere = "Colors_Group_DetailWhere"
Const ewSessionTblAdvSrch = "Colors_Group_AdvSrch"
Const ewTblBasicSrch = "psearch"
Const ewSessionTblBasicSrch = "Colors_Group_psearch"
Const ewTblBasicSrchType = "psearchtype"
Const ewSessionTblBasicSrchType = "Colors_Group_psearchtype"
Const ewSessionTblSearchWhere = "Colors_Group_SearchWhere"
Const ewSessionTblSort = "Colors_Group_Sort"
Const ewSessionTblOrderBy = "Colors_Group_OrderBy"
Const ewSessionTblKey = "Colors_Group_Key"

' Table Level SQL
Const ewSqlSelect = "SELECT * FROM [Colors_Group]"
Const ewSqlWhere = ""
Const ewSqlGroupBy = ""
Const ewSqlHaving = ""
Const ewSqlOrderBy = ""
Const ewSqlOrderBySessions = ""
Const ewSqlKeyWhere = "[Index] = @Index"
Const ewSqlMasterSelect = "SELECT * FROM [Color_Groups]"
Const ewSqlMasterWhere = ""
Const ewSqlMasterGroupBy = ""
Const ewSqlMasterHaving = ""
Const ewSqlMasterOrderBy = ""
Const ewSqlMasterFilter = "[Index] = @Grp"
Const ewSqlDetailFilter = "[Grp] = @Grp"
Const ewSqlUserIDFilter = ""
%>
