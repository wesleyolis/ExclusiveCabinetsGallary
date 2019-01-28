<%

' ASPMaker 5 configuration
' Table Level Constants

Const ewTblVar = "Image_Groups"
Const ewTblRecPerPage = "RecPerPage"
Const ewSessionTblRecPerPage = "Image_Groups_RecPerPage"
Const ewTblStartRec = "start"
Const ewSessionTblStartRec = "Image_Groups_start"
Const ewTblShowMaster = "showmaster"
Const ewSessionTblMasterKey = "Image_Groups_MasterKey"
Const ewSessionTblMasterWhere = "Image_Groups_MasterWhere"
Const ewSessionTblDetailWhere = "Image_Groups_DetailWhere"
Const ewSessionTblAdvSrch = "Image_Groups_AdvSrch"
Const ewTblBasicSrch = "psearch"
Const ewSessionTblBasicSrch = "Image_Groups_psearch"
Const ewTblBasicSrchType = "psearchtype"
Const ewSessionTblBasicSrchType = "Image_Groups_psearchtype"
Const ewSessionTblSearchWhere = "Image_Groups_SearchWhere"
Const ewSessionTblSort = "Image_Groups_Sort"
Const ewSessionTblOrderBy = "Image_Groups_OrderBy"
Const ewSessionTblKey = "Image_Groups_Key"

' Table Level SQL
Const ewSqlSelect = "SELECT * FROM [Image_Groups]"
Const ewSqlWhere = ""
Const ewSqlGroupBy = ""
Const ewSqlHaving = ""
Const ewSqlOrderBy = ""
Const ewSqlOrderBySessions = ""
Const ewSqlKeyWhere = "[Image] = @Image AND [Group] = @Group"
Const ewSqlMasterSelect = "SELECT * FROM [Images]"
Const ewSqlMasterWhere = ""
Const ewSqlMasterGroupBy = ""
Const ewSqlMasterHaving = ""
Const ewSqlMasterOrderBy = ""
Const ewSqlMasterFilter = "[Image] = @Image"
Const ewSqlDetailFilter = "[Image] = @Image"
Const ewSqlUserIDFilter = ""
%>
