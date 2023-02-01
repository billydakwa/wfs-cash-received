<%
strTableName = "##@TABLE.strDataSourceTable s##"
##if @TABLE.strOriginalTable##
strOriginalTableName = "##@TABLE.strOriginalTable s##"
##else##
strOriginalTableName = "##@TABLE.strDataSourceTable s##"
##endif##
##if @TABLE.nType==titTABLE || @TABLE.nType==titVIEW##
gPageSize = ##@TABLE.nNumberOfRecords##
##endif##

gstrOrderBy = "##@TABLE.strOrderBy s##"
##if @TABLE.nType==titTABLE || @TABLE.nType==titVIEW || @TABLE.nType==titREPORT##
##foreach @TABLE.arrOrderIndexes as @o order @o.nOrderIndex##
##if @first##
Set g_orderindexes = CreateDictionary()
##endif##
Set orderindex=CreateDictionary()
orderindex(0)=##@o.nIndex##
orderindex(1)=IIF(##@o.bAsc##,"ASC","DESC")
orderindex(2)="##@o.strOrderField s##"
g_orderindexes.Add g_orderindexes.Count, orderindex
##endfor##
##endif##

gsqlHead = "##@TABLE.sqlHead ls##"
gsqlFrom = "##@TABLE.sqlFrom ls##"
gsqlWhereExpr = "##@TABLE.sqlWhere ls##"
gsqlTail = "##@TABLE.sqlTail ls##"
%>
