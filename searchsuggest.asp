<%
%>
<!--#include file="include/dbcommon.asp"-->
<%
add_nocache_headers 
asp_include "classes/searchclause.asp",false
doAssignmentByRef table,postvalue("table")
doAssignmentByRef strTableName,GetTableByShort(table)
if not bValue(checkTableName(table,false)) then
	Response.End
end if
asp_include ("include/" & CSmartStr(table)) & "_variables.asp",false
if IsEqual(postvalue("searchFor"),"") then
	response.end
end if
doAssignment sessionPrefix,strTableName
doAssignmentByRef allSearchFields,GetTableData(strTableName,".allSearchFields",CreateDictionary())
if not IsEmpty(Session(CSmartStr(sessionPrefix) & "_advsearch")) then
	doAssignmentByRef searchClauseObj,unserialize(Session(CSmartStr(sessionPrefix) & "_advsearch"))
else
	Set params = (CreateDictionary())
	setArrElement params,"tName",strTableName
	setArrElement params,"searchFieldsArr",allSearchFields
	setArrElement params,"sessionPrefix",sessionPrefix
	setArrElement params,"panelSearchFields",GetTableData(strTableName,".panelSearchFields",CreateDictionary())
	setArrElement params,"googleLikeFields",GetTableData(strTableName,".googleLikeFields",CreateDictionary())
	doClassAssignment this_object,"searchClauseObj",CreateClass("SearchClause",1,params,Empty,Empty,Empty,Empty,Empty,Empty)
end if
Set var_response = (CreateDictionary())
suggestAllContent = true
if bValue(postvalue("start")) then
	suggestAllContent = false
end if
doAssignmentByRef searchFor,postvalue("searchFor")
doAssignmentByRef searchField,GoodFieldName(postvalue("searchField"))
doAssignmentByRef strSecuritySql,SecuritySQL("Search",strTableName)
GetCollectionBounds allSearchFields,searchsuggest_loopIdx2,searchsuggest_loopMax2
searchsuggest_exitLoop2=false
do while searchsuggest_loopIdx2<=searchsuggest_loopMax2
	searchsuggest_exitLoop2=false
	do
		searchsuggest_arrKey2 = GetCollectionKey(allSearchFields,searchsuggest_loopIdx2)
		doAssignment f,ArrayElement(allSearchFields,searchsuggest_arrKey2)
		doAssignmentByRef fType,GetFieldType(f,strTableName)
		if not bValue(IsCharType(fType)) and not bValue(IsNumberType(fType)) or bValue(IsTextType(fType)) then
			exit do
		end if
		if (IsEqual(searchField,"") or IsEqual(searchField,GoodFieldName(f))) and bValue(CheckFieldPermissions(f,"")) then
			where = ""
			having = ""
			if not bValue(gQuery.IsAggrFuncField_p1(CSmartDbl(GetFieldIndex(f,""))-1)) then
				doAssignmentByRef where,searchClauseObj.getSuggestWhere_p4(f,fType,suggestAllContent,searchFor)
			else
				if bValue(gQuery.IsAggrFuncField_p1(CSmartDbl(GetFieldIndex(f,""))-1)) then
					doAssignmentByRef having,searchClauseObj.getSuggestWhere_p4(f,fType,suggestAllContent,searchFor)
				end if
			end if
			sqlHead = ("SELECT DISTINCT " & CSmartStr(GetFullFieldName(f,""))) & " "
			doAssignmentByRef oHaving,gQuery.Having()
			doAssignmentByRef sqlHaving,oHaving.toSql_p1(gQuery)
			doAssignmentByRef sqlGroupBy,gQuery.GroupByToSql()
			doAssignmentByRef where,whereAdd(where,strSecuritySql)
			doAssignmentByRef strSQL,gSQLWhere_having(sqlHead,gsqlFrom,gsqlWhereExpr,sqlGroupBy,sqlHaving,where,having)
			strSQL = CSmartStr(strSQL) & " ORDER BY 1 "
			doAssignmentByRef rs,db_query(strSQL,conn)
			i = 0
			do while bValue(doAssignmentByRef(row,db_fetch_numarray(rs)))
				i = CSmartDbl(i)+1
				doAssignmentByRef pos,asp_strpos(ArrayElement(row,0),vblf,empty)
				if not IsFalse(pos) then
					setArrElement var_response,asp_count(var_response),asp_substr(ArrayElement(row,0),0,pos)
				else
					setArrElement var_response,asp_count(var_response),ArrayElement(row,0)
				end if
				if IsLess(10,i) then
					exit do
				end if
			loop
		end if
	loop while false
	if searchsuggest_exitLoop2 then _
		exit do
	searchsuggest_loopIdx2=searchsuggest_loopIdx2+1
loop
db_close conn
doAssignmentByRef var_response,asp_array_unique(var_response)
asp_sort var_response
ResponseWrite "suggest_success"
i = 0
do while IsLess(i,10) and IsLess(i,asp_count(var_response))
	doAssignment value,ArrayElement(var_response,i)
	if bValue(suggestAllContent) then
		doAssignmentByRef str,asp_substr(value,0,50)
		doAssignmentByRef pos,my_stripos(str,searchFor,0)
		if IsFalse(pos) then
			ResponseWrite str
		else
			ResponseWrite (((CSmartStr(asp_substr(str,0,pos)) & "<b>") & CSmartStr(asp_substr(str,pos,asp_strlen(searchFor)))) & "</b>") & CSmartStr(asp_substr(str,CSmartDbl(pos)+CSmartDbl(asp_strlen(searchFor)),empty))
		end if
		ResponseWrite vblf
	else
		ResponseWrite ((("<b>" & CSmartStr(asp_substr(value,0,asp_strlen(searchFor)))) & "</b>") & CSmartStr(asp_substr(value,asp_strlen(searchFor),50-CSmartDbl(asp_strlen(searchFor))))) & vblf
	end if
	i = CSmartDbl(i)+1
loop
%>
