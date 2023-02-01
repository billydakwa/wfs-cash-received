<%
%>
<!--#include file="include/dbcommon.asp"-->
<%
doAssignmentByRef table,postvalue("table")
if not bValue(checkTableName(table,false)) then
	Response.End
end if
asp_include ("include/" & CSmartStr(table)) & "_variables.asp",false
doAssignmentByRef id,postvalue("id")
doAssignmentByRef field,postvalue("field")
if not bValue(CheckFieldPermissions(field,"")) then
	Set returnJSON = (CreateDictionary2("success",false,"error","Error: You have not permission for read this text"))
	ResponseWrite my_json_encode(returnJSON)
	response.end
end if
if not bValue(gQuery.HasGroupBy()) then
	gQuery.RemoveAllFieldsExcept_p1 GetFieldIndex(field,"")
end if
doAssignmentByRef keysArr,GetTableData(strTableName,".Keys",CreateDictionary())
Set keys = (CreateDictionary())
GetCollectionBounds keysArr,fulltext_loopIdx2,fulltext_loopMax2
do while fulltext_loopIdx2<=fulltext_loopMax2
	ind = GetCollectionKey(keysArr,fulltext_loopIdx2)
	doAssignment k,ArrayElement(keysArr,ind)
	setArrElement keys,k,postvalue("key" & CSmartStr(CSmartDbl(ind)+1))
	fulltext_loopIdx2=fulltext_loopIdx2+1
loop
doAssignmentByRef where,KeyWhere(keys,"")
doAssignmentByRef secOpt,GetTableData(strTableName,".nSecOptions",CreateDictionary())
doAssignmentByRef sql,gSQLWhere(where,"")
doAssignmentByRef rs,db_query(sql,conn)
if not bValue(rs) or not bValue(doAssignmentByRef(data,db_fetch_array(rs))) then
	Set returnJSON = (CreateDictionary2("success",false,"error","Error: Wrong SQL query"))
	ResponseWrite my_json_encode(returnJSON)
	response.end
end if
doAssignmentByRef value,asp_nl2br(htmlspecialchars(ArrayElement(data,field)))
Set returnJSON = (CreateDictionary2("success",true,"textCont",value))
ResponseWrite my_json_encode(returnJSON)
response.end
%>
