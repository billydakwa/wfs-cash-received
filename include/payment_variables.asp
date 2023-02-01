<%
strTableName = "payment"
setArrElement Session,"OwnerID",Session(("_" & CSmartStr(strTableName)) & "_OwnerID")
strOriginalTableName = "payment"
gstrOrderBy = "order by payment.payment_date"
if bValue(asp_strlen(gstrOrderBy)) and not IsEqual(asp_strtolower(asp_substr(gstrOrderBy,0,8)),"order by") then
	gstrOrderBy = "order by " & CSmartStr(gstrOrderBy)
end if
Set g_orderindexes = (CreateDictionary())
setArrElement g_orderindexes,asp_count(g_orderindexes),CreateDictionary3(Empty,3,Empty,IIF(1,"ASC","DESC"),Empty,"payment.payment_date")
gsqlHead = "SELECT payment.station,  part.name,  payment.payment_date,  payment.fop,  awb.[prefix],  awb.serial,  payment.amount_received"
gsqlFrom = "FROM payment  INNER JOIN part ON payment.part = part.code  INNER JOIN awb ON payment.awb_seq = awb.seq"
gsqlWhereExpr = "(payment.payment_date between '2022-12-31' and GETDATE())"
gsqlTail = ""
asp_include getabspath("include/payment_settings.asp"),false
doAssignmentByRef gQuery,queryData_payment
doAssignmentByRef eventObj,ArrayElement(tableEvents,"payment")
reportCaseSensitiveGroupFields = false
doAssignmentByRef gstrSQL,gSQLWhere("","")
%>
