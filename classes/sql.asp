<%
Function SQLQueryIsNull(ByVal value)
	SQLQueryIsNull = IsIdentical(value,"null")
	Exit Function
End Function

'------ Class SQLEntity ------
Class SQLEntity
	Public Function init_SQLEntity()
		DoAssignmentByRef init_SQLEntity,method_SQLEntity_SQLEntity(me)
	End Function
	Public Function IsAggregationFunctionCall()
		DoAssignmentByRef IsAggregationFunctionCall,method_SQLEntity_IsAggregationFunctionCall(me)
	End Function
	Public Function IsBinary()
		DoAssignmentByRef IsBinary,method_SQLEntity_IsBinary(me)
	End Function
	Public Function IsAsterisk()
		DoAssignmentByRef IsAsterisk,method_SQLEntity_IsAsterisk(me)
	End Function
	Public Function SetQuery_p1(ByRef query)
		DoAssignmentByRef SetQuery_p1,method_SQLEntity_SetQuery(me,query)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
	End Sub
' end serialize
End Class
' SQLEntity implementation 
Function method_SQLEntity_SQLEntity(byref this_object)
End Function
Function method_SQLEntity_IsAggregationFunctionCall(byref this_object)
	method_SQLEntity_IsAggregationFunctionCall = false
	Exit Function
End Function
Function method_SQLEntity_IsBinary(byref this_object)
	method_SQLEntity_IsBinary = false
	Exit Function
End Function
Function method_SQLEntity_IsAsterisk(byref this_object)
	method_SQLEntity_IsAsterisk = false
	Exit Function
End Function
Function method_SQLEntity_SetQuery(byref this_object,ByRef query)
End Function

'------ Class SQLNonParsed extends SQLEntity ------
Class SQLNonParsed
	Public m_sql
	Public Function init_SQLNonParsed_p1(ByVal proto)
		DoAssignmentByRef init_SQLNonParsed_p1,method_SQLNonParsed_SQLNonParsed(me,proto)
	End Function
	Public Function toSql_p1(ByVal query)
		DoAssignmentByRef toSql_p1,method_SQLNonParsed_toSql(me,query)
	End Function
	Public Function IsAsterisk()
		DoAssignmentByRef IsAsterisk,method_SQLNonParsed_IsAsterisk(me)
	End Function
	Public Function IsAggregationFunctionCall()
		DoAssignmentByRef IsAggregationFunctionCall,method_SQLEntity_IsAggregationFunctionCall(me)
	End Function
	Public Function IsBinary()
		DoAssignmentByRef IsBinary,method_SQLEntity_IsBinary(me)
	End Function
	Public Function SetQuery_p1(ByRef query)
		DoAssignmentByRef SetQuery_p1,method_SQLEntity_SetQuery(me,query)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"m_sql", m_sql
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment m_sql, dict("m_sql")
	End Sub
' end serialize
End Class
' SQLNonParsed implementation 
Function method_SQLNonParsed_SQLNonParsed(byref this_object,ByVal proto)
	doClassAssignment this_object,"m_sql",IIF(not IsEmpty(ArrayElement(proto,"m_sql")),ArrayElement(proto,"m_sql"),"")
End Function
Function method_SQLNonParsed_toSql(byref this_object,ByVal query)
	doAssignmentByRef method_SQLNonParsed_toSql,this_object.m_sql
	Exit Function
End Function
Function method_SQLNonParsed_IsAsterisk(byref this_object)
	Dim last
	doAssignmentByRef last,asp_substr(this_object.m_sql,CSmartDbl(asp_strlen(this_object.m_sql))-1,empty)
	method_SQLNonParsed_IsAsterisk = IsEqual(last,"*")
	Exit Function
End Function

'------ Class SQLField extends SQLEntity ------
Class SQLField
	Public m_strName
	Public m_strTable
	Public Function init_SQLField_p1(ByVal proto)
		DoAssignmentByRef init_SQLField_p1,method_SQLField_SQLField(me,proto)
	End Function
	Public Function toSql_p1(ByVal query)
		DoAssignmentByRef toSql_p1,method_SQLField_toSql(me,query)
	End Function
	Public Function GetType()
		DoAssignmentByRef GetType,method_SQLField_GetType(me)
	End Function
	Public Function IsBinary()
		DoAssignmentByRef IsBinary,method_SQLField_IsBinary(me)
	End Function
	Public Function IsAggregationFunctionCall()
		DoAssignmentByRef IsAggregationFunctionCall,method_SQLEntity_IsAggregationFunctionCall(me)
	End Function
	Public Function IsAsterisk()
		DoAssignmentByRef IsAsterisk,method_SQLEntity_IsAsterisk(me)
	End Function
	Public Function SetQuery_p1(ByRef query)
		DoAssignmentByRef SetQuery_p1,method_SQLEntity_SetQuery(me,query)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"m_strName", m_strName
		setArrElement out,"m_strTable", m_strTable
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment m_strName, dict("m_strName")
		DoAssignment m_strTable, dict("m_strTable")
	End Sub
' end serialize
End Class
' SQLField implementation 
Function method_SQLField_SQLField(byref this_object,ByVal proto)
	doClassAssignment this_object,"m_strName",IIF(not IsEmpty(ArrayElement(proto,"m_strName")),ArrayElement(proto,"m_strName"),null)
	doClassAssignment this_object,"m_strTable",IIF(not IsEmpty(ArrayElement(proto,"m_strTable")),ArrayElement(proto,"m_strTable"),null)
End Function
Function method_SQLField_toSql(byref this_object,ByVal query)
	if IsEqual(this_object.m_strTable,"") or IsEqual(query.TableCount(),1) then
		doAssignmentByRef method_SQLField_toSql,AddFieldWrappers(this_object.m_strName)
		Exit Function
	else
		method_SQLField_toSql = (CSmartStr(AddTableWrappers(this_object.m_strTable)) & ".") & CSmartStr(AddFieldWrappers(this_object.m_strName))
		Exit Function
	end if
End Function
Function method_SQLField_GetType(byref this_object)
	doAssignmentByRef method_SQLField_GetType,GetFieldType(this_object.m_strName,this_object.m_strTable)
	Exit Function
End Function
Function method_SQLField_IsBinary(byref this_object)
	Dim this
	doAssignmentByRef method_SQLField_IsBinary,IsBinaryType(this_object.GetType())
	Exit Function
End Function

'------ Class SQLFieldListItem extends SQLEntity ------
Class SQLFieldListItem
	Public m_expr
	Public m_alias
	Public Function init_SQLFieldListItem_p1(ByVal proto)
		DoAssignmentByRef init_SQLFieldListItem_p1,method_SQLFieldListItem_SQLFieldListItem(me,proto)
	End Function
	Public Function toSql_p1(ByVal query)
		DoAssignmentByRef toSql_p1,method_SQLFieldListItem_toSql(me,query)
	End Function
	Public Function IsAsterisk()
		DoAssignmentByRef IsAsterisk,method_SQLFieldListItem_IsAsterisk(me)
	End Function
	Public Function IsAggregationFunctionCall()
		DoAssignmentByRef IsAggregationFunctionCall,method_SQLFieldListItem_IsAggregationFunctionCall(me)
	End Function
	Public Function IsBinary()
		DoAssignmentByRef IsBinary,method_SQLEntity_IsBinary(me)
	End Function
	Public Function SetQuery_p1(ByRef query)
		DoAssignmentByRef SetQuery_p1,method_SQLEntity_SetQuery(me,query)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"m_expr", m_expr
		setArrElement out,"m_alias", m_alias
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment m_expr, dict("m_expr")
		DoAssignment m_alias, dict("m_alias")
	End Sub
' end serialize
End Class
' SQLFieldListItem implementation 
Function method_SQLFieldListItem_SQLFieldListItem(byref this_object,ByVal proto)
	doClassAssignment this_object,"m_expr",IIF(not IsEmpty(ArrayElement(proto,"m_expr")),ArrayElement(proto,"m_expr"),null)
	doClassAssignment this_object,"m_alias",IIF(not IsEmpty(ArrayElement(proto,"m_alias")),ArrayElement(proto,"m_alias"),null)
End Function
Function method_SQLFieldListItem_toSql(byref this_object,ByVal query)
	Dim ret
	ret = ""
	if bValue(this_object.m_expr) then
		if bValue(is_string(this_object.m_expr)) then
			doAssignment ret,this_object.m_expr
		else
			doAssignmentByRef ret,this_object.m_expr.toSql_p1(query)
			if bValue(is_a(this_object.m_expr,"SQLQuery")) then
				ret = ("(" & CSmartStr(ret)) & ")"
			end if
		end if
	end if
	if not IsEqual(this_object.m_alias,"") then
		ret = CSmartStr(ret) & (" as " & CSmartStr(AddFieldWrappers(this_object.m_alias)))
	end if
	doAssignmentByRef method_SQLFieldListItem_toSql,ret
	Exit Function
End Function
Function method_SQLFieldListItem_IsAsterisk(byref this_object)
	if bValue(is_object(this_object.m_expr)) then
		doAssignmentByRef method_SQLFieldListItem_IsAsterisk,this_object.m_expr.IsAsterisk()
		Exit Function
	end if
	method_SQLFieldListItem_IsAsterisk = false
	Exit Function
End Function
Function method_SQLFieldListItem_IsAggregationFunctionCall(byref this_object)
	if bValue(is_object(this_object.m_expr)) then
		doAssignmentByRef method_SQLFieldListItem_IsAggregationFunctionCall,this_object.m_expr.IsAggregationFunctionCall()
		Exit Function
	end if
	method_SQLFieldListItem_IsAggregationFunctionCall = false
	Exit Function
End Function

'------ Class SQLFromListItem extends SQLEntity ------
Class SQLFromListItem
	Public m_link
	Public m_table
	Public m_alias
	Public m_joinon
	Public Function init_SQLFromListItem_p1(ByVal proto)
		DoAssignmentByRef init_SQLFromListItem_p1,method_SQLFromListItem_SQLFromListItem(me,proto)
	End Function
	Public Function SetQuery_p1(ByRef query)
		DoAssignmentByRef SetQuery_p1,method_SQLFromListItem_SetQuery(me,query)
	End Function
	Public Function toSql_p2(ByVal query,ByVal first)
		DoAssignmentByRef toSql_p2,method_SQLFromListItem_toSql(me,query,first)
	End Function
	Public Function IsAggregationFunctionCall()
		DoAssignmentByRef IsAggregationFunctionCall,method_SQLEntity_IsAggregationFunctionCall(me)
	End Function
	Public Function IsBinary()
		DoAssignmentByRef IsBinary,method_SQLEntity_IsBinary(me)
	End Function
	Public Function IsAsterisk()
		DoAssignmentByRef IsAsterisk,method_SQLEntity_IsAsterisk(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"m_link", m_link
		setArrElement out,"m_table", m_table
		setArrElement out,"m_alias", m_alias
		setArrElement out,"m_joinon", m_joinon
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment m_link, dict("m_link")
		DoAssignment m_table, dict("m_table")
		DoAssignment m_alias, dict("m_alias")
		DoAssignment m_joinon, dict("m_joinon")
	End Sub
' end serialize
End Class
' SQLFromListItem implementation 
Function method_SQLFromListItem_SQLFromListItem(byref this_object,ByVal proto)
	doClassAssignment this_object,"m_link",IIF(not IsEmpty(ArrayElement(proto,"m_link")),ArrayElement(proto,"m_link"),null)
	doClassAssignment this_object,"m_table",IIF(not IsEmpty(ArrayElement(proto,"m_table")),ArrayElement(proto,"m_table"),null)
	doClassAssignment this_object,"m_alias",IIF(not IsEmpty(ArrayElement(proto,"m_alias")),ArrayElement(proto,"m_alias"),null)
	doClassAssignment this_object,"m_joinon",IIF(not IsEmpty(ArrayElement(proto,"m_joinon")),ArrayElement(proto,"m_joinon"),null)
End Function
Function method_SQLFromListItem_SetQuery(byref this_object,ByRef query)
	if bValue(is_object(this_object.m_table)) then
		this_object.m_table.SetQuery_p1 query
	end if
End Function
Function method_SQLFromListItem_toSql(byref this_object,ByVal query,ByVal first)
	Dim ret,joinStr
	ret = ""
	if bValue(is_a(this_object.m_table,"SQLTable")) then
		ret = CSmartStr(ret) & CSmartStr(this_object.m_table.toSql_p1(query))
	else
		if bValue(this_object.m_table) then
			ret = CSmartStr(ret) & (("(" & CSmartStr(this_object.m_table.toSql_p1(query))) & ")")
		end if
	end if
	if not IsEqual(this_object.m_alias,"") then
		ret = CSmartStr(ret) & (" as " & CSmartStr(AddFieldWrappers(this_object.m_alias)))
	end if
	if IsEqual(this_object.m_link,"SQLL_MAIN") then
		doAssignmentByRef method_SQLFromListItem_toSql,ret
		Exit Function
	end if
	do
		If IsEqual(this_object.m_link,"SQLL_INNERJOIN") then
			ret = " INNER JOIN " & CSmartStr(ret)
			exit do
		End If
		If IsEqual(this_object.m_link,"SQLL_INNERJOIN") or IsEqual(this_object.m_link,"SQLL_NATURALJOIN") then
			ret = " NATURAL JOIN " & CSmartStr(ret)
			exit do
		End If
		If IsEqual(this_object.m_link,"SQLL_INNERJOIN") or IsEqual(this_object.m_link,"SQLL_NATURALJOIN") or IsEqual(this_object.m_link,"SQLL_LEFTJOIN") then
			ret = " LEFT OUTER JOIN " & CSmartStr(ret)
			exit do
		End If
		If IsEqual(this_object.m_link,"SQLL_INNERJOIN") or IsEqual(this_object.m_link,"SQLL_NATURALJOIN") or IsEqual(this_object.m_link,"SQLL_LEFTJOIN") or IsEqual(this_object.m_link,"SQLL_RIGHTJOIN") then
			ret = " RIGHT OUTER JOIN " & CSmartStr(ret)
			exit do
		End If
		If IsEqual(this_object.m_link,"SQLL_INNERJOIN") or IsEqual(this_object.m_link,"SQLL_NATURALJOIN") or IsEqual(this_object.m_link,"SQLL_LEFTJOIN") or IsEqual(this_object.m_link,"SQLL_RIGHTJOIN") or IsEqual(this_object.m_link,"SQLL_FULLOUTERJOIN") then
			ret = " FULL OUTER JOIN " & CSmartStr(ret)
			exit do
		End If
		If IsEqual(this_object.m_link,"SQLL_INNERJOIN") or IsEqual(this_object.m_link,"SQLL_NATURALJOIN") or IsEqual(this_object.m_link,"SQLL_LEFTJOIN") or IsEqual(this_object.m_link,"SQLL_RIGHTJOIN") or IsEqual(this_object.m_link,"SQLL_FULLOUTERJOIN") or IsEqual(this_object.m_link,"SQLL_CROSSJOIN") then
			ret = CSmartStr(IIF(not bValue(first),",","")) & CSmartStr(ret)
			exit do
		End If
	Loop While false
	doAssignmentByRef joinStr,this_object.m_joinon.toSql_p1(query)
	if not IsEqual(joinStr,"") then
		ret = CSmartStr(ret) & (" ON " & CSmartStr(joinStr))
	end if
	doAssignmentByRef method_SQLFromListItem_toSql,ret
	Exit Function
End Function

'------ Class SQLJoinOn extends SQLEntity ------
Class SQLJoinOn
	Public m_field1
	Public m_field2
	Public Function init_SQLJoinOn_p1(ByVal proto)
		DoAssignmentByRef init_SQLJoinOn_p1,method_SQLJoinOn_SQLJoinOn(me,proto)
	End Function
	Public Function IsAggregationFunctionCall()
		DoAssignmentByRef IsAggregationFunctionCall,method_SQLEntity_IsAggregationFunctionCall(me)
	End Function
	Public Function IsBinary()
		DoAssignmentByRef IsBinary,method_SQLEntity_IsBinary(me)
	End Function
	Public Function IsAsterisk()
		DoAssignmentByRef IsAsterisk,method_SQLEntity_IsAsterisk(me)
	End Function
	Public Function SetQuery_p1(ByRef query)
		DoAssignmentByRef SetQuery_p1,method_SQLEntity_SetQuery(me,query)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"m_field1", m_field1
		setArrElement out,"m_field2", m_field2
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment m_field1, dict("m_field1")
		DoAssignment m_field2, dict("m_field2")
	End Sub
' end serialize
End Class
' SQLJoinOn implementation 
Function method_SQLJoinOn_SQLJoinOn(byref this_object,ByVal proto)
	doClassAssignment this_object,"m_field1",IIF(not IsEmpty(ArrayElement(proto,"m_field1")),ArrayElement(proto,"m_field1"),null)
	doClassAssignment this_object,"m_field2",IIF(not IsEmpty(ArrayElement(proto,"m_field2")),ArrayElement(proto,"m_field2"),null)
End Function

'------ Class SQLFunctionCall extends SQLEntity ------
Class SQLFunctionCall
	Public m_functiontype
	Public m_strFunctionName
	Public m_arguments
	Public Function init_SQLFunctionCall_p1(ByVal proto)
		DoAssignmentByRef init_SQLFunctionCall_p1,method_SQLFunctionCall_SQLFunctionCall(me,proto)
	End Function
	Public Function toSql_p1(ByVal query)
		DoAssignmentByRef toSql_p1,method_SQLFunctionCall_toSql(me,query)
	End Function
	Public Function IsAggregationFunctionCall()
		DoAssignmentByRef IsAggregationFunctionCall,method_SQLFunctionCall_IsAggregationFunctionCall(me)
	End Function
	Public Function IsBinary()
		DoAssignmentByRef IsBinary,method_SQLEntity_IsBinary(me)
	End Function
	Public Function IsAsterisk()
		DoAssignmentByRef IsAsterisk,method_SQLEntity_IsAsterisk(me)
	End Function
	Public Function SetQuery_p1(ByRef query)
		DoAssignmentByRef SetQuery_p1,method_SQLEntity_SetQuery(me,query)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"m_functiontype", m_functiontype
		setArrElement out,"m_strFunctionName", m_strFunctionName
		setArrElement out,"m_arguments", m_arguments
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment m_functiontype, dict("m_functiontype")
		DoAssignment m_strFunctionName, dict("m_strFunctionName")
		DoAssignment m_arguments, dict("m_arguments")
	End Sub
' end serialize
End Class
' SQLFunctionCall implementation 
Function method_SQLFunctionCall_SQLFunctionCall(byref this_object,ByVal proto)
	doClassAssignment this_object,"m_functiontype",IIF(not IsEmpty(ArrayElement(proto,"m_functiontype")),ArrayElement(proto,"m_functiontype"),null)
	doClassAssignment this_object,"m_strFunctionName",IIF(not IsEmpty(ArrayElement(proto,"m_strFunctionName")),ArrayElement(proto,"m_strFunctionName"),null)
	doClassAssignment this_object,"m_arguments",IIF(not IsEmpty(ArrayElement(proto,"m_arguments")),ArrayElement(proto,"m_arguments"),null)
End Function
Function method_SQLFunctionCall_toSql(byref this_object,ByVal query)
	Dim ret,args,a
	ret = ""
	do
		If IsEqual(this_object.m_functiontype,"SQLF_AVG") then
			ret = CSmartStr(ret) & " AVG"
			exit do
		End If
		If IsEqual(this_object.m_functiontype,"SQLF_AVG") or IsEqual(this_object.m_functiontype,"SQLF_SUM") then
			ret = CSmartStr(ret) & " SUM"
			exit do
		End If
		If IsEqual(this_object.m_functiontype,"SQLF_AVG") or IsEqual(this_object.m_functiontype,"SQLF_SUM") or IsEqual(this_object.m_functiontype,"SQLF_MIN") then
			ret = CSmartStr(ret) & " MIN"
			exit do
		End If
		If IsEqual(this_object.m_functiontype,"SQLF_AVG") or IsEqual(this_object.m_functiontype,"SQLF_SUM") or IsEqual(this_object.m_functiontype,"SQLF_MIN") or IsEqual(this_object.m_functiontype,"SQLF_MAX") then
			ret = CSmartStr(ret) & " MAX"
			exit do
		End If
		If IsEqual(this_object.m_functiontype,"SQLF_AVG") or IsEqual(this_object.m_functiontype,"SQLF_SUM") or IsEqual(this_object.m_functiontype,"SQLF_MIN") or IsEqual(this_object.m_functiontype,"SQLF_MAX") or IsEqual(this_object.m_functiontype,"SQLF_COUNT") then
			ret = CSmartStr(ret) & " COUNT"
			exit do
		End If
		ret = CSmartStr(ret) & CSmartStr(this_object.m_strFunctionName)
	Loop While false
	Set args = (CreateDictionary())
	GetCollectionBounds this_object.m_arguments,c_sql_loopIdx2,c_sql_loopMax2
	do while c_sql_loopIdx2<=c_sql_loopMax2
		c_sql_arrKey2 = GetCollectionKey(this_object.m_arguments,c_sql_loopIdx2)
		doAssignment a,ArrayElement(this_object.m_arguments,c_sql_arrKey2)
		setArrElement args,asp_count(args),a.toSql_p1(query)
		c_sql_loopIdx2=c_sql_loopIdx2+1
	loop
	ret = CSmartStr(ret) & (("(" & CSmartStr(asp_implode(", ",args))) & ")")
	doAssignmentByRef method_SQLFunctionCall_toSql,ret
	Exit Function
End Function
Function method_SQLFunctionCall_IsAggregationFunctionCall(byref this_object)
	do
		If IsEqual(this_object.m_functiontype,"SQLF_AVG") or IsEqual(this_object.m_functiontype,"SQLF_SUM") or IsEqual(this_object.m_functiontype,"SQLF_MIN") or IsEqual(this_object.m_functiontype,"SQLF_MAX") or IsEqual(this_object.m_functiontype,"SQLF_COUNT") then
			method_SQLFunctionCall_IsAggregationFunctionCall = true
			Exit Function
		End If
	Loop While false
	method_SQLFunctionCall_IsAggregationFunctionCall = false
	Exit Function
End Function

'------ Class SQLGroupByItem extends SQLEntity ------
Class SQLGroupByItem
	Public m_column
	Public Function init_SQLGroupByItem_p1(ByVal proto)
		DoAssignmentByRef init_SQLGroupByItem_p1,method_SQLGroupByItem_SQLGroupByItem(me,proto)
	End Function
	Public Function toSql_p1(ByVal query)
		DoAssignmentByRef toSql_p1,method_SQLGroupByItem_toSql(me,query)
	End Function
	Public Function IsAggregationFunctionCall()
		DoAssignmentByRef IsAggregationFunctionCall,method_SQLEntity_IsAggregationFunctionCall(me)
	End Function
	Public Function IsBinary()
		DoAssignmentByRef IsBinary,method_SQLEntity_IsBinary(me)
	End Function
	Public Function IsAsterisk()
		DoAssignmentByRef IsAsterisk,method_SQLEntity_IsAsterisk(me)
	End Function
	Public Function SetQuery_p1(ByRef query)
		DoAssignmentByRef SetQuery_p1,method_SQLEntity_SetQuery(me,query)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"m_column", m_column
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment m_column, dict("m_column")
	End Sub
' end serialize
End Class
' SQLGroupByItem implementation 
Function method_SQLGroupByItem_SQLGroupByItem(byref this_object,ByVal proto)
	doClassAssignment this_object,"m_column",IIF(not IsEmpty(ArrayElement(proto,"m_column")),ArrayElement(proto,"m_column"),null)
End Function
Function method_SQLGroupByItem_toSql(byref this_object,ByVal query)
	doAssignmentByRef method_SQLGroupByItem_toSql,this_object.m_column.toSql_p1(query)
	Exit Function
End Function
Const LE_AND = 1
Const LE_OR = 2
Const LE_ISNULL = 1
Const LE_EQ = 2
Const F_AGG = 1
Const F_SIMPLE = 2

'------ Class SQLLogicalExpr extends SQLEntity ------
Class SQLLogicalExpr
	Public m_uniontype
	Public m_column
	Public m_strCase
	Public m_havingmode
	Public m_contained
	Public m_inBrackets
	Public m_useAlias
	Public m_sql
	Public query
	Public postSql
	Public Function init_SQLLogicalExpr_p1(ByVal proto)
		DoAssignmentByRef init_SQLLogicalExpr_p1,method_SQLLogicalExpr_SQLLogicalExpr(me,proto)
	End Function
	Public Function SetQuery_p1(ByRef query)
		DoAssignmentByRef SetQuery_p1,method_SQLLogicalExpr_SetQuery(me,query)
	End Function
	Public Function toSql_p1(ByVal query)
		DoAssignmentByRef toSql_p1,method_SQLLogicalExpr_toSql(me,query)
	End Function
	Public Function haveContained()
		DoAssignmentByRef haveContained,method_SQLLogicalExpr_haveContained(me)
	End Function
	Public Function InjectSelf()
		DoAssignmentByRef InjectSelf,method_SQLLogicalExpr_InjectSelf(me)
	End Function
	Public Function AddFilter_p3(ByVal field,ByVal func,ByVal value)
		DoAssignmentByRef AddFilter_p3,method_SQLLogicalExpr_AddFilter(me,field,func,value)
	End Function
	Public Function AddFilter_p2(ByVal field,ByVal func)
		DoAssignmentByRef AddFilter_p2,method_SQLLogicalExpr_AddFilter(me,field,func,"")
	End Function
	Public Function AddSqlFilter_p1(ByVal value)
		DoAssignmentByRef AddSqlFilter_p1,method_SQLLogicalExpr_AddSqlFilter(me,value)
	End Function
	Public Function AddFilterByKeys_p1(ByRef keys)
		DoAssignmentByRef AddFilterByKeys_p1,method_SQLLogicalExpr_AddFilterByKeys(me,keys)
	End Function
	Public Function IsAggregationFunctionCall()
		DoAssignmentByRef IsAggregationFunctionCall,method_SQLEntity_IsAggregationFunctionCall(me)
	End Function
	Public Function IsBinary()
		DoAssignmentByRef IsBinary,method_SQLEntity_IsBinary(me)
	End Function
	Public Function IsAsterisk()
		DoAssignmentByRef IsAsterisk,method_SQLEntity_IsAsterisk(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"m_uniontype", m_uniontype
		setArrElement out,"m_column", m_column
		setArrElement out,"m_strCase", m_strCase
		setArrElement out,"m_havingmode", m_havingmode
		setArrElement out,"m_contained", m_contained
		setArrElement out,"m_inBrackets", m_inBrackets
		setArrElement out,"m_useAlias", m_useAlias
		setArrElement out,"m_sql", m_sql
		setArrElement out,"query", query
		setArrElement out,"postSql", postSql
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment m_uniontype, dict("m_uniontype")
		DoAssignment m_column, dict("m_column")
		DoAssignment m_strCase, dict("m_strCase")
		DoAssignment m_havingmode, dict("m_havingmode")
		DoAssignment m_contained, dict("m_contained")
		DoAssignment m_inBrackets, dict("m_inBrackets")
		DoAssignment m_useAlias, dict("m_useAlias")
		DoAssignment m_sql, dict("m_sql")
		DoAssignment query, dict("query")
		DoAssignment postSql, dict("postSql")
	End Sub
' end serialize
End Class
' SQLLogicalExpr implementation 
Function method_SQLLogicalExpr_SQLLogicalExpr(byref this_object,ByVal proto)
	doClassAssignment this_object,"m_sql",IIF(not IsEmpty(ArrayElement(proto,"m_sql")),ArrayElement(proto,"m_sql"),null)
	doClassAssignment this_object,"m_uniontype",IIF(not IsEmpty(ArrayElement(proto,"m_uniontype")),ArrayElement(proto,"m_uniontype"),null)
	doClassAssignment this_object,"m_column",IIF(not IsEmpty(ArrayElement(proto,"m_column")),ArrayElement(proto,"m_column"),null)
	doClassAssignment this_object,"m_strCase",IIF(not IsEmpty(ArrayElement(proto,"m_strCase")),ArrayElement(proto,"m_strCase"),null)
	doClassAssignment this_object,"m_havingmode",IIF(not IsEmpty(ArrayElement(proto,"m_havingmode")),ArrayElement(proto,"m_havingmode"),null)
	doClassAssignment this_object,"m_contained",IIF(not IsEmpty(ArrayElement(proto,"m_contained")),ArrayElement(proto,"m_contained"),null)
	doClassAssignment this_object,"m_inBrackets",IIF(not IsEmpty(ArrayElement(proto,"m_inBrackets")),ArrayElement(proto,"m_inBrackets"),null)
	doClassAssignment this_object,"m_useAlias",IIF(not IsEmpty(ArrayElement(proto,"m_useAlias")),ArrayElement(proto,"m_useAlias"),null)
	doClassAssignment this_object,"postSql",CreateDictionary()
End Function
Function method_SQLLogicalExpr_SetQuery(byref this_object,ByRef query)
	Dim nCnt
	doClassAssignmentByRef this_object,"query",query
	nCnt = 0
	do while IsLess(nCnt,asp_count(this_object.m_contained))
		ArrayElement(this_object.m_contained,nCnt).SetQuery_p1 query
		nCnt = CSmartDbl(nCnt)+1
	loop
End Function
Function method_SQLLogicalExpr_toSql(byref this_object,ByVal query)
	Dim ret,this,glue,contained,cSql,c,nCnt
	ret = ""
	if bValue(this_object.haveContained()) then
		glue = ""
		if IsEqual(this_object.m_uniontype,"SQLL_AND") then
			glue = " AND "
		else
			if IsEqual(this_object.m_uniontype,"SQLL_OR") then
				glue = " OR "
			else
				Response.End
			end if
		end if
		Set contained = (CreateDictionary())
		GetCollectionBounds this_object.m_contained,c_sql_loopIdx4,c_sql_loopMax4
		do while c_sql_loopIdx4<=c_sql_loopMax4
			c_sql_arrKey4 = GetCollectionKey(this_object.m_contained,c_sql_loopIdx4)
			doAssignment c,ArrayElement(this_object.m_contained,c_sql_arrKey4)
			doAssignmentByRef cSql,c.toSql_p1(query)
			if not IsEqual(cSql,"") then
				setArrElement contained,asp_count(contained),cSql
			end if
			c_sql_loopIdx4=c_sql_loopIdx4+1
		loop
		if IsLess(0,asp_count(contained)) then
			doAssignmentByRef ret,asp_implode(glue,contained)
		end if
		if IsLess(0,asp_count(this_object.postSql)) then
			if IsEqual(ret,"") then
				ret = CSmartStr(ret) & (("(" & CSmartStr(ArrayElement(this_object.postSql,0))) & ")")
			else
				ret = ((((("(" & CSmartStr(ret)) & ")") & CSmartStr(glue)) & "(") & CSmartStr(ArrayElement(this_object.postSql,0))) & ")"
			end if
			nCnt = 1
			do while IsLess(nCnt,asp_count(this_object.postSql))
				ret = CSmartStr(ret) & (((CSmartStr(glue) & "(") & CSmartStr(ArrayElement(this_object.postSql,nCnt))) & ")")
				nCnt = CSmartDbl(nCnt)+1
			loop
		end if
		if bValue(this_object.m_inBrackets) then
			ret = ("(" & CSmartStr(ret)) & ")"
		end if
		doAssignmentByRef method_SQLLogicalExpr_toSql,ret
		Exit Function
	end if
	if not bValue(this_object.m_column) then
		doAssignment ret,this_object.m_sql
	else
		if bValue(this_object.m_useAlias) then
			doAssignment ret,this_object.m_column.m_alias
		else
			doAssignmentByRef ret,this_object.m_column.toSql_p1(query)
		end if
	end if
	if IsEqual(this_object.m_strCase,"NOT") then
		method_SQLLogicalExpr_toSql = " NOT " & CSmartStr(ret)
		Exit Function
	else
		if not IsEqual(ret,"") then
			ret = CSmartStr(ret) & (" " & CSmartStr(this_object.m_strCase))
		end if
	end if
	if bValue(this_object.m_inBrackets) then
		ret = ("(" & CSmartStr(ret)) & ")"
	end if
	doAssignmentByRef method_SQLLogicalExpr_toSql,ret
	Exit Function
End Function
Function method_SQLLogicalExpr_haveContained(byref this_object)
	method_SQLLogicalExpr_haveContained = IsLess(0,asp_count(this_object.m_contained)) or IsLess(0,asp_count(this_object.postSql))
	Exit Function
End Function
Function method_SQLLogicalExpr_InjectSelf(byref this_object)
	Dim insertLE
	Set insertLE = (CreateClass("SQLLogicalExpr",1,CreateDictionary6("m_uniontype",this_object.m_uniontype,"m_strCase",this_object.m_strCase,"m_havingmode",this_object.m_havingmode,"m_column",this_object.m_column,"m_contained",this_object.m_contained,"m_inBrackets",this_object.m_inBrackets),Empty,Empty,Empty,Empty,Empty,Empty))
	doClassAssignment insertLE,"postSql",this_object.postSql
	doClassAssignment insertLE,"query",this_object.query
	this_object.m_uniontype = "SQLL_AND"
	this_object.m_strCase = ""
	this_object.m_havingmode = false
	this_object.m_column = null
	this_object.m_inBrackets = false
	doClassAssignment this_object,"postSql",CreateDictionary()
	doClassAssignment this_object,"m_contained",CreateDictionary1(Empty,insertLE)
End Function
Function method_SQLLogicalExpr_AddFilter(byref this_object,ByVal field,ByVal func,ByVal value)
	Dim oDefaultTable,var_case,this
	if bValue(is_string(field)) then
		doAssignmentByRef oDefaultTable,this_object.query.DefaultTable()
		doAssignmentByRef field,oDefaultTable.FieldByName_p1(field)
	end if
	var_case = ""
	do
		If IsEqual(func,LE_ISNULL) then
			var_case = " is null "
			exit do
		End If
		If IsEqual(func,LE_ISNULL) or IsEqual(func,LE_EQ) then
			var_case = " = " & CSmartStr(value)
			exit do
		End If
		Response.End
	Loop While false
	if not IsEqual(this_object.m_uniontype,"SQLL_AND") then
		this_object.InjectSelf 
	end if
	setArrElement this_object.m_contained,asp_count(this_object.m_contained),CreateClass("SQLLogicalExpr",1,CreateDictionary4("m_uniontype","SQLL_UNKNOWN","m_strCase",var_case,"m_havingmode",false,"m_column",field),Empty,Empty,Empty,Empty,Empty,Empty)
End Function
Function method_SQLLogicalExpr_AddSqlFilter(byref this_object,ByVal value)
	Dim this
	if bValue(is_string(value)) and not IsEqual(value,"") then
		if not IsEqual(this_object.m_uniontype,"SQLL_AND") then
			this_object.InjectSelf 
		end if
		setArrElement this_object.postSql,asp_count(this_object.postSql),value
	end if
End Function
Function method_SQLLogicalExpr_AddFilterByKeys(byref this_object,ByRef keys)
	Dim oDefaultTable,tableKeys,value,tk,this
	doAssignmentByRef oDefaultTable,this_object.query.DefaultTable()
	doAssignmentByRef tableKeys,oDefaultTable.GetKeyFields()
	GetCollectionBounds tableKeys,c_sql_loopIdx6,c_sql_loopMax6
	do while c_sql_loopIdx6<=c_sql_loopMax6
		c_sql_arrKey6 = GetCollectionKey(tableKeys,c_sql_loopIdx6)
		doAssignment tk,ArrayElement(tableKeys,c_sql_arrKey6)
		doAssignmentByRef value,make_db_value(tk,ArrayElement(keys,tk),"","","")
		if bValue(SQLQueryIsNull(value)) then
			this_object.AddFilter_p2 tk,LE_ISNULL
		else
			this_object.AddFilter_p3 tk,LE_EQ,value
		end if
		c_sql_loopIdx6=c_sql_loopIdx6+1
	loop
End Function

'------ Class SQLOrderByItem extends SQLEntity ------
Class SQLOrderByItem
	Public m_column
	Public m_bAsc
	Public m_nColumn
	Public Function init_SQLOrderByItem_p1(ByVal proto)
		DoAssignmentByRef init_SQLOrderByItem_p1,method_SQLOrderByItem_SQLOrderByItem(me,proto)
	End Function
	Public Function toSql_p1(ByVal query)
		DoAssignmentByRef toSql_p1,method_SQLOrderByItem_toSql(me,query)
	End Function
	Public Function FixIfGreater_p1(ByVal num)
		DoAssignmentByRef FixIfGreater_p1,method_SQLOrderByItem_FixIfGreater(me,num)
	End Function
	Public Function IsAggregationFunctionCall()
		DoAssignmentByRef IsAggregationFunctionCall,method_SQLEntity_IsAggregationFunctionCall(me)
	End Function
	Public Function IsBinary()
		DoAssignmentByRef IsBinary,method_SQLEntity_IsBinary(me)
	End Function
	Public Function IsAsterisk()
		DoAssignmentByRef IsAsterisk,method_SQLEntity_IsAsterisk(me)
	End Function
	Public Function SetQuery_p1(ByRef query)
		DoAssignmentByRef SetQuery_p1,method_SQLEntity_SetQuery(me,query)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"m_column", m_column
		setArrElement out,"m_bAsc", m_bAsc
		setArrElement out,"m_nColumn", m_nColumn
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment m_column, dict("m_column")
		DoAssignment m_bAsc, dict("m_bAsc")
		DoAssignment m_nColumn, dict("m_nColumn")
	End Sub
' end serialize
End Class
' SQLOrderByItem implementation 
Function method_SQLOrderByItem_SQLOrderByItem(byref this_object,ByVal proto)
	doClassAssignment this_object,"m_column",IIF(not IsEmpty(ArrayElement(proto,"m_column")),ArrayElement(proto,"m_column"),null)
	doClassAssignment this_object,"m_bAsc",IIF(not IsEmpty(ArrayElement(proto,"m_bAsc")),ArrayElement(proto,"m_bAsc"),null)
	doClassAssignment this_object,"m_nColumn",IIF(not IsEmpty(ArrayElement(proto,"m_nColumn")),ArrayElement(proto,"m_nColumn"),null)
End Function
Function method_SQLOrderByItem_toSql(byref this_object,ByVal query)
	Dim ret
	ret = ""
	if IsEqual(0,this_object.m_nColumn) then
		ret = CSmartStr(ret) & CSmartStr(this_object.m_column.toSql_p1(query))
	else
		ret = CSmartStr(ret) & CSmartStr(this_object.m_nColumn)
	end if
	if not bValue(this_object.m_bAsc) then
		ret = CSmartStr(ret) & " DESC "
	end if
	doAssignmentByRef method_SQLOrderByItem_toSql,ret
	Exit Function
End Function
Function method_SQLOrderByItem_FixIfGreater(byref this_object,ByVal num)
	if not IsEqual(0,this_object.m_nColumn) and IsLessOrEqual(num,this_object.m_nColumn) then
		this_object.m_nColumn = CSmartDbl(this_object.m_nColumn)-1
	end if
End Function

'------ Class SQLTable extends SQLEntity ------
Class SQLTable
	Public m_strName
	Public m_columns
	Public query
	Public Function init_SQLTable_p1(ByVal proto)
		DoAssignmentByRef init_SQLTable_p1,method_SQLTable_SQLTable(me,proto)
	End Function
	Public Function SetQuery_p1(ByRef query)
		DoAssignmentByRef SetQuery_p1,method_SQLTable_SetQuery(me,query)
	End Function
	Public Function toSql_p1(ByVal query)
		DoAssignmentByRef toSql_p1,method_SQLTable_toSql(me,query)
	End Function
	Public Function GetKeyFields()
		DoAssignmentByRef GetKeyFields,method_SQLTable_GetKeyFields(me)
	End Function
	Public Function FieldByName_p1(ByVal name)
		DoAssignmentByRef FieldByName_p1,method_SQLTable_FieldByName(me,name)
	End Function
	Public Function IsAggregationFunctionCall()
		DoAssignmentByRef IsAggregationFunctionCall,method_SQLEntity_IsAggregationFunctionCall(me)
	End Function
	Public Function IsBinary()
		DoAssignmentByRef IsBinary,method_SQLEntity_IsBinary(me)
	End Function
	Public Function IsAsterisk()
		DoAssignmentByRef IsAsterisk,method_SQLEntity_IsAsterisk(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"m_strName", m_strName
		setArrElement out,"m_columns", m_columns
		setArrElement out,"query", query
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment m_strName, dict("m_strName")
		DoAssignment m_columns, dict("m_columns")
		DoAssignment query, dict("query")
	End Sub
' end serialize
End Class
' SQLTable implementation 
Function method_SQLTable_SQLTable(byref this_object,ByVal proto)
	doClassAssignment this_object,"m_strName",IIF(not IsEmpty(ArrayElement(proto,"m_strName")),ArrayElement(proto,"m_strName"),null)
	doClassAssignment this_object,"m_columns",IIF(not IsEmpty(ArrayElement(proto,"m_columns")),ArrayElement(proto,"m_columns"),null)
End Function
Function method_SQLTable_SetQuery(byref this_object,ByRef query)
	doClassAssignment this_object,"query",query
End Function
Function method_SQLTable_toSql(byref this_object,ByVal query)
	doAssignmentByRef method_SQLTable_toSql,AddTableWrappers(this_object.m_strName)
	Exit Function
End Function
Function method_SQLTable_GetKeyFields(byref this_object)
	doAssignmentByRef method_SQLTable_GetKeyFields,dal.GetDBTableKeys(this_object.m_strName)
	Exit Function
End Function
Function method_SQLTable_FieldByName(byref this_object,ByVal name)
	doAssignmentByRef method_SQLTable_FieldByName,this_object.query.FieldByNameAndTable_p2(this_object.m_strName,name)
	Exit Function
End Function

'------ Class SQLQuery extends SQLEntity ------
Class SQLQuery
	Public m_strHead
	Public m_strFieldList
	Public m_strFrom
	Public m_strWhere
	Public m_strOrderBy
	Public m_strTail
	Public m_where
	Public m_having
	Public m_fieldlist
	Public m_fromlist
	Public m_groupby
	Public m_orderby
	Public bHasAsterisks
	Public Function init_SQLQuery_p1(ByVal proto)
		DoAssignmentByRef init_SQLQuery_p1,method_SQLQuery_SQLQuery(me,proto)
	End Function
	Public Function HeadToSql()
		DoAssignmentByRef HeadToSql,method_SQLQuery_HeadToSql(me)
	End Function
	Public Function GroupByToSql()
		DoAssignmentByRef GroupByToSql,method_SQLQuery_GroupByToSql(me)
	End Function
	Public Function toSql_p3(ByVal strwhere,ByVal strorderby,ByVal strhaving)
		DoAssignmentByRef toSql_p3,method_SQLQuery_toSql(me,strwhere,strorderby,strhaving)
	End Function
	Public Function toSql_p2(ByVal strwhere,ByVal strorderby)
		DoAssignmentByRef toSql_p2,method_SQLQuery_toSql(me,strwhere,strorderby,null)
	End Function
	Public Function toSql_p1(ByVal strwhere)
		DoAssignmentByRef toSql_p1,method_SQLQuery_toSql(me,strwhere,null,null)
	End Function
	Public Function toSql()
		DoAssignmentByRef toSql,method_SQLQuery_toSql(me,null,null,null)
	End Function
	Public Function TableCount()
		DoAssignmentByRef TableCount,method_SQLQuery_TableCount(me)
	End Function
	Public Function TableByField_p1(ByVal field)
		DoAssignmentByRef TableByField_p1,method_SQLQuery_TableByField(me,field)
	End Function
	Public Function TableByName_p1(ByVal tableName)
		DoAssignmentByRef TableByName_p1,method_SQLQuery_TableByName(me,tableName)
	End Function
	Public Function FieldByNameAndTable_p2(ByVal table,ByVal field)
		DoAssignmentByRef FieldByNameAndTable_p2,method_SQLQuery_FieldByNameAndTable(me,table,field)
	End Function
	Public Function IsAggrFuncField_p1(ByVal idx)
		DoAssignmentByRef IsAggrFuncField_p1,method_SQLQuery_IsAggrFuncField(me,idx)
	End Function
	Public Function DefaultTable()
		DoAssignmentByRef DefaultTable,method_SQLQuery_DefaultTable(me)
	End Function
	Public Function dropOrderByField_p1(ByVal field)
		DoAssignmentByRef dropOrderByField_p1,method_SQLQuery_dropOrderByField(me,field)
	End Function
	Public Function ReplaceFieldsWithDummies_p1(ByVal fieldindices)
		DoAssignmentByRef ReplaceFieldsWithDummies_p1,method_SQLQuery_ReplaceFieldsWithDummies(me,fieldindices)
	End Function
	Public Function RemoveAllFieldsExcept_p1(ByVal idx)
		DoAssignmentByRef RemoveAllFieldsExcept_p1,method_SQLQuery_RemoveAllFieldsExcept(me,idx)
	End Function
	Public Function Where()
		DoAssignmentByRef Where,method_SQLQuery_Where(me)
	End Function
	Public Function Having()
		DoAssignmentByRef Having,method_SQLQuery_Having(me)
	End Function
	Public Function Copy()
		DoAssignmentByRef Copy,method_SQLQuery_Copy(me)
	End Function
	Public Function HasGroupBy()
		DoAssignmentByRef HasGroupBy,method_SQLQuery_HasGroupBy(me)
	End Function
	Public Function HasAsterisks()
		DoAssignmentByRef HasAsterisks,method_SQLQuery_HasAsterisks(me)
	End Function
	Public Function IsAggregationFunctionCall()
		DoAssignmentByRef IsAggregationFunctionCall,method_SQLEntity_IsAggregationFunctionCall(me)
	End Function
	Public Function IsBinary()
		DoAssignmentByRef IsBinary,method_SQLEntity_IsBinary(me)
	End Function
	Public Function IsAsterisk()
		DoAssignmentByRef IsAsterisk,method_SQLEntity_IsAsterisk(me)
	End Function
	Public Function SetQuery_p1(ByRef query)
		DoAssignmentByRef SetQuery_p1,method_SQLEntity_SetQuery(me,query)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"m_strHead", m_strHead
		setArrElement out,"m_strFieldList", m_strFieldList
		setArrElement out,"m_strFrom", m_strFrom
		setArrElement out,"m_strWhere", m_strWhere
		setArrElement out,"m_strOrderBy", m_strOrderBy
		setArrElement out,"m_strTail", m_strTail
		setArrElement out,"m_where", m_where
		setArrElement out,"m_having", m_having
		setArrElement out,"m_fieldlist", m_fieldlist
		setArrElement out,"m_fromlist", m_fromlist
		setArrElement out,"m_groupby", m_groupby
		setArrElement out,"m_orderby", m_orderby
		setArrElement out,"bHasAsterisks", bHasAsterisks
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment m_strHead, dict("m_strHead")
		DoAssignment m_strFieldList, dict("m_strFieldList")
		DoAssignment m_strFrom, dict("m_strFrom")
		DoAssignment m_strWhere, dict("m_strWhere")
		DoAssignment m_strOrderBy, dict("m_strOrderBy")
		DoAssignment m_strTail, dict("m_strTail")
		DoAssignment m_where, dict("m_where")
		DoAssignment m_having, dict("m_having")
		DoAssignment m_fieldlist, dict("m_fieldlist")
		DoAssignment m_fromlist, dict("m_fromlist")
		DoAssignment m_groupby, dict("m_groupby")
		DoAssignment m_orderby, dict("m_orderby")
		DoAssignment bHasAsterisks, dict("bHasAsterisks")
	End Sub
' end serialize
End Class
' SQLQuery implementation 
Function method_SQLQuery_SQLQuery(byref this_object,ByVal proto)
	this_object.bHasAsterisks = false
	Dim nCnt,i
	doClassAssignment this_object,"m_strHead",IIF(not IsEmpty(ArrayElement(proto,"m_strHead")),ArrayElement(proto,"m_strHead"),null)
	doClassAssignment this_object,"m_strFieldList",IIF(not IsEmpty(ArrayElement(proto,"m_strFieldList")),ArrayElement(proto,"m_strFieldList"),null)
	doClassAssignment this_object,"m_strFrom",IIF(not IsEmpty(ArrayElement(proto,"m_strFrom")),ArrayElement(proto,"m_strFrom"),null)
	doClassAssignment this_object,"m_strWhere",IIF(not IsEmpty(ArrayElement(proto,"m_strWhere")),ArrayElement(proto,"m_strWhere"),null)
	doClassAssignment this_object,"m_strOrderBy",IIF(not IsEmpty(ArrayElement(proto,"m_strOrderBy")),ArrayElement(proto,"m_strOrderBy"),null)
	doClassAssignment this_object,"m_strTail",IIF(not IsEmpty(ArrayElement(proto,"m_strTail")),ArrayElement(proto,"m_strTail"),null)
	doClassAssignment this_object,"m_where",IIF(not IsEmpty(ArrayElement(proto,"m_where")),ArrayElement(proto,"m_where"),null)
	doClassAssignment this_object,"m_having",IIF(not IsEmpty(ArrayElement(proto,"m_having")),ArrayElement(proto,"m_having"),null)
	doClassAssignment this_object,"m_fieldlist",IIF(not IsEmpty(ArrayElement(proto,"m_fieldlist")),ArrayElement(proto,"m_fieldlist"),null)
	doClassAssignment this_object,"m_fromlist",IIF(not IsEmpty(ArrayElement(proto,"m_fromlist")),ArrayElement(proto,"m_fromlist"),null)
	doClassAssignment this_object,"m_groupby",IIF(not IsEmpty(ArrayElement(proto,"m_groupby")),ArrayElement(proto,"m_groupby"),null)
	doClassAssignment this_object,"m_orderby",IIF(not IsEmpty(ArrayElement(proto,"m_orderby")),ArrayElement(proto,"m_orderby"),null)
	nCnt = 0
	do while IsLess(nCnt,asp_count(this_object.m_fromlist))
		ArrayElement(this_object.m_fromlist,nCnt).SetQuery_p1 this_object
		nCnt = CSmartDbl(nCnt)+1
	loop
	this_object.m_where.SetQuery_p1 this_object
	if not bValue(asp_is_array(this_object.m_fieldlist)) then
		Exit Function
	end if
	i = 0
	do while IsLess(i,asp_count(this_object.m_fieldlist))
		if bValue(ArrayElement(this_object.m_fieldlist,i).IsAsterisk()) then
			this_object.bHasAsterisks = true
			exit do
		end if
		i = CSmartDbl(i)+1
	loop
End Function
Function method_SQLQuery_HeadToSql(byref this_object)
	Dim sql,fields,f
	sql = ""
	sql = CSmartStr(sql) & CSmartStr(this_object.m_strHead)
	if not IsEqual(sql,"") then
		sql = CSmartStr(sql) & vbcrlf
	end if
	Set fields = (CreateDictionary())
	GetCollectionBounds this_object.m_fieldlist,c_sql_loopIdx9,c_sql_loopMax9
	do while c_sql_loopIdx9<=c_sql_loopMax9
		c_sql_arrKey9 = GetCollectionKey(this_object.m_fieldlist,c_sql_loopIdx9)
		doAssignment f,ArrayElement(this_object.m_fieldlist,c_sql_arrKey9)
		setArrElement fields,asp_count(fields),f.toSql_p1(this_object)
		c_sql_loopIdx9=c_sql_loopIdx9+1
	loop
	if IsLess(0,asp_count(fields)) then
		sql = CSmartStr(sql) & CSmartStr(asp_implode(", ",fields))
	end if
	doAssignmentByRef method_SQLQuery_HeadToSql,sql
	Exit Function
End Function
Function method_SQLQuery_GroupByToSql(byref this_object)
	Dim sql,groupby,g
	sql = ""
	Set groupby = (CreateDictionary())
	GetCollectionBounds this_object.m_groupby,c_sql_loopIdx10,c_sql_loopMax10
	do while c_sql_loopIdx10<=c_sql_loopMax10
		c_sql_arrKey10 = GetCollectionKey(this_object.m_groupby,c_sql_loopIdx10)
		doAssignment g,ArrayElement(this_object.m_groupby,c_sql_arrKey10)
		setArrElement groupby,asp_count(groupby),g.toSql_p1(this_object)
		c_sql_loopIdx10=c_sql_loopIdx10+1
	loop
	if IsLess(0,asp_count(groupby)) then
		sql = CSmartStr(sql) & " GROUP BY "
		sql = CSmartStr(sql) & CSmartStr(asp_implode(",",groupby))
		sql = CSmartStr(sql) & " "
	end if
	doAssignmentByRef method_SQLQuery_GroupByToSql,sql
	Exit Function
End Function
Function method_SQLQuery_toSql(byref this_object,ByVal strwhere,ByVal strorderby,ByVal strhaving)
	Dim sql,this,sqlFromList,nCnt,fromlist,where,having,orderby,g
	if bValue(is_a(strwhere,"SQLQuery")) then
		strwhere = null
	end if
	doAssignmentByRef sql,this_object.HeadToSql()
	if IsLess(0,asp_count(this_object.m_fromlist)) then
		sql = CSmartStr(sql) & vbcrlf
		sql = CSmartStr(sql) & " FROM "
	end if
	Set fromlist = (CreateDictionary())
	nCnt = 0
	do while IsLess(nCnt,asp_count(this_object.m_fromlist))
		setArrElement fromlist,asp_count(fromlist),ArrayElement(this_object.m_fromlist,nCnt).toSql_p2(this_object,IsEqual(nCnt,0))
		nCnt = CSmartDbl(nCnt)+1
	loop
	sql = CSmartStr(sql) & CSmartStr(asp_implode(vbcrlf,fromlist))
	if not IsNull(strwhere) then
		if not IsIdentical(strwhere,"") then
			sql = CSmartStr(sql) & ((" WHERE " & CSmartStr(strwhere)) & vbcrlf)
		end if
	else
		doAssignmentByRef where,this_object.m_where.toSql_p1(this_object)
		if not IsEqual(where,"") then
			sql = CSmartStr(sql) & ((" WHERE " & CSmartStr(where)) & vbcrlf)
		end if
	end if
	sql = CSmartStr(sql) & CSmartStr(this_object.GroupByToSql())
	if not IsNull(strhaving) then
		if not IsIdentical(strhaving,"") then
			sql = CSmartStr(sql) & ((" HAVING " & CSmartStr(strhaving)) & vbcrlf)
		end if
	else
		doAssignmentByRef having,this_object.m_having.toSql_p1(this_object)
		if not IsEqual(having,"") then
			sql = CSmartStr(sql) & ((" HAVING " & CSmartStr(having)) & vbcrlf)
		end if
	end if
	if not IsNull(strorderby) then
		sql = CSmartStr(sql) & (CSmartStr(strorderby) & vbcrlf)
	else
		Set orderby = (CreateDictionary())
		GetCollectionBounds this_object.m_orderby,c_sql_loopIdx13,c_sql_loopMax13
		do while c_sql_loopIdx13<=c_sql_loopMax13
			c_sql_arrKey13 = GetCollectionKey(this_object.m_orderby,c_sql_loopIdx13)
			doAssignment g,ArrayElement(this_object.m_orderby,c_sql_arrKey13)
			setArrElement orderby,asp_count(orderby),g.toSql_p1(this_object)
			c_sql_loopIdx13=c_sql_loopIdx13+1
		loop
		if IsLess(0,asp_count(orderby)) then
			sql = CSmartStr(sql) & " ORDER BY "
			sql = CSmartStr(sql) & CSmartStr(asp_implode(",",orderby))
			sql = CSmartStr(sql) & vbcrlf
		end if
	end if
	doAssignmentByRef method_SQLQuery_toSql,sql
	Exit Function
End Function
Function method_SQLQuery_TableCount(byref this_object)
	doAssignmentByRef method_SQLQuery_TableCount,asp_count(this_object.m_fromlist)
	Exit Function
End Function
Function method_SQLQuery_TableByField(byref this_object,ByVal field)
	Dim f
	GetCollectionBounds this_object.m_fromlist,c_sql_loopIdx14,c_sql_loopMax14
	do while c_sql_loopIdx14<=c_sql_loopMax14
		c_sql_arrKey14 = GetCollectionKey(this_object.m_fromlist,c_sql_loopIdx14)
		doAssignment f,ArrayElement(this_object.m_fromlist,c_sql_arrKey14)
		if IsEqual(f.m_table.m_strName,field.m_strTable) then
			doAssignmentByRef method_SQLQuery_TableByField,f
			Exit Function
		end if
		c_sql_loopIdx14=c_sql_loopIdx14+1
	loop
	method_SQLQuery_TableByField = null
	Exit Function
End Function
Function method_SQLQuery_TableByName(byref this_object,ByVal tableName)
	Dim f
	GetCollectionBounds this_object.m_fromlist,c_sql_loopIdx15,c_sql_loopMax15
	do while c_sql_loopIdx15<=c_sql_loopMax15
		c_sql_arrKey15 = GetCollectionKey(this_object.m_fromlist,c_sql_loopIdx15)
		doAssignment f,ArrayElement(this_object.m_fromlist,c_sql_arrKey15)
		if IsEqual(f.m_table.m_strName,tableName) then
			doAssignmentByRef method_SQLQuery_TableByName,f
			Exit Function
		end if
		c_sql_loopIdx15=c_sql_loopIdx15+1
	loop
	method_SQLQuery_TableByName = null
	Exit Function
End Function
Function method_SQLQuery_FieldByNameAndTable(byref this_object,ByVal table,ByVal field)
	Dim f
	GetCollectionBounds this_object.m_fieldlist,c_sql_loopIdx16,c_sql_loopMax16
	do while c_sql_loopIdx16<=c_sql_loopMax16
		c_sql_arrKey16 = GetCollectionKey(this_object.m_fieldlist,c_sql_loopIdx16)
		doAssignment f,ArrayElement(this_object.m_fieldlist,c_sql_arrKey16)
		if IsEqual(f.m_expr.m_strName,field) and IsEqual(f.m_expr.m_strTable,table) then
			doAssignmentByRef method_SQLQuery_FieldByNameAndTable,f
			Exit Function
		end if
		c_sql_loopIdx16=c_sql_loopIdx16+1
	loop
	method_SQLQuery_FieldByNameAndTable = null
	Exit Function
End Function
Function method_SQLQuery_IsAggrFuncField(byref this_object,ByVal idx)
	Dim this
	if bValue(this_object.HasAsterisks()) then
		method_SQLQuery_IsAggrFuncField = false
		Exit Function
	end if
	if not (not IsEmpty(ArrayElement(this_object.m_fieldlist,idx))) then
		method_SQLQuery_IsAggrFuncField = false
		Exit Function
	end if
	doAssignmentByRef method_SQLQuery_IsAggrFuncField,ArrayElement(this_object.m_fieldlist,idx).IsAggregationFunctionCall()
	Exit Function
End Function
Function method_SQLQuery_DefaultTable(byref this_object)
	doAssignmentByRef method_SQLQuery_DefaultTable,ArrayElement(this_object.m_fromlist,0).m_table
	Exit Function
End Function
Function method_SQLQuery_dropOrderByField(byref this_object,ByVal field)
	Dim nOb,obField,nOb2
	nOb = 0
	do while IsLess(nOb,asp_count(this_object.m_orderby))
		doAssignment obField,ArrayElement(this_object.m_orderby,nOb).m_column
		if IsEqual(obField.m_strName,field.m_strName) and IsEqual(obField.m_strTable,field.m_strTable) then
			nOb2 = 0
			do while IsLess(nOb2,asp_count(this_object.m_orderby))
				ArrayElement(this_object.m_orderby,nOb2).FixIfGreater_p1 nOb2
				nOb2 = CSmartDbl(nOb2)+1
			loop
			asp_array_splice this_object.m_orderby,nOb,1
			exit do
		end if
		nOb = CSmartDbl(nOb)+1
	loop
End Function
Function method_SQLQuery_ReplaceFieldsWithDummies(byref this_object,ByVal fieldindices)
	Dim this,idx
	if bValue(this_object.HasAsterisks()) then
		Exit Function
	end if
	GetCollectionBounds fieldindices,c_sql_loopIdx19,c_sql_loopMax19
	do while c_sql_loopIdx19<=c_sql_loopMax19
		c_sql_arrKey19 = GetCollectionKey(fieldindices,c_sql_loopIdx19)
		doAssignment idx,ArrayElement(fieldindices,c_sql_arrKey19)
		if not (not IsEmpty(ArrayElement(this_object.m_fieldlist,CSmartDbl(idx)-1))) then
			Exit Function
		end if
		setArrElement this_object.m_fieldlist,CSmartDbl(idx)-1,CreateClass("SQLFieldListItem",1,CreateDictionary2("m_alias","runnerdummy" & CSmartStr(idx),"m_expr","1"),Empty,Empty,Empty,Empty,Empty,Empty)
		c_sql_loopIdx19=c_sql_loopIdx19+1
	loop
End Function
Function method_SQLQuery_RemoveAllFieldsExcept(byref this_object,ByVal idx)
	Dim this,removeindices,i
	if bValue(this_object.HasAsterisks()) then
		Exit Function
	end if
	Set removeindices = (CreateDictionary())
	i = 0
	do while IsLess(i,asp_count(this_object.m_fieldlist))
		if not IsEqual(i,CSmartDbl(idx)-1) then
			setArrElement removeindices,asp_count(removeindices),CSmartDbl(i)+1
		end if
		i = CSmartDbl(i)+1
	loop
	this_object.ReplaceFieldsWithDummies_p1 removeindices
End Function
Function method_SQLQuery_Where(byref this_object)
	doAssignmentByRef method_SQLQuery_Where,this_object.m_where
	Exit Function
End Function
Function method_SQLQuery_Having(byref this_object)
	doAssignmentByRef method_SQLQuery_Having,this_object.m_having
	Exit Function
End Function
Function method_SQLQuery_Copy(byref this_object)
	Dim this
	doAssignmentByRef method_SQLQuery_Copy,unserialize(serialize(this_object))
	Exit Function
End Function
Function method_SQLQuery_HasGroupBy(byref this_object)
	method_SQLQuery_HasGroupBy = IsLess(0,asp_count(this_object.m_groupby))
	Exit Function
End Function
Function method_SQLQuery_HasAsterisks(byref this_object)
	doAssignmentByRef method_SQLQuery_HasAsterisks,this_object.bHasAsterisks
	Exit Function
End Function

'------ Class SQLCountQuery extends SQLQuery ------
Class SQLCountQuery
	Public m_strHead
	Public m_strFieldList
	Public m_strFrom
	Public m_strWhere
	Public m_strOrderBy
	Public m_strTail
	Public m_where
	Public m_having
	Public m_fieldlist
	Public m_fromlist
	Public m_groupby
	Public m_orderby
	Public bHasAsterisks
	Public Function init_SQLCountQuery_p1(ByVal src)
		DoAssignmentByRef init_SQLCountQuery_p1,method_SQLCountQuery_SQLCountQuery(me,src)
	End Function
	Public Function toSql_p1(ByVal where)
		DoAssignmentByRef toSql_p1,method_SQLCountQuery_toSql(me,where)
	End Function
	Public Function HeadToSql()
		DoAssignmentByRef HeadToSql,method_SQLQuery_HeadToSql(me)
	End Function
	Public Function GroupByToSql()
		DoAssignmentByRef GroupByToSql,method_SQLQuery_GroupByToSql(me)
	End Function
	Public Function TableCount()
		DoAssignmentByRef TableCount,method_SQLQuery_TableCount(me)
	End Function
	Public Function TableByField_p1(ByVal field)
		DoAssignmentByRef TableByField_p1,method_SQLQuery_TableByField(me,field)
	End Function
	Public Function TableByName_p1(ByVal tableName)
		DoAssignmentByRef TableByName_p1,method_SQLQuery_TableByName(me,tableName)
	End Function
	Public Function FieldByNameAndTable_p2(ByVal table,ByVal field)
		DoAssignmentByRef FieldByNameAndTable_p2,method_SQLQuery_FieldByNameAndTable(me,table,field)
	End Function
	Public Function IsAggrFuncField_p1(ByVal idx)
		DoAssignmentByRef IsAggrFuncField_p1,method_SQLQuery_IsAggrFuncField(me,idx)
	End Function
	Public Function DefaultTable()
		DoAssignmentByRef DefaultTable,method_SQLQuery_DefaultTable(me)
	End Function
	Public Function dropOrderByField_p1(ByVal field)
		DoAssignmentByRef dropOrderByField_p1,method_SQLQuery_dropOrderByField(me,field)
	End Function
	Public Function ReplaceFieldsWithDummies_p1(ByVal fieldindices)
		DoAssignmentByRef ReplaceFieldsWithDummies_p1,method_SQLQuery_ReplaceFieldsWithDummies(me,fieldindices)
	End Function
	Public Function RemoveAllFieldsExcept_p1(ByVal idx)
		DoAssignmentByRef RemoveAllFieldsExcept_p1,method_SQLQuery_RemoveAllFieldsExcept(me,idx)
	End Function
	Public Function Where()
		DoAssignmentByRef Where,method_SQLQuery_Where(me)
	End Function
	Public Function Having()
		DoAssignmentByRef Having,method_SQLQuery_Having(me)
	End Function
	Public Function Copy()
		DoAssignmentByRef Copy,method_SQLQuery_Copy(me)
	End Function
	Public Function HasGroupBy()
		DoAssignmentByRef HasGroupBy,method_SQLQuery_HasGroupBy(me)
	End Function
	Public Function HasAsterisks()
		DoAssignmentByRef HasAsterisks,method_SQLQuery_HasAsterisks(me)
	End Function
	Public Function IsAggregationFunctionCall()
		DoAssignmentByRef IsAggregationFunctionCall,method_SQLEntity_IsAggregationFunctionCall(me)
	End Function
	Public Function IsBinary()
		DoAssignmentByRef IsBinary,method_SQLEntity_IsBinary(me)
	End Function
	Public Function IsAsterisk()
		DoAssignmentByRef IsAsterisk,method_SQLEntity_IsAsterisk(me)
	End Function
	Public Function SetQuery_p1(ByRef query)
		DoAssignmentByRef SetQuery_p1,method_SQLEntity_SetQuery(me,query)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"m_strHead", m_strHead
		setArrElement out,"m_strFieldList", m_strFieldList
		setArrElement out,"m_strFrom", m_strFrom
		setArrElement out,"m_strWhere", m_strWhere
		setArrElement out,"m_strOrderBy", m_strOrderBy
		setArrElement out,"m_strTail", m_strTail
		setArrElement out,"m_where", m_where
		setArrElement out,"m_having", m_having
		setArrElement out,"m_fieldlist", m_fieldlist
		setArrElement out,"m_fromlist", m_fromlist
		setArrElement out,"m_groupby", m_groupby
		setArrElement out,"m_orderby", m_orderby
		setArrElement out,"bHasAsterisks", bHasAsterisks
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment m_strHead, dict("m_strHead")
		DoAssignment m_strFieldList, dict("m_strFieldList")
		DoAssignment m_strFrom, dict("m_strFrom")
		DoAssignment m_strWhere, dict("m_strWhere")
		DoAssignment m_strOrderBy, dict("m_strOrderBy")
		DoAssignment m_strTail, dict("m_strTail")
		DoAssignment m_where, dict("m_where")
		DoAssignment m_having, dict("m_having")
		DoAssignment m_fieldlist, dict("m_fieldlist")
		DoAssignment m_fromlist, dict("m_fromlist")
		DoAssignment m_groupby, dict("m_groupby")
		DoAssignment m_orderby, dict("m_orderby")
		DoAssignment bHasAsterisks, dict("bHasAsterisks")
	End Sub
' end serialize
End Class
' SQLCountQuery implementation 
Function method_SQLCountQuery_SQLCountQuery(byref this_object,ByVal src)
	this_object.bHasAsterisks = false
	Dim this,k,v
	GetCollectionBounds src,c_sql_loopIdx21,c_sql_loopMax21
	do while c_sql_loopIdx21<=c_sql_loopMax21
		k = GetCollectionKey(src,c_sql_loopIdx21)
		doAssignment v,ArrayElement(src,k)
		setArrElement this_object,k,v
		c_sql_loopIdx21=c_sql_loopIdx21+1
	loop
	this_object.m_strHead = ""
	if not bValue(this_object.HasGroupBy()) then
		doClassAssignment this_object,"m_fieldlist",CreateDictionary()
	end if
End Function
Function method_SQLCountQuery_toSql(byref this_object,ByVal where)
	Dim sql,this,ret
	doAssignmentByRef sql,method_SQLQuery_toSql(this_object,where,null,null)
	if bValue(this_object.HasGroupBy()) then
		ret = ("select count(*) from (" & CSmartStr(sql)) & ") a"
	else
		ret = "select count(*) from " & CSmartStr(sql)
	end if
	doAssignmentByRef method_SQLCountQuery_toSql,ret
	Exit Function
End Function
%>
