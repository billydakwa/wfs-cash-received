<%

Set report_classfields = (CreateDictionary())
Function create_reportfield(ByVal name,ByVal var_type,ByVal interval,ByVal alias)
	if bValue(var_type) then
		if not IsEmpty(ArrayElement(report_classfields,var_type)) then
			doAssignmentByRef create_reportfield,CreateClass(ArrayElement(report_classfields,var_type),3,name,interval,alias,Empty,Empty,Empty,Empty)
			Exit Function
		else
			Response.End
		end if
	end if
End Function

'------ Class ReportField ------
Class ReportField
	Public var_interval
	Public var_name
	Public var_alias
	Public var_sqlname
	Public var_start
	Public var_caseSensitive
	Public var_recordBasedRequest
	Public var_rowsInSummary
	Public var_rowsInHeader
	Public var_viewFormat
	Public var_orderBy
	Public var_oldAlgorithm
	Public Function init_ReportField_p3(ByVal name,ByVal interval,ByVal alias)
		DoAssignmentByRef init_ReportField_p3,method_ReportField_ReportField(me,name,interval,alias)
	End Function
	Public Function getStringSql()
		DoAssignmentByRef getStringSql,method_ReportField_getStringSql(me)
	End Function
	Public Function getFieldName_p2(ByVal fieldValue,ByVal data)
		DoAssignmentByRef getFieldName_p2,method_ReportField_getFieldName(me,fieldValue,data)
	End Function
	Public Function getFieldName_p1(ByVal fieldValue)
		DoAssignmentByRef getFieldName_p1,method_ReportField_getFieldName(me,fieldValue,null)
	End Function
	Public Function getSelectSql_p1(ByVal hasGrouping)
		DoAssignmentByRef getSelectSql_p1,method_ReportField_getSelectSql(me,hasGrouping)
	End Function
	Public Function getSelectSql()
		DoAssignmentByRef getSelectSql,method_ReportField_getSelectSql(me,false)
	End Function
	Public Function getGroupSql()
		DoAssignmentByRef getGroupSql,method_ReportField_getGroupSql(me)
	End Function
	Public Function getOrderSql()
		DoAssignmentByRef getOrderSql,method_ReportField_getOrderSql(me)
	End Function
	Public Function getWhereSql_p1(ByVal groups)
		DoAssignmentByRef getWhereSql_p1,method_ReportField_getWhereSql(me,groups)
	End Function
	Public Function getGroup_p1(ByVal data)
		DoAssignmentByRef getGroup_p1,method_ReportField_getGroup(me,data)
	End Function
	Public Function getKey_p1(ByVal data)
		DoAssignmentByRef getKey_p1,method_ReportField_getKey(me,data)
	End Function
	Public Function setStart_p1(ByVal start)
		DoAssignmentByRef setStart_p1,method_ReportField_setStart(me,start)
	End Function
	Public Function name()
		DoAssignmentByRef name,method_ReportField_name(me)
	End Function
	Public Function alias()
		DoAssignmentByRef alias,method_ReportField_alias(me)
	End Function
	Public Function overrideFormat()
		DoAssignmentByRef overrideFormat,method_ReportField_overrideFormat(me)
	End Function
	Public Function setCaseSensitive_p1(ByVal cs)
		DoAssignmentByRef setCaseSensitive_p1,method_ReportField_setCaseSensitive(me,cs)
	End Function
	Public Function cutNull_p2(ByRef range,ByVal checkEmty)
		DoAssignmentByRef cutNull_p2,method_ReportField_cutNull(me,range,checkEmty)
	End Function
	Public Function cutNull_p1(ByRef range)
		DoAssignmentByRef cutNull_p1,method_ReportField_cutNull(me,range,false)
	End Function
	Public Function getLtGt_p2(ByRef lt,ByRef gt)
		DoAssignmentByRef getLtGt_p2,method_ReportField_getLtGt(me,lt,gt)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"var_interval", var_interval
		setArrElement out,"var_name", var_name
		setArrElement out,"var_alias", var_alias
		setArrElement out,"var_sqlname", var_sqlname
		setArrElement out,"var_start", var_start
		setArrElement out,"var_caseSensitive", var_caseSensitive
		setArrElement out,"var_recordBasedRequest", var_recordBasedRequest
		setArrElement out,"var_rowsInSummary", var_rowsInSummary
		setArrElement out,"var_rowsInHeader", var_rowsInHeader
		setArrElement out,"var_viewFormat", var_viewFormat
		setArrElement out,"var_orderBy", var_orderBy
		setArrElement out,"var_oldAlgorithm", var_oldAlgorithm
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment var_interval, dict("var_interval")
		DoAssignment var_name, dict("var_name")
		DoAssignment var_alias, dict("var_alias")
		DoAssignment var_sqlname, dict("var_sqlname")
		DoAssignment var_start, dict("var_start")
		DoAssignment var_caseSensitive, dict("var_caseSensitive")
		DoAssignment var_recordBasedRequest, dict("var_recordBasedRequest")
		DoAssignment var_rowsInSummary, dict("var_rowsInSummary")
		DoAssignment var_rowsInHeader, dict("var_rowsInHeader")
		DoAssignment var_viewFormat, dict("var_viewFormat")
		DoAssignment var_orderBy, dict("var_orderBy")
		DoAssignment var_oldAlgorithm, dict("var_oldAlgorithm")
	End Sub
' end serialize
End Class
' ReportField implementation 
Function method_ReportField_ReportField(byref this_object,ByVal name,ByVal interval,ByVal alias)
	this_object.var_interval = 0
	this_object.var_name = ""
	this_object.var_alias = ""
	this_object.var_sqlname = ""
	this_object.var_start = 0
	this_object.var_caseSensitive = false
	this_object.var_recordBasedRequest = false
	this_object.var_rowsInSummary = 0
	this_object.var_rowsInHeader = 0
	this_object.var_viewFormat = ""
	this_object.var_orderBy = "ASC"
	this_object.var_oldAlgorithm = false
	doClassAssignment this_object,"var_name",name
	doClassAssignment this_object,"var_interval",interval
	doClassAssignment this_object,"var_alias",alias
	doClassAssignment this_object,"var_sqlname",alias
End Function
Function method_ReportField_getStringSql(byref this_object)
	Response.End
End Function
Function method_ReportField_getFieldName(byref this_object,ByVal fieldValue,ByVal data)
	Response.End
End Function
Function method_ReportField_getSelectSql(byref this_object,ByVal hasGrouping)
	Dim this
	method_ReportField_getSelectSql = CSmartStr(this_object.getStringSql()) & CSmartStr(IIF(this_object.alias()," as " & CSmartStr(cached_ffn(this_object.alias())),""))
	Exit Function
End Function
Function method_ReportField_getGroupSql(byref this_object)
	Dim this
	doAssignmentByRef method_ReportField_getGroupSql,this_object.getStringSql()
	Exit Function
End Function
Function method_ReportField_getOrderSql(byref this_object)
	Dim this
	method_ReportField_getOrderSql = ((CSmartStr(this_object.getStringSql()) & " ") & CSmartStr(this_object.var_orderBy)) & " "
	Exit Function
End Function
Function method_ReportField_getWhereSql(byref this_object,ByVal groups)
	Response.End
End Function
Function method_ReportField_getGroup(byref this_object,ByVal data)
	Dim this
	doAssignmentByRef method_ReportField_getGroup,ArrayElement(data,this_object.alias())
	Exit Function
End Function
Function method_ReportField_getKey(byref this_object,ByVal data)
	Dim this
	doAssignmentByRef method_ReportField_getKey,ArrayElement(data,this_object.alias())
	Exit Function
End Function
Function method_ReportField_setStart(byref this_object,ByVal start)
	Dim this
	doClassAssignment this_object,"var_start",start
	doClassAssignment this_object,"var_sqlname",this_object.alias()
	method_ReportField_setStart = CSmartDbl(start)+1
	Exit Function
End Function
Function method_ReportField_name(byref this_object)
	doAssignmentByRef method_ReportField_name,this_object.var_name
	Exit Function
End Function
Function method_ReportField_alias(byref this_object)
	method_ReportField_alias = CSmartStr(this_object.var_alias) & CSmartStr(this_object.var_start)
	Exit Function
End Function
Function method_ReportField_overrideFormat(byref this_object)
	method_ReportField_overrideFormat = false
	Exit Function
End Function
Function method_ReportField_setCaseSensitive(byref this_object,ByVal cs)
	doClassAssignment this_object,"var_caseSensitive",cs
End Function
Function method_ReportField_cutNull(byref this_object,ByRef range,ByVal checkEmty)
	Dim ret,out,nCnt,b
	ret = false
	Set out = (CreateDictionary())
	nCnt = 0
	do while IsLess(nCnt,asp_count(range))
		b = false
		if IsNull(ArrayElement(range,nCnt)) then
			b = true
			ret = true
		else
			if bValue(checkEmty) and (not bValue(ArrayElement(range,nCnt)) or IsEqual(strcasecmp(ArrayElement(range,nCnt),"null"),0)) then
				b = true
				ret = true
			end if
		end if
		if not bValue(b) then
			setArrElement out,asp_count(out),ArrayElement(range,nCnt)
		end if
		nCnt = CSmartDbl(nCnt)+1
	loop
	doAssignment range,out
	doAssignmentByRef method_ReportField_cutNull,ret
	Exit Function
End Function
Function method_ReportField_getLtGt(byref this_object,ByRef lt,ByRef gt)
	if not IsEqual(this_object.var_orderBy,"ASC") then
		lt = " >= "
		gt = " <= "
	else
		lt = " <= "
		gt = " >= "
	end if
End Function
setArrElement report_classfields,"numeric","ReportNumericField"

'------ Class ReportNumericField extends ReportField ------
Class ReportNumericField
	Public var_interval
	Public var_name
	Public var_alias
	Public var_sqlname
	Public var_start
	Public var_caseSensitive
	Public var_recordBasedRequest
	Public var_rowsInSummary
	Public var_rowsInHeader
	Public var_viewFormat
	Public var_orderBy
	Public var_oldAlgorithm
	Public Function init_ReportNumericField_p3(ByVal name,ByVal interval,ByVal alias)
		DoAssignmentByRef init_ReportNumericField_p3,method_ReportNumericField_ReportNumericField(me,name,interval,alias)
	End Function
	Public Function getStringSql()
		DoAssignmentByRef getStringSql,method_ReportNumericField_getStringSql(me)
	End Function
	Public Function getFieldName_p2(ByVal fieldValue,ByVal data)
		DoAssignmentByRef getFieldName_p2,method_ReportNumericField_getFieldName(me,fieldValue,data)
	End Function
	Public Function getKey_p1(ByVal data)
		DoAssignmentByRef getKey_p1,method_ReportNumericField_getKey(me,data)
	End Function
	Public Function getWhereSql_p1(ByVal groups)
		DoAssignmentByRef getWhereSql_p1,method_ReportNumericField_getWhereSql(me,groups)
	End Function
	Public Function getSelectSql_p1(ByVal hasGrouping)
		DoAssignmentByRef getSelectSql_p1,method_ReportField_getSelectSql(me,hasGrouping)
	End Function
	Public Function getSelectSql()
		DoAssignmentByRef getSelectSql,method_ReportField_getSelectSql(me,false)
	End Function
	Public Function getGroupSql()
		DoAssignmentByRef getGroupSql,method_ReportField_getGroupSql(me)
	End Function
	Public Function getOrderSql()
		DoAssignmentByRef getOrderSql,method_ReportField_getOrderSql(me)
	End Function
	Public Function getGroup_p1(ByVal data)
		DoAssignmentByRef getGroup_p1,method_ReportField_getGroup(me,data)
	End Function
	Public Function setStart_p1(ByVal start)
		DoAssignmentByRef setStart_p1,method_ReportField_setStart(me,start)
	End Function
	Public Function name()
		DoAssignmentByRef name,method_ReportField_name(me)
	End Function
	Public Function alias()
		DoAssignmentByRef alias,method_ReportField_alias(me)
	End Function
	Public Function overrideFormat()
		DoAssignmentByRef overrideFormat,method_ReportField_overrideFormat(me)
	End Function
	Public Function setCaseSensitive_p1(ByVal cs)
		DoAssignmentByRef setCaseSensitive_p1,method_ReportField_setCaseSensitive(me,cs)
	End Function
	Public Function cutNull_p2(ByRef range,ByVal checkEmty)
		DoAssignmentByRef cutNull_p2,method_ReportField_cutNull(me,range,checkEmty)
	End Function
	Public Function cutNull_p1(ByRef range)
		DoAssignmentByRef cutNull_p1,method_ReportField_cutNull(me,range,false)
	End Function
	Public Function getLtGt_p2(ByRef lt,ByRef gt)
		DoAssignmentByRef getLtGt_p2,method_ReportField_getLtGt(me,lt,gt)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"var_interval", var_interval
		setArrElement out,"var_name", var_name
		setArrElement out,"var_alias", var_alias
		setArrElement out,"var_sqlname", var_sqlname
		setArrElement out,"var_start", var_start
		setArrElement out,"var_caseSensitive", var_caseSensitive
		setArrElement out,"var_recordBasedRequest", var_recordBasedRequest
		setArrElement out,"var_rowsInSummary", var_rowsInSummary
		setArrElement out,"var_rowsInHeader", var_rowsInHeader
		setArrElement out,"var_viewFormat", var_viewFormat
		setArrElement out,"var_orderBy", var_orderBy
		setArrElement out,"var_oldAlgorithm", var_oldAlgorithm
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment var_interval, dict("var_interval")
		DoAssignment var_name, dict("var_name")
		DoAssignment var_alias, dict("var_alias")
		DoAssignment var_sqlname, dict("var_sqlname")
		DoAssignment var_start, dict("var_start")
		DoAssignment var_caseSensitive, dict("var_caseSensitive")
		DoAssignment var_recordBasedRequest, dict("var_recordBasedRequest")
		DoAssignment var_rowsInSummary, dict("var_rowsInSummary")
		DoAssignment var_rowsInHeader, dict("var_rowsInHeader")
		DoAssignment var_viewFormat, dict("var_viewFormat")
		DoAssignment var_orderBy, dict("var_orderBy")
		DoAssignment var_oldAlgorithm, dict("var_oldAlgorithm")
	End Sub
' end serialize
End Class
' ReportNumericField implementation 
Function method_ReportNumericField_ReportNumericField(byref this_object,ByVal name,ByVal interval,ByVal alias)
	this_object.var_interval = 0
	this_object.var_name = ""
	this_object.var_alias = ""
	this_object.var_sqlname = ""
	this_object.var_start = 0
	this_object.var_caseSensitive = false
	this_object.var_recordBasedRequest = false
	this_object.var_rowsInSummary = 0
	this_object.var_rowsInHeader = 0
	this_object.var_viewFormat = ""
	this_object.var_orderBy = "ASC"
	this_object.var_oldAlgorithm = false
	method_ReportField_ReportField this_object,name,interval,alias
End Function
Function method_ReportNumericField_getStringSql(byref this_object)
	Dim fname
	doAssignment fname,IIF(this_object.var_oldAlgorithm,GetFullFieldName(this_object.var_name,""),cached_ffn(this_object.var_name))
	if IsLess(0,this_object.var_interval) then
		method_ReportNumericField_getStringSql = (((("floor(" & CSmartStr(fname)) & "/") & CSmartStr(this_object.var_interval)) & ")*") & CSmartStr(this_object.var_interval)
		Exit Function
	else
		doAssignmentByRef method_ReportNumericField_getStringSql,fname
		Exit Function
	end if
End Function
Function method_ReportNumericField_getFieldName(byref this_object,ByVal fieldValue,ByVal data)
	Dim value
	doAssignment value,ArrayElement(data,IIF(this_object.var_recordBasedRequest,this_object.var_name,this_object.var_sqlname))
	if IsNull(value) then
		method_ReportNumericField_getFieldName = "NULL"
		Exit Function
	end if
	if IsLess(0,this_object.var_interval) then
		method_ReportNumericField_getFieldName = (CSmartStr(asp_intval(value)) & " - ") & CSmartStr(CSmartDbl(asp_intval(value))+CSmartDbl(this_object.var_interval))
		Exit Function
	else
		doAssignmentByRef method_ReportNumericField_getFieldName,asp_intval(value)
		Exit Function
	end if
End Function
Function method_ReportNumericField_getKey(byref this_object,ByVal data)
	if bValue(this_object.var_recordBasedRequest) then
		if IsLess(0,this_object.var_interval) then
			method_ReportNumericField_getKey = CSmartDbl(asp_intval(CSmartDbl(ArrayElement(data,this_object.var_name))/CSmartDbl(this_object.var_interval)))*CSmartDbl(this_object.var_interval)
			Exit Function
		else
			doAssignmentByRef method_ReportNumericField_getKey,ArrayElement(data,this_object.var_name)
			Exit Function
		end if
	else
		doAssignmentByRef method_ReportNumericField_getKey,method_ReportField_getKey(this_object,data)
		Exit Function
	end if
End Function
Function method_ReportNumericField_getWhereSql(byref this_object,ByVal groups)
	Dim ret,ssql,this,hasNull,lt,gt
	ret = ""
	doAssignmentByRef ssql,this_object.getStringSql()
	doAssignmentByRef hasNull,this_object.cutNull_p1(groups)
	if IsLess(0,asp_count(groups)) then
		lt = ""
		gt = ""
		this_object.getLtGt_p2 lt,gt
		ret = ((((((("(" & CSmartStr(ssql)) & CSmartStr(gt)) & CSmartStr(ArrayElement(groups,0))) & " AND ") & CSmartStr(ssql)) & CSmartStr(lt)) & CSmartStr(ArrayElement(groups,CSmartDbl(asp_count(groups))-1))) & ")"
	end if
	if bValue(hasNull) then
		ret = CSmartStr(ret) & ((CSmartStr(IIF(ret," OR ","")) & CSmartStr(ssql)) & " IS NULL")
	end if
	doAssignmentByRef method_ReportNumericField_getWhereSql,IIF(ret,("(" & CSmartStr(ret)) & ")","")
	Exit Function
End Function
setArrElement report_classfields,"char","ReportCharField"

'------ Class ReportCharField extends ReportField ------
Class ReportCharField
	Public var_interval
	Public var_name
	Public var_alias
	Public var_sqlname
	Public var_start
	Public var_caseSensitive
	Public var_recordBasedRequest
	Public var_rowsInSummary
	Public var_rowsInHeader
	Public var_viewFormat
	Public var_orderBy
	Public var_oldAlgorithm
	Public Function init_ReportCharField_p3(ByVal name,ByVal interval,ByVal alias)
		DoAssignmentByRef init_ReportCharField_p3,method_ReportCharField_ReportCharField(me,name,interval,alias)
	End Function
	Public Function getStringSql()
		DoAssignmentByRef getStringSql,method_ReportCharField_getStringSql(me)
	End Function
	Public Function getFieldName_p2(ByVal fieldValue,ByVal data)
		DoAssignmentByRef getFieldName_p2,method_ReportCharField_getFieldName(me,fieldValue,data)
	End Function
	Public Function getKey_p1(ByVal data)
		DoAssignmentByRef getKey_p1,method_ReportCharField_getKey(me,data)
	End Function
	Public Function getWhereSql_p1(ByVal groups)
		DoAssignmentByRef getWhereSql_p1,method_ReportCharField_getWhereSql(me,groups)
	End Function
	Public Function getSelectSql_p1(ByVal hasGrouping)
		DoAssignmentByRef getSelectSql_p1,method_ReportField_getSelectSql(me,hasGrouping)
	End Function
	Public Function getSelectSql()
		DoAssignmentByRef getSelectSql,method_ReportField_getSelectSql(me,false)
	End Function
	Public Function getGroupSql()
		DoAssignmentByRef getGroupSql,method_ReportField_getGroupSql(me)
	End Function
	Public Function getOrderSql()
		DoAssignmentByRef getOrderSql,method_ReportField_getOrderSql(me)
	End Function
	Public Function getGroup_p1(ByVal data)
		DoAssignmentByRef getGroup_p1,method_ReportField_getGroup(me,data)
	End Function
	Public Function setStart_p1(ByVal start)
		DoAssignmentByRef setStart_p1,method_ReportField_setStart(me,start)
	End Function
	Public Function name()
		DoAssignmentByRef name,method_ReportField_name(me)
	End Function
	Public Function alias()
		DoAssignmentByRef alias,method_ReportField_alias(me)
	End Function
	Public Function overrideFormat()
		DoAssignmentByRef overrideFormat,method_ReportField_overrideFormat(me)
	End Function
	Public Function setCaseSensitive_p1(ByVal cs)
		DoAssignmentByRef setCaseSensitive_p1,method_ReportField_setCaseSensitive(me,cs)
	End Function
	Public Function cutNull_p2(ByRef range,ByVal checkEmty)
		DoAssignmentByRef cutNull_p2,method_ReportField_cutNull(me,range,checkEmty)
	End Function
	Public Function cutNull_p1(ByRef range)
		DoAssignmentByRef cutNull_p1,method_ReportField_cutNull(me,range,false)
	End Function
	Public Function getLtGt_p2(ByRef lt,ByRef gt)
		DoAssignmentByRef getLtGt_p2,method_ReportField_getLtGt(me,lt,gt)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"var_interval", var_interval
		setArrElement out,"var_name", var_name
		setArrElement out,"var_alias", var_alias
		setArrElement out,"var_sqlname", var_sqlname
		setArrElement out,"var_start", var_start
		setArrElement out,"var_caseSensitive", var_caseSensitive
		setArrElement out,"var_recordBasedRequest", var_recordBasedRequest
		setArrElement out,"var_rowsInSummary", var_rowsInSummary
		setArrElement out,"var_rowsInHeader", var_rowsInHeader
		setArrElement out,"var_viewFormat", var_viewFormat
		setArrElement out,"var_orderBy", var_orderBy
		setArrElement out,"var_oldAlgorithm", var_oldAlgorithm
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment var_interval, dict("var_interval")
		DoAssignment var_name, dict("var_name")
		DoAssignment var_alias, dict("var_alias")
		DoAssignment var_sqlname, dict("var_sqlname")
		DoAssignment var_start, dict("var_start")
		DoAssignment var_caseSensitive, dict("var_caseSensitive")
		DoAssignment var_recordBasedRequest, dict("var_recordBasedRequest")
		DoAssignment var_rowsInSummary, dict("var_rowsInSummary")
		DoAssignment var_rowsInHeader, dict("var_rowsInHeader")
		DoAssignment var_viewFormat, dict("var_viewFormat")
		DoAssignment var_orderBy, dict("var_orderBy")
		DoAssignment var_oldAlgorithm, dict("var_oldAlgorithm")
	End Sub
' end serialize
End Class
' ReportCharField implementation 
Function method_ReportCharField_ReportCharField(byref this_object,ByVal name,ByVal interval,ByVal alias)
	this_object.var_interval = 0
	this_object.var_name = ""
	this_object.var_alias = ""
	this_object.var_sqlname = ""
	this_object.var_start = 0
	this_object.var_caseSensitive = false
	this_object.var_recordBasedRequest = false
	this_object.var_rowsInSummary = 0
	this_object.var_rowsInHeader = 0
	this_object.var_viewFormat = ""
	this_object.var_orderBy = "ASC"
	this_object.var_oldAlgorithm = false
	method_ReportField_ReportField this_object,name,interval,alias
End Function
Function method_ReportCharField_getStringSql(byref this_object)
	Dim fname
	doAssignment fname,IIF(this_object.var_oldAlgorithm,GetFullFieldName(this_object.var_name,""),cached_ffn(this_object.var_name))
	if IsLess(0,this_object.var_interval) then
		method_ReportCharField_getStringSql = ((("substring(" & CSmartStr(fname)) & ", 1, ") & CSmartStr(this_object.var_interval)) & ")"
		Exit Function
	else
		doAssignmentByRef method_ReportCharField_getStringSql,fname
		Exit Function
	end if
End Function
Function method_ReportCharField_getFieldName(byref this_object,ByVal fieldValue,ByVal data)
	Dim value
	doAssignment value,ArrayElement(data,IIF(this_object.var_recordBasedRequest,this_object.var_name,this_object.var_sqlname))
	if IsNull(value) then
		method_ReportCharField_getFieldName = "NULL"
		Exit Function
	end if
	if IsLess(0,this_object.var_interval) then
		doAssignmentByRef method_ReportCharField_getFieldName,asp_substr(value,0,this_object.var_interval)
		Exit Function
	else
		doAssignmentByRef method_ReportCharField_getFieldName,value
		Exit Function
	end if
End Function
Function method_ReportCharField_getKey(byref this_object,ByVal data)
	Dim this
	if bValue(this_object.var_recordBasedRequest) then
		if IsLess(0,this_object.var_interval) then
			if bValue(this_object.var_caseSensitive) then
				doAssignmentByRef method_ReportCharField_getKey,asp_substr(ArrayElement(data,this_object.var_name),0,this_object.var_interval)
				Exit Function
			else
				doAssignmentByRef method_ReportCharField_getKey,asp_strtolower(asp_substr(ArrayElement(data,this_object.var_name),0,this_object.var_interval))
				Exit Function
			end if
		else
			if bValue(this_object.var_caseSensitive) then
				doAssignmentByRef method_ReportCharField_getKey,ArrayElement(data,this_object.var_name)
				Exit Function
			else
				doAssignmentByRef method_ReportCharField_getKey,asp_strtolower(ArrayElement(data,this_object.var_name))
				Exit Function
			end if
		end if
	else
		if bValue(this_object.var_caseSensitive) then
			doAssignmentByRef method_ReportCharField_getKey,ArrayElement(data,this_object.alias())
			Exit Function
		else
			doAssignmentByRef method_ReportCharField_getKey,asp_strtolower(ArrayElement(data,this_object.alias()))
			Exit Function
		end if
	end if
End Function
Function method_ReportCharField_getWhereSql(byref this_object,ByVal groups)
	Dim ret,ssql,this,hasNull,gr,g,lt,gt
	ret = ""
	doAssignmentByRef ssql,this_object.getStringSql()
	doAssignmentByRef hasNull,this_object.cutNull_p1(groups)
	if IsLess(0,asp_count(groups)) then
		Set gr = (CreateDictionary())
		GetCollectionBounds groups,c_reportlib_loopIdx3,c_reportlib_loopMax3
		do while c_reportlib_loopIdx3<=c_reportlib_loopMax3
			c_reportlib_arrKey3 = GetCollectionKey(groups,c_reportlib_loopIdx3)
			doAssignment g,ArrayElement(groups,c_reportlib_arrKey3)
			setArrElement gr,asp_count(gr),("'" & CSmartStr(g)) & "'"
			c_reportlib_loopIdx3=c_reportlib_loopIdx3+1
		loop
		lt = ""
		gt = ""
		this_object.getLtGt_p2 lt,gt
		ret = ((((((((("(" & CSmartStr(ssql)) & CSmartStr(gt)) & "'") & CSmartStr(db_addslashes(ArrayElement(groups,0)))) & "' AND ") & CSmartStr(ssql)) & CSmartStr(lt)) & "'") & CSmartStr(db_addslashes(ArrayElement(groups,CSmartDbl(asp_count(groups))-1)))) & "')"
	end if
	if bValue(hasNull) then
		ret = CSmartStr(ret) & ((CSmartStr(IIF(ret," OR ","")) & CSmartStr(ssql)) & " IS NULL")
		ret = CSmartStr(ret) & ((" OR " & CSmartStr(ssql)) & "=''")
	end if
	doAssignmentByRef method_ReportCharField_getWhereSql,IIF(ret,("(" & CSmartStr(ret)) & ")","")
	Exit Function
End Function
setArrElement report_classfields,"date","ReportDateField"

'------ Class ReportDateField extends ReportField ------
Class ReportDateField
	Public var_interval
	Public var_name
	Public var_alias
	Public var_sqlname
	Public var_start
	Public var_caseSensitive
	Public var_recordBasedRequest
	Public var_rowsInSummary
	Public var_rowsInHeader
	Public var_viewFormat
	Public var_orderBy
	Public var_oldAlgorithm
	Public Function init_ReportDateField_p3(ByVal name,ByVal interval,ByVal alias)
		DoAssignmentByRef init_ReportDateField_p3,method_ReportDateField_ReportDateField(me,name,interval,alias)
	End Function
	Public Function setStart_p1(ByVal start)
		DoAssignmentByRef setStart_p1,method_ReportDateField_setStart(me,start)
	End Function
	Public Function getSqlList_p1(ByVal all)
		DoAssignmentByRef getSqlList_p1,method_ReportDateField_getSqlList(me,all)
	End Function
	Public Function getSqlList()
		DoAssignmentByRef getSqlList,method_ReportDateField_getSqlList(me,true)
	End Function
	Public Function getSelectSql_p1(ByVal hasGrouping)
		DoAssignmentByRef getSelectSql_p1,method_ReportDateField_getSelectSql(me,hasGrouping)
	End Function
	Public Function getGroupSql()
		DoAssignmentByRef getGroupSql,method_ReportDateField_getGroupSql(me)
	End Function
	Public Function getOrderSql()
		DoAssignmentByRef getOrderSql,method_ReportDateField_getOrderSql(me)
	End Function
	Public Function getWhereSql_p1(ByVal groups)
		DoAssignmentByRef getWhereSql_p1,method_ReportDateField_getWhereSql(me,groups)
	End Function
	Public Function getFieldName_p2(ByVal fieldValue,ByVal data)
		DoAssignmentByRef getFieldName_p2,method_ReportDateField_getFieldName(me,fieldValue,data)
	End Function
	Public Function getGroup_p1(ByVal data)
		DoAssignmentByRef getGroup_p1,method_ReportDateField_getGroup(me,data)
	End Function
	Public Function getKey_p1(ByVal data)
		DoAssignmentByRef getKey_p1,method_ReportDateField_getKey(me,data)
	End Function
	Public Function overrideFormat()
		DoAssignmentByRef overrideFormat,method_ReportDateField_overrideFormat(me)
	End Function
	Public Function cutNull_p2(ByRef range,ByVal checkEmpty)
		DoAssignmentByRef cutNull_p2,method_ReportDateField_cutNull(me,range,checkEmpty)
	End Function
	Public Function cutNull_p1(ByRef range)
		DoAssignmentByRef cutNull_p1,method_ReportDateField_cutNull(me,range,false)
	End Function
	Public Function getStringSql()
		DoAssignmentByRef getStringSql,method_ReportField_getStringSql(me)
	End Function
	Public Function name()
		DoAssignmentByRef name,method_ReportField_name(me)
	End Function
	Public Function alias()
		DoAssignmentByRef alias,method_ReportField_alias(me)
	End Function
	Public Function setCaseSensitive_p1(ByVal cs)
		DoAssignmentByRef setCaseSensitive_p1,method_ReportField_setCaseSensitive(me,cs)
	End Function
	Public Function getLtGt_p2(ByRef lt,ByRef gt)
		DoAssignmentByRef getLtGt_p2,method_ReportField_getLtGt(me,lt,gt)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"var_interval", var_interval
		setArrElement out,"var_name", var_name
		setArrElement out,"var_alias", var_alias
		setArrElement out,"var_sqlname", var_sqlname
		setArrElement out,"var_start", var_start
		setArrElement out,"var_caseSensitive", var_caseSensitive
		setArrElement out,"var_recordBasedRequest", var_recordBasedRequest
		setArrElement out,"var_rowsInSummary", var_rowsInSummary
		setArrElement out,"var_rowsInHeader", var_rowsInHeader
		setArrElement out,"var_viewFormat", var_viewFormat
		setArrElement out,"var_orderBy", var_orderBy
		setArrElement out,"var_oldAlgorithm", var_oldAlgorithm
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment var_interval, dict("var_interval")
		DoAssignment var_name, dict("var_name")
		DoAssignment var_alias, dict("var_alias")
		DoAssignment var_sqlname, dict("var_sqlname")
		DoAssignment var_start, dict("var_start")
		DoAssignment var_caseSensitive, dict("var_caseSensitive")
		DoAssignment var_recordBasedRequest, dict("var_recordBasedRequest")
		DoAssignment var_rowsInSummary, dict("var_rowsInSummary")
		DoAssignment var_rowsInHeader, dict("var_rowsInHeader")
		DoAssignment var_viewFormat, dict("var_viewFormat")
		DoAssignment var_orderBy, dict("var_orderBy")
		DoAssignment var_oldAlgorithm, dict("var_oldAlgorithm")
	End Sub
' end serialize
End Class
' ReportDateField implementation 
Function method_ReportDateField_ReportDateField(byref this_object,ByVal name,ByVal interval,ByVal alias)
	this_object.var_interval = 0
	this_object.var_name = ""
	this_object.var_alias = ""
	this_object.var_sqlname = ""
	this_object.var_start = 0
	this_object.var_caseSensitive = false
	this_object.var_recordBasedRequest = false
	this_object.var_rowsInSummary = 0
	this_object.var_rowsInHeader = 0
	this_object.var_viewFormat = ""
	this_object.var_orderBy = "ASC"
	this_object.var_oldAlgorithm = false
	method_ReportField_ReportField this_object,name,interval,alias
End Function
Function method_ReportDateField_setStart(byref this_object,ByVal start)
	Dim this
	doClassAssignment this_object,"var_start",start
	if IsEqual(this_object.var_interval,0) then
		doClassAssignment this_object,"var_sqlname",this_object.alias()
	else
		this_object.var_sqlname = CSmartStr(this_object.alias()) & "MIN"
	end if
	method_ReportDateField_setStart = CSmartDbl(start)+CSmartDbl(IIF(IsLess(0,this_object.var_interval),this_object.var_interval,1))
	Exit Function
End Function
Function method_ReportDateField_getSqlList(byref this_object,ByVal all)
	Dim grp,fname,symbols,idx
	Set grp = (CreateDictionary())
	doAssignment fname,IIF(this_object.var_oldAlgorithm,GetFullFieldName(this_object.var_name,""),cached_ffn(this_object.var_name))
	if bValue(all) then
		Set symbols = (CreateDictionary7(Empty,CreateDictionary3(Empty,"DATEPART(year, ",Empty,")",Empty,-1),Empty,CreateDictionary3(Empty,"DATEPART(quarter, ",Empty,")",Empty,0),Empty,CreateDictionary3(Empty,"DATEPART(month, ",Empty,")",Empty,0),Empty,CreateDictionary3(Empty,"DATEPART(week, ",Empty,")",Empty,0),Empty,CreateDictionary3(Empty,"DATEPART(day, ",Empty,")",Empty,2),Empty,CreateDictionary3(Empty,"DATEPART(hour, ",Empty,")",Empty,4),Empty,CreateDictionary3(Empty,"DATEPART(minute, ",Empty,")",Empty,5)))
		idx = CSmartDbl(this_object.var_interval)-1
		do
			asp_array_unshift grp,(CSmartStr(ArrayElement(ArrayElement(symbols,idx),0)) & CSmartStr(cached_ffn(this_object.var_name))) & CSmartStr(ArrayElement(ArrayElement(symbols,idx),1))
			doAssignment idx,ArrayElement(ArrayElement(symbols,idx),2)
		loop while IsLessOrEqual(0,idx)
	end if
	doAssignmentByRef method_ReportDateField_getSqlList,grp
	Exit Function
End Function
Function method_ReportDateField_getSelectSql(byref this_object,ByVal hasGrouping)
	Dim fname,this,grp,nCnt
	doAssignment fname,IIF(this_object.var_oldAlgorithm,GetFullFieldName(this_object.var_name,""),cached_ffn(this_object.var_name))
	if IsEqual(this_object.var_interval,0) then
		method_ReportDateField_getSelectSql = (CSmartStr(fname) & " as ") & CSmartStr(cached_ffn(this_object.alias()))
		Exit Function
	else
		doAssignmentByRef grp,this_object.getSqlList()
		nCnt = 0
		do while IsLess(nCnt,asp_count(grp))
			setArrElement grp,nCnt,CSmartStr(ArrayElement(grp,nCnt)) & (" as " & CSmartStr(cached_ffn(CSmartStr(this_object.var_alias) & CSmartStr(CSmartDbl(nCnt)+CSmartDbl(this_object.var_start)))))
			nCnt = CSmartDbl(nCnt)+1
		loop
		if bValue(hasGrouping) then
			setArrElement grp,asp_count(grp),(("MIN(" & CSmartStr(fname)) & ") as ") & CSmartStr(cached_ffn(CSmartStr(this_object.alias()) & "MIN"))
			setArrElement grp,asp_count(grp),(("MAX(" & CSmartStr(fname)) & ") as ") & CSmartStr(cached_ffn(CSmartStr(this_object.alias()) & "MAX"))
		else
			setArrElement grp,asp_count(grp),(CSmartStr(fname) & " as ") & CSmartStr(cached_ffn(CSmartStr(this_object.alias()) & "MIN"))
		end if
		doAssignmentByRef method_ReportDateField_getSelectSql,asp_join(", ",grp)
		Exit Function
	end if
End Function
Function method_ReportDateField_getGroupSql(byref this_object)
	Dim grp,this
	if IsEqual(this_object.var_interval,0) then
		doAssignmentByRef method_ReportDateField_getGroupSql,cached_ffn(this_object.var_name)
		Exit Function
	else
		doAssignmentByRef grp,this_object.getSqlList()
		doAssignmentByRef method_ReportDateField_getGroupSql,asp_join(", ",grp)
		Exit Function
	end if
End Function
Function method_ReportDateField_getOrderSql(byref this_object)
	Dim fname,grp,this,newgrp,g
	if IsEqual(this_object.var_interval,0) then
		doAssignment fname,IIF(this_object.var_oldAlgorithm,GetFullFieldName(this_object.var_name,""),cached_ffn(this_object.var_name))
		method_ReportDateField_getOrderSql = ((CSmartStr(fname) & " ") & CSmartStr(this_object.var_orderBy)) & " "
		Exit Function
	else
		doAssignmentByRef grp,this_object.getSqlList()
		Set newgrp = (CreateDictionary())
		GetCollectionBounds grp,c_reportlib_loopIdx6,c_reportlib_loopMax6
		do while c_reportlib_loopIdx6<=c_reportlib_loopMax6
			c_reportlib_arrKey6 = GetCollectionKey(grp,c_reportlib_loopIdx6)
			doAssignment g,ArrayElement(grp,c_reportlib_arrKey6)
			setArrElement newgrp,asp_count(newgrp),((CSmartStr(g) & " ") & CSmartStr(this_object.var_orderBy)) & " "
			c_reportlib_loopIdx6=c_reportlib_loopIdx6+1
		loop
		doAssignmentByRef method_ReportDateField_getOrderSql,asp_join(", ",newgrp)
		Exit Function
	end if
End Function
Function method_ReportDateField_getWhereSql(byref this_object,ByVal groups)
	Dim ret,hasNull,this,lt,gt
	ret = ""
	doAssignmentByRef hasNull,this_object.cutNull_p2(groups,true)
	if IsLess(0,asp_count(groups)) then
		lt = ""
		gt = ""
		if IsEqual(this_object.var_interval,0) then
			this_object.getLtGt_p2 lt,gt
			ret = ((((((((((("(" & CSmartStr(cached_ffn(this_object.var_name))) & " ") & CSmartStr(gt)) & " ") & CSmartStr(db_datequotes(ArrayElement(groups,0)))) & " AND ") & CSmartStr(cached_ffn(this_object.var_name))) & " ") & CSmartStr(lt)) & " ") & CSmartStr(db_datequotes(ArrayElement(groups,CSmartDbl(asp_count(groups))-1)))) & ")"
		else
			if IsLessOrEqual(ArrayElement(ArrayElement(groups,0),"MIN"),ArrayElement(ArrayElement(groups,CSmartDbl(asp_count(groups))-1),"MAX")) then
				ret = ((((((("(" & CSmartStr(cached_ffn(this_object.var_name))) & " >= ") & CSmartStr(db_datequotes(ArrayElement(ArrayElement(groups,0),"MIN")))) & " AND ") & CSmartStr(cached_ffn(this_object.var_name))) & " <= ") & CSmartStr(db_datequotes(ArrayElement(ArrayElement(groups,CSmartDbl(asp_count(groups))-1),"MAX")))) & ")"
			else
				ret = ((((((("(" & CSmartStr(cached_ffn(this_object.var_name))) & " <= ") & CSmartStr(db_datequotes(ArrayElement(ArrayElement(groups,0),"MAX")))) & " AND ") & CSmartStr(cached_ffn(this_object.var_name))) & " >= ") & CSmartStr(db_datequotes(ArrayElement(ArrayElement(groups,CSmartDbl(asp_count(groups))-1),"MIN")))) & ")"
			end if
		end if
	end if
	if bValue(hasNull) then
		ret = CSmartStr(ret) & ((CSmartStr(IIF(ret," OR ","")) & CSmartStr(cached_ffn(this_object.var_name))) & " IS NULL ")
	end if
	doAssignmentByRef method_ReportDateField_getWhereSql,IIF(ret,("(" & CSmartStr(ret)) & ")","")
	Exit Function
End Function
Function method_ReportDateField_getFieldName(byref this_object,ByVal fieldValue,ByVal data)
	Dim value,date
	doAssignment value,ArrayElement(data,IIF(this_object.var_recordBasedRequest,this_object.var_name,this_object.var_sqlname))
	if (IsNull(value) or not bValue(value)) or IsEqual(strcasecmp(value,"null"),0) then
		method_ReportDateField_getFieldName = "NULL"
		Exit Function
	end if
	if IsEqual(this_object.var_interval,0) then
		if bValue(this_object.var_viewFormat) then
			doAssignmentByRef method_ReportDateField_getFieldName,GetDataInt(value,CreateDictionary1(this_object.var_name,value),this_object.var_name,this_object.var_viewFormat)
			Exit Function
		else
			doAssignmentByRef date,db2time(value)
			doAssignmentByRef method_ReportDateField_getFieldName,str_format_datetime(date)
			Exit Function
		end if
	end if
	do
		If IsEqual(this_object.var_interval,1) then
			doAssignmentByRef date,cached_db2time(value)
			doAssignmentByRef method_ReportDateField_getFieldName,ArrayElement(date,0)
			Exit Function
		End If
		If IsEqual(this_object.var_interval,1) or IsEqual(this_object.var_interval,2) then
			doAssignmentByRef date,cached_db2time(value)
			method_ReportDateField_getFieldName = (CSmartStr(ArrayElement(date,0)) & "/Q") & CSmartStr(asp_intval(CSmartDbl(ArrayElement(date,1))/3))
			Exit Function
		End If
		If IsEqual(this_object.var_interval,1) or IsEqual(this_object.var_interval,2) or IsEqual(this_object.var_interval,3) then
			doAssignmentByRef date,cached_db2time(value)
			method_ReportDateField_getFieldName = (CSmartStr(ArrayElement(locale_info,"LOCALE_SABBREVMONTHNAME" & CSmartStr(ArrayElement(date,1)))) & " ") & CSmartStr(ArrayElement(date,0))
			Exit Function
		End If
		If IsEqual(this_object.var_interval,1) or IsEqual(this_object.var_interval,2) or IsEqual(this_object.var_interval,3) or IsEqual(this_object.var_interval,4) then
			doAssignmentByRef method_ReportDateField_getFieldName,cached_formatweekstart(value)
			Exit Function
		End If
		If IsEqual(this_object.var_interval,1) or IsEqual(this_object.var_interval,2) or IsEqual(this_object.var_interval,3) or IsEqual(this_object.var_interval,4) or IsEqual(this_object.var_interval,5) then
			doAssignmentByRef date,cached_db2time(value)
			doAssignmentByRef method_ReportDateField_getFieldName,format_shortdate(date)
			Exit Function
		End If
		If IsEqual(this_object.var_interval,1) or IsEqual(this_object.var_interval,2) or IsEqual(this_object.var_interval,3) or IsEqual(this_object.var_interval,4) or IsEqual(this_object.var_interval,5) or IsEqual(this_object.var_interval,6) then
			doAssignmentByRef date,db2time(value)
			setArrElement date,4,0
			setArrElement date,5,0
			doAssignmentByRef method_ReportDateField_getFieldName,str_format_datetime(date)
			Exit Function
		End If
		If IsEqual(this_object.var_interval,1) or IsEqual(this_object.var_interval,2) or IsEqual(this_object.var_interval,3) or IsEqual(this_object.var_interval,4) or IsEqual(this_object.var_interval,5) or IsEqual(this_object.var_interval,6) or IsEqual(this_object.var_interval,7) then
			doAssignmentByRef date,db2time(value)
			setArrElement date,5,0
			doAssignmentByRef method_ReportDateField_getFieldName,str_format_datetime(date)
			Exit Function
		End If
	Loop While false
End Function
Function method_ReportDateField_getGroup(byref this_object,ByVal data)
	Dim this
	if IsEqual(this_object.var_interval,0) then
		doAssignmentByRef method_ReportDateField_getGroup,ArrayElement(data,this_object.alias())
		Exit Function
	else
		if IsNull(ArrayElement(data,CSmartStr(this_object.alias()) & "MIN")) or IsNull(ArrayElement(data,CSmartStr(this_object.alias()) & "MAX")) then
			method_ReportDateField_getGroup = null
			Exit Function
		else
			doAssignmentByRef method_ReportDateField_getGroup,CreateDictionary2("MIN",ArrayElement(data,CSmartStr(this_object.alias()) & "MIN"),"MAX",ArrayElement(data,CSmartStr(this_object.alias()) & "MAX"))
			Exit Function
		end if
	end if
End Function
Function method_ReportDateField_getKey(byref this_object,ByVal data)
	Dim this,key,nCnt,strdate,date,start
	if not bValue(this_object.var_recordBasedRequest) then
		if IsEqual(this_object.var_interval,0) then
			doAssignmentByRef method_ReportDateField_getKey,ArrayElement(data,this_object.alias())
			Exit Function
		else
			Set key = (CreateDictionary())
			doAssignment nCnt,this_object.var_start
			do while IsLess(nCnt,CSmartDbl(this_object.var_interval)+CSmartDbl(this_object.var_start))
				setArrElement key,asp_count(key),ArrayElement(data,CSmartStr(this_object.var_alias) & CSmartStr(nCnt))
				nCnt = CSmartDbl(nCnt)+1
			loop
			doAssignmentByRef method_ReportDateField_getKey,asp_join("-",key)
			Exit Function
		end if
	else
		doAssignment strdate,ArrayElement(data,this_object.var_name)
		if IsNull(strdate) then
			method_ReportDateField_getKey = "NULL"
			Exit Function
		end if
		if IsEqual(this_object.var_interval,0) then
			doAssignmentByRef method_ReportDateField_getKey,strdate
			Exit Function
		else
			do
				If IsEqual(this_object.var_interval,1) then
					doAssignmentByRef date,cached_db2time(strdate)
					doAssignmentByRef method_ReportDateField_getKey,ArrayElement(date,0)
					Exit Function
				End If
				If IsEqual(this_object.var_interval,1) or IsEqual(this_object.var_interval,2) then
					doAssignmentByRef date,cached_db2time(strdate)
					method_ReportDateField_getKey = (CSmartStr(ArrayElement(date,0)) & "-") & CSmartStr(asp_intval(CSmartDbl(ArrayElement(date,1))/3))
					Exit Function
				End If
				If IsEqual(this_object.var_interval,1) or IsEqual(this_object.var_interval,2) or IsEqual(this_object.var_interval,3) then
					doAssignmentByRef date,cached_db2time(strdate)
					method_ReportDateField_getKey = (CSmartStr(ArrayElement(date,0)) & "-") & CSmartStr(ArrayElement(date,1))
					Exit Function
				End If
				If IsEqual(this_object.var_interval,1) or IsEqual(this_object.var_interval,2) or IsEqual(this_object.var_interval,3) or IsEqual(this_object.var_interval,4) then
					doAssignmentByRef start,cached_getweekstart(strdate)
					method_ReportDateField_getKey = (((CSmartStr(ArrayElement(start,0)) & "-") & CSmartStr(ArrayElement(start,1))) & "-") & CSmartStr(ArrayElement(start,2))
					Exit Function
				End If
				If IsEqual(this_object.var_interval,1) or IsEqual(this_object.var_interval,2) or IsEqual(this_object.var_interval,3) or IsEqual(this_object.var_interval,4) or IsEqual(this_object.var_interval,5) then
					doAssignmentByRef date,cached_db2time(strdate)
					method_ReportDateField_getKey = (((CSmartStr(ArrayElement(date,0)) & "-") & CSmartStr(ArrayElement(date,1))) & "-") & CSmartStr(ArrayElement(date,2))
					Exit Function
				End If
				If IsEqual(this_object.var_interval,1) or IsEqual(this_object.var_interval,2) or IsEqual(this_object.var_interval,3) or IsEqual(this_object.var_interval,4) or IsEqual(this_object.var_interval,5) or IsEqual(this_object.var_interval,6) then
					doAssignmentByRef date,db2time(strdate)
					method_ReportDateField_getKey = (((((CSmartStr(ArrayElement(date,0)) & "-") & CSmartStr(ArrayElement(date,1))) & "-") & CSmartStr(ArrayElement(date,2))) & "-") & CSmartStr(ArrayElement(date,3))
					Exit Function
				End If
				If IsEqual(this_object.var_interval,1) or IsEqual(this_object.var_interval,2) or IsEqual(this_object.var_interval,3) or IsEqual(this_object.var_interval,4) or IsEqual(this_object.var_interval,5) or IsEqual(this_object.var_interval,6) or IsEqual(this_object.var_interval,7) then
					doAssignmentByRef date,db2time(strdate)
					method_ReportDateField_getKey = (((((((CSmartStr(ArrayElement(date,0)) & "-") & CSmartStr(ArrayElement(date,1))) & "-") & CSmartStr(ArrayElement(date,2))) & "-") & CSmartStr(ArrayElement(date,3))) & "-") & CSmartStr(ArrayElement(date,4))
					Exit Function
				End If
			Loop While false
		end if
	end if
End Function
Function method_ReportDateField_overrideFormat(byref this_object)
	method_ReportDateField_overrideFormat = true
	Exit Function
End Function
Function method_ReportDateField_cutNull(byref this_object,ByRef range,ByVal checkEmpty)
	Dim ret,out,nCnt,b
	ret = false
	Set out = (CreateDictionary())
	nCnt = 0
	do while IsLess(nCnt,asp_count(range))
		b = false
		if IsNull(ArrayElement(range,nCnt)) then
			b = true
			ret = true
		else
			if bValue(checkEmpty) then
				if bValue(asp_is_array(ArrayElement(range,nCnt))) then
					if not bValue(ArrayElement(ArrayElement(range,nCnt),"MIN")) or IsEqual(strcasecmp(ArrayElement(ArrayElement(range,nCnt),"MIN"),"null"),0) then
						b = true
						ret = true
					end if
				else
					if not bValue(ArrayElement(range,nCnt)) or IsEqual(strcasecmp(ArrayElement(range,nCnt),"null"),0) then
						b = true
						ret = true
					end if
				end if
			end if
		end if
		if not bValue(b) then
			setArrElement out,asp_count(out),ArrayElement(range,nCnt)
		end if
		nCnt = CSmartDbl(nCnt)+1
	loop
	doAssignment range,out
	doAssignmentByRef method_ReportDateField_cutNull,ret
	Exit Function
End Function
Function getFormattedValue(ByVal value,ByVal fieldName,ByVal strViewFormat,ByVal strEditFormat,ByVal mode)
	Dim val,d,h,m,s
	if IsEqual(strViewFormat,FORMAT_TIME) and bValue(asp_is_numeric(value)) then
		val = ""
		doAssignmentByRef d,asp_intval(CSmartDbl(value)/86400)
		doAssignmentByRef h,asp_intval((value mod 86400)/3600)
		doAssignmentByRef m,asp_intval(((value mod 86400) mod 3600)/60)
		s = ((value mod 86400) mod 3600) mod 60
		val = CSmartStr(val) & CSmartStr(IIF(IsLess(0,d),CSmartStr(d) & "d ",""))
		val = CSmartStr(val) & CSmartStr(str_format_time(CreateDictionary6(Empty,0,Empty,0,Empty,0,Empty,h,Empty,m,Empty,s)))
	else
		if IsEqual(strEditFormat,EDIT_FORMAT_LOOKUP_WIZARD) or IsEqual(strEditFormat,EDIT_FORMAT_RADIO) then
			doAssignmentByRef val,DisplayLookupWizard(fieldName,value,null,"",mode)
		else
			doAssignmentByRef val,GetDataInt(value,CreateDictionary1(fieldName,value),fieldName,strViewFormat)
		end if
	end if
	doAssignmentByRef getFormattedValue,val
	Exit Function
End Function
Set cache_db2time = (CreateDictionary())
Function cached_db2time(ByVal strtime)
	Dim res
	if not (not IsEmpty(ArrayElement(cache_db2time,strtime))) then
		doAssignmentByRef res,db2time(strtime)
		setArrElement cache_db2time,strtime,res
		doAssignmentByRef cached_db2time,res
		Exit Function
	else
		doAssignmentByRef cached_db2time,ArrayElement(cache_db2time,strtime)
		Exit Function
	end if
End Function
Set cache_getdayofweek = (CreateDictionary())
Function cached_getdayofweek(ByVal strtime)
	Dim date,res
	if not (not IsEmpty(ArrayElement(cache_getdayofweek,strtime))) then
		doAssignmentByRef date,cached_db2time(strtime)
		doAssignmentByRef res,getdayofweek(date)
		setArrElement cache_getdayofweek,strtime,res
		doAssignmentByRef cached_getdayofweek,res
		Exit Function
	else
		doAssignmentByRef cached_getdayofweek,ArrayElement(cache_getdayofweek,strtime)
		Exit Function
	end if
End Function
Set cache_getweekstart = (CreateDictionary())
Function cached_getweekstart(ByVal strtime)
	Dim date,res
	if not (not IsEmpty(ArrayElement(cache_getweekstart,strtime))) then
		doAssignmentByRef date,cached_db2time(strtime)
		doAssignmentByRef res,getweekstart(date)
		setArrElement cache_getweekstart,strtime,res
		doAssignmentByRef cached_getweekstart,res
		Exit Function
	else
		doAssignmentByRef cached_getweekstart,ArrayElement(cache_getweekstart,strtime)
		Exit Function
	end if
End Function
Set cache_formatweekstart = (CreateDictionary())
Function cached_formatweekstart(ByVal strtime)
	Dim start,var_end,res
	if not (not IsEmpty(ArrayElement(cache_formatweekstart,strtime))) then
		doAssignmentByRef start,cached_getweekstart(strtime)
		doAssignmentByRef var_end,adddays(start,6)
		res = (CSmartStr(format_shortdate(start)) & " - ") & CSmartStr(format_shortdate(var_end))
		setArrElement cache_formatweekstart,strtime,res
		doAssignmentByRef cached_formatweekstart,res
		Exit Function
	else
		doAssignmentByRef cached_formatweekstart,ArrayElement(cache_formatweekstart,strtime)
		Exit Function
	end if
End Function
Set cache_fullfieldname = (CreateDictionary())
Function cached_ffn(ByVal field)
	Dim res
	if not (not IsEmpty(ArrayElement(cache_fullfieldname,field))) then
		doAssignmentByRef res,AddFieldWrappers(field)
		setArrElement cache_fullfieldname,field,res
		doAssignmentByRef cached_ffn,res
		Exit Function
	else
		doAssignmentByRef cached_ffn,ArrayElement(cache_fullfieldname,field)
		Exit Function
	end if
End Function

'------ Class SQLStatement ------
Class SQLStatement
	Public var_fields
	Public var_hasDetails
	Public var_originalSql
	Public var_order_in
	Public var_order_out
	Public var_order_old
	Public var_aggregates
	Public var_skipCount
	Public var_reportGlobalSummary
	Public var_reportSummary
	Public var_details
	Public var_from
	Public var_pagesize
	Public var_limitLevel
	Public var_hasGroups
	Public searchClauseObj
	Public var_recordBasedRequest
	Public var_oldAlgorithm
	Public tName
	Public shortTName
	Public repGroupFieldsCount
	Public repPageSummary
	Public repGlobalSummary
	Public repLayout
	Public showGroupSummaryCount
	Public repShowDet
	Public repGroupFields
	Public tKeyFields
	Public isExistTotalFields
	Public fieldsArr
	Public orderIndexes
	Public Function init_SQLStatement_p6(ByVal sql,ByVal order,ByVal pagesize,ByVal connection,ByRef searchClauseObj,ByRef params)
		DoAssignmentByRef init_SQLStatement_p6,method_SQLStatement_SQLStatement(me,sql,order,pagesize,connection,searchClauseObj,params)
	End Function
	Public Function getOriginal_p1(ByVal useOriginalOrder)
		DoAssignmentByRef getOriginal_p1,method_SQLStatement_getOriginal(me,useOriginalOrder)
	End Function
	Public Function getOriginal()
		DoAssignmentByRef getOriginal,method_SQLStatement_getOriginal(me,true)
	End Function
	Public Function setRecordBasedRequest_p1(ByVal recordBasedRequest)
		DoAssignmentByRef setRecordBasedRequest_p1,method_SQLStatement_setRecordBasedRequest(me,recordBasedRequest)
	End Function
	Public Function getGroup_p1(ByVal data)
		DoAssignmentByRef getGroup_p1,method_SQLStatement_getGroup(me,data)
	End Function
	Public Function field_p1(ByVal num)
		DoAssignmentByRef field_p1,method_SQLStatement_field(me,num)
	End Function
	Public Function getSQLLimits_p2(ByVal sql,ByVal from)
		DoAssignmentByRef getSQLLimits_p2,method_SQLStatement_getSQLLimits(me,sql,from)
	End Function
	Public Function sqlg_p2(ByVal donotlimit,ByVal doorder)
		DoAssignmentByRef sqlg_p2,method_SQLStatement_sqlg(me,donotlimit,doorder)
	End Function
	Public Function sqlg_p1(ByVal donotlimit)
		DoAssignmentByRef sqlg_p1,method_SQLStatement_sqlg(me,donotlimit,true)
	End Function
	Public Function sqlg()
		DoAssignmentByRef sqlg,method_SQLStatement_sqlg(me,false,true)
	End Function
	Public Function sqlcg()
		DoAssignmentByRef sqlcg,method_SQLStatement_sqlcg(me)
	End Function
	Public Function sqlt()
		DoAssignmentByRef sqlt,method_SQLStatement_sqlt(me)
	End Function
	Public Function sql2_p1(ByVal groups)
		DoAssignmentByRef sql2_p1,method_SQLStatement_sql2(me,groups)
	End Function
	Public Function sql2()
		DoAssignmentByRef sql2,method_SQLStatement_sql2(me,null)
	End Function
	Public Function buildsql_p1(ByVal hsql)
		DoAssignmentByRef buildsql_p1,method_SQLStatement_buildsql(me,hsql)
	End Function
	Public Function initWhere()
		DoAssignmentByRef initWhere,method_SQLStatement_initWhere(me)
	End Function
	Public Function getWhere()
		DoAssignmentByRef getWhere,method_SQLStatement_getWhere(me)
	End Function
	Public Function applyWhere_p1(ByRef sql)
		DoAssignmentByRef applyWhere_p1,method_SQLStatement_applyWhere(me,sql)
	End Function
	Public Function setOldAlgorithm_p1(ByVal useOldAlgorithm)
		DoAssignmentByRef setOldAlgorithm_p1,method_SQLStatement_setOldAlgorithm(me,useOldAlgorithm)
	End Function
	Public Function setOldAlgorithm()
		DoAssignmentByRef setOldAlgorithm,method_SQLStatement_setOldAlgorithm(me,true)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"var_fields", var_fields
		setArrElement out,"var_hasDetails", var_hasDetails
		setArrElement out,"var_originalSql", var_originalSql
		setArrElement out,"var_order_in", var_order_in
		setArrElement out,"var_order_out", var_order_out
		setArrElement out,"var_order_old", var_order_old
		setArrElement out,"var_aggregates", var_aggregates
		setArrElement out,"var_skipCount", var_skipCount
		setArrElement out,"var_reportGlobalSummary", var_reportGlobalSummary
		setArrElement out,"var_reportSummary", var_reportSummary
		setArrElement out,"var_details", var_details
		setArrElement out,"var_from", var_from
		setArrElement out,"var_pagesize", var_pagesize
		setArrElement out,"var_limitLevel", var_limitLevel
		setArrElement out,"var_hasGroups", var_hasGroups
		setArrElement out,"searchClauseObj", searchClauseObj
		setArrElement out,"var_recordBasedRequest", var_recordBasedRequest
		setArrElement out,"var_oldAlgorithm", var_oldAlgorithm
		setArrElement out,"tName", tName
		setArrElement out,"shortTName", shortTName
		setArrElement out,"repGroupFieldsCount", repGroupFieldsCount
		setArrElement out,"repPageSummary", repPageSummary
		setArrElement out,"repGlobalSummary", repGlobalSummary
		setArrElement out,"repLayout", repLayout
		setArrElement out,"showGroupSummaryCount", showGroupSummaryCount
		setArrElement out,"repShowDet", repShowDet
		setArrElement out,"repGroupFields", repGroupFields
		setArrElement out,"tKeyFields", tKeyFields
		setArrElement out,"isExistTotalFields", isExistTotalFields
		setArrElement out,"fieldsArr", fieldsArr
		setArrElement out,"orderIndexes", orderIndexes
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment var_fields, dict("var_fields")
		DoAssignment var_hasDetails, dict("var_hasDetails")
		DoAssignment var_originalSql, dict("var_originalSql")
		DoAssignment var_order_in, dict("var_order_in")
		DoAssignment var_order_out, dict("var_order_out")
		DoAssignment var_order_old, dict("var_order_old")
		DoAssignment var_aggregates, dict("var_aggregates")
		DoAssignment var_skipCount, dict("var_skipCount")
		DoAssignment var_reportGlobalSummary, dict("var_reportGlobalSummary")
		DoAssignment var_reportSummary, dict("var_reportSummary")
		DoAssignment var_details, dict("var_details")
		DoAssignment var_from, dict("var_from")
		DoAssignment var_pagesize, dict("var_pagesize")
		DoAssignment var_limitLevel, dict("var_limitLevel")
		DoAssignment var_hasGroups, dict("var_hasGroups")
		DoAssignment searchClauseObj, dict("searchClauseObj")
		DoAssignment var_recordBasedRequest, dict("var_recordBasedRequest")
		DoAssignment var_oldAlgorithm, dict("var_oldAlgorithm")
		DoAssignment tName, dict("tName")
		DoAssignment shortTName, dict("shortTName")
		DoAssignment repGroupFieldsCount, dict("repGroupFieldsCount")
		DoAssignment repPageSummary, dict("repPageSummary")
		DoAssignment repGlobalSummary, dict("repGlobalSummary")
		DoAssignment repLayout, dict("repLayout")
		DoAssignment showGroupSummaryCount, dict("showGroupSummaryCount")
		DoAssignment repShowDet, dict("repShowDet")
		DoAssignment repGroupFields, dict("repGroupFields")
		DoAssignment tKeyFields, dict("tKeyFields")
		DoAssignment isExistTotalFields, dict("isExistTotalFields")
		DoAssignment fieldsArr, dict("fieldsArr")
		DoAssignment orderIndexes, dict("orderIndexes")
	End Sub
' end serialize
End Class
' SQLStatement implementation 
Function method_SQLStatement_SQLStatement(byref this_object,ByVal sql,ByVal order,ByVal pagesize,ByVal connection,ByRef searchClauseObj,ByRef params)
	doClassAssignment this_object,"var_fields",CreateDictionary()
	this_object.var_hasDetails = true
	this_object.var_originalSql = null
	doClassAssignment this_object,"var_aggregates",CreateDictionary()
	this_object.var_skipCount = 0
	this_object.var_reportGlobalSummary = true
	this_object.var_reportSummary = true
	this_object.var_details = true
	this_object.var_from = 0
	this_object.var_pagesize = 20
	this_object.var_limitLevel = 0
	this_object.var_hasGroups = true
	this_object.searchClauseObj = null
	this_object.var_recordBasedRequest = false
	this_object.var_oldAlgorithm = false
	this_object.tName = ""
	this_object.shortTName = ""
	this_object.repGroupFieldsCount = 0
	this_object.repPageSummary = 0
	this_object.repGlobalSummary = 0
	this_object.repLayout = 0
	this_object.showGroupSummaryCount = 0
	this_object.repShowDet = 0
	doClassAssignment this_object,"repGroupFields",CreateDictionary()
	doClassAssignment this_object,"tKeyFields",CreateDictionary()
	this_object.isExistTotalFields = false
	doClassAssignment this_object,"fieldsArr",CreateDictionary()
	Dim this,start,fields,i,j,add,f,field,order_in,order_out,order_old,o,groupField,fieldIndex,n
	RunnerApply this_object,params
	doClassAssignment this_object,"searchClauseObj",searchClauseObj
	if not bValue(asp_is_array(sql)) then
		Response.End
	end if
	doClassAssignment this_object,"var_originalSql",sql
	start = 0
	Set fields = (CreateDictionary())
	i = 0
	do while IsLess(i,asp_count(this_object.repGroupFields))
		j = 0
		do while IsLess(j,asp_count(this_object.fieldsArr))
			if IsEqual(ArrayElement(ArrayElement(this_object.repGroupFields,i),"strGroupField"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"name")) then
				Set add = (CreateDictionary())
				setArrElement add,"name",ArrayElement(ArrayElement(this_object.fieldsArr,j),"name")
				if bValue(IsNumberType(GetFieldType(ArrayElement(ArrayElement(this_object.fieldsArr,j),"name"),this_object.tName))) then
					setArrElement add,"type","numeric"
				else
					if bValue(IsCharType(GetFieldType(ArrayElement(ArrayElement(this_object.fieldsArr,j),"name"),this_object.tName))) then
						setArrElement add,"type","char"
						setArrElement add,"case_sensitive",reportCaseSensitiveGroupFields
					else
						if bValue(IsDateFieldType(GetFieldType(ArrayElement(ArrayElement(this_object.fieldsArr,j),"name"),this_object.tName))) then
							setArrElement add,"type","date"
						else
							setArrElement add,"type","char"
						end if
					end if
				end if
				setArrElement add,"interval",ArrayElement(ArrayElement(this_object.repGroupFields,i),"groupInterval")
				setArrElement add,"viewformat",ArrayElement(ArrayElement(this_object.fieldsArr,j),"viewFormat")
				setArrElement add,"rowsinsummary",1
				if ((bValue(ArrayElement(ArrayElement(this_object.fieldsArr,j),"totalMax")) or bValue(ArrayElement(ArrayElement(this_object.fieldsArr,j),"totalMin"))) or bValue(ArrayElement(ArrayElement(this_object.fieldsArr,j),"totalAvg"))) or bValue(ArrayElement(ArrayElement(this_object.fieldsArr,j),"totalSum")) then
					setArrElement add,"rowsinsummary",CSmartDbl(ArrayElement(add,"rowsinsummary"))+1
				end if
				if IsEqual(this_object.repLayout,REPORT_STEPPED) then
					setArrElement add,"rowsinheader",1
				else
					if IsEqual(this_object.repLayout,REPORT_BLOCK) then
						setArrElement add,"rowsinheader",0
					else
						if IsEqual(this_object.repLayout,REPORT_OUTLINE) or IsEqual(this_object.repLayout,REPORT_ALIGN) then
							if IsEqual(j,CSmartDbl(asp_count(this_object.fieldsArr))-1) then
								setArrElement add,"rowsinheader",2
							else
								setArrElement add,"rowsinheader",1
							end if
						else
							if IsEqual(this_object.repLayout,REPORT_TABULAR) then
								setArrElement add,"rowsinheader",0
							end if
						end if
					end if
				end if
				setArrElement fields,asp_count(fields),add
			end if
			j = CSmartDbl(j)+1
		loop
		i = CSmartDbl(i)+1
	loop
	this_object.var_hasGroups = IsLess(0,asp_count(fields))
	GetCollectionBounds fields,c_reportlib_loopIdx11,c_reportlib_loopMax11
	do while c_reportlib_loopIdx11<=c_reportlib_loopMax11
		c_reportlib_arrKey11 = GetCollectionKey(fields,c_reportlib_loopIdx11)
		doAssignment field,ArrayElement(fields,c_reportlib_arrKey11)
		doAssignmentByRef f,create_reportfield(ArrayElement(field,"name"),ArrayElement(field,"type"),ArrayElement(field,"interval"),"grp")
		doAssignmentByRef start,f.setStart_p1(start)
		if not IsEmpty(ArrayElement(field,"case_sensitive")) then
			f.setCaseSensitive_p1 ArrayElement(field,"case_sensitive")
		end if
		if not IsEmpty(ArrayElement(field,"rowsinsummary")) then
			doClassAssignment f,"var_rowsInSummary",ArrayElement(field,"rowsinsummary")
		end if
		if not IsEmpty(ArrayElement(field,"rowsinheader")) then
			doClassAssignment f,"var_rowsInHeader",ArrayElement(field,"rowsinheader")
		end if
		doClassAssignment f,"var_viewFormat",ArrayElement(field,"viewformat")
		setArrElement this_object.var_fields,asp_count(this_object.var_fields),f
		c_reportlib_loopIdx11=c_reportlib_loopIdx11+1
	loop
	if bValue(order) then
		Set order_in = (CreateDictionary())
		Set order_out = (CreateDictionary())
		Set order_old = (CreateDictionary())
		GetCollectionBounds order,c_reportlib_loopIdx12,c_reportlib_loopMax12
		do while c_reportlib_loopIdx12<=c_reportlib_loopMax12
			c_reportlib_arrKey12 = GetCollectionKey(order,c_reportlib_loopIdx12)
			doAssignment o,ArrayElement(order,c_reportlib_arrKey12)
			setArrElement order_in,asp_count(order_in),(CSmartStr(ArrayElement(o,2)) & " as ") & CSmartStr(cached_ffn("originalorder" & CSmartStr(ArrayElement(o,0))))
			setArrElement order_out,asp_count(order_out),(CSmartStr(cached_ffn("originalorder" & CSmartStr(ArrayElement(o,0)))) & " ") & CSmartStr(ArrayElement(o,1))
			groupField = false
			i = 0
			do while IsLess(i,asp_count(this_object.repGroupFields))
				j = 0
				do while IsLess(j,asp_count(this_object.fieldsArr))
					if IsEqual(ArrayElement(ArrayElement(this_object.repGroupFields,i),"strGroupField"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"name")) then
						doAssignmentByRef fieldIndex,GetFieldIndex(ArrayElement(ArrayElement(this_object.repGroupFields,i),"strGroupField"),"")
						if IsEqual(fieldIndex,ArrayElement(o,0)) then
							n = CSmartDbl(ArrayElement(ArrayElement(this_object.repGroupFields,i),"groupOrder"))-1
							doClassAssignment ArrayElement(this_object.var_fields,n),"var_orderBy",ArrayElement(o,1)
							groupField = true
						end if
					end if
					j = CSmartDbl(j)+1
				loop
				i = CSmartDbl(i)+1
			loop
			if not bValue(groupField) then
				setArrElement order_old,asp_count(order_old),(CSmartStr(ArrayElement(o,2)) & " ") & CSmartStr(ArrayElement(o,1))
			end if
			c_reportlib_loopIdx12=c_reportlib_loopIdx12+1
		loop
		doClassAssignment this_object,"var_order_in",asp_join(", ",order_in)
		doClassAssignment this_object,"var_order_out",asp_join(", ",order_out)
		doClassAssignment this_object,"var_order_old",asp_join(", ",order_old)
	end if
	i = 0
	do while IsLess(i,asp_count(this_object.fieldsArr))
		if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalMax")) then
			setArrElement this_object.var_aggregates,asp_count(this_object.var_aggregates),(("MAX(" & CSmartStr(cached_ffn(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")))) & ") as ") & CSmartStr(cached_ffn(CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) & "MAX"))
		end if
		if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalMin")) then
			setArrElement this_object.var_aggregates,asp_count(this_object.var_aggregates),(("MIN(" & CSmartStr(cached_ffn(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")))) & ") as ") & CSmartStr(cached_ffn(CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) & "MIN"))
		end if
		if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalAvg")) then
			if not bValue(IsDateFieldType(GetFieldType(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),this_object.tName))) then
				setArrElement this_object.var_aggregates,asp_count(this_object.var_aggregates),(("AVG(" & CSmartStr(cached_ffn(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")))) & ") as ") & CSmartStr(cached_ffn(CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) & "AVG"))
				setArrElement this_object.var_aggregates,asp_count(this_object.var_aggregates),(("COUNT(" & CSmartStr(cached_ffn(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")))) & ") as ") & CSmartStr(cached_ffn(CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) & "NAVG"))
			end if
		end if
		if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalSum")) then
			if not bValue(IsDateFieldType(GetFieldType(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),this_object.tName))) then
				setArrElement this_object.var_aggregates,asp_count(this_object.var_aggregates),(("SUM(" & CSmartStr(cached_ffn(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")))) & ") as ") & CSmartStr(cached_ffn(CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) & "SUM"))
			end if
		end if
		i = CSmartDbl(i)+1
	loop
	this_object.var_reportSummary = bValue(this_object.repPageSummary) or bValue(this_object.repGlobalSummary)
	doClassAssignment this_object,"var_pagesize",pagesize
End Function
Function method_SQLStatement_getOriginal(byref this_object,ByVal useOriginalOrder)
	Dim sql,this,hwhere,eventObj
	doAssignmentByRef sql,this_object.applyWhere_p1(this_object.var_originalSql)
	if bValue(tableEventExists("BeforeQueryReport",strTableName)) then
		doAssignment hwhere,ArrayElement(sql,2)
		doAssignmentByRef eventObj,getEventObject(strTableName)
		eventObj.BeforeQueryReport_p1 hwhere
		setArrElement sql,2,hwhere
	end if
	method_SQLStatement_getOriginal = ((((((((CSmartStr(ArrayElement(sql,0)) & " ") & CSmartStr(IIF((bValue(useOriginalOrder) and bValue(this_object.var_order_in)) and not bValue(this_object.var_oldAlgorithm),(", " & CSmartStr(this_object.var_order_in)) & " ",""))) & CSmartStr(ArrayElement(sql,1))) & " ") & CSmartStr(IIF(ArrayElement(sql,2)," WHERE " & CSmartStr(ArrayElement(sql,2)),""))) & " ") & CSmartStr(ArrayElement(sql,3))) & " ") & CSmartStr(IIF(ArrayElement(sql,4)," HAVING " & CSmartStr(ArrayElement(sql,4)),""))
	Exit Function
End Function
Function method_SQLStatement_setRecordBasedRequest(byref this_object,ByVal recordBasedRequest)
	Dim nCnt
	doClassAssignment this_object,"var_recordBasedRequest",recordBasedRequest
	nCnt = 0
	do while IsLess(nCnt,asp_count(this_object.var_fields))
		doClassAssignment ArrayElement(this_object.var_fields,nCnt),"var_recordBasedRequest",recordBasedRequest
		nCnt = CSmartDbl(nCnt)+1
	loop
End Function
Function method_SQLStatement_getGroup(byref this_object,ByVal data)
	doAssignmentByRef method_SQLStatement_getGroup,ArrayElement(this_object.var_fields,0).getGroup_p1(data)
	Exit Function
End Function
Function method_SQLStatement_field(byref this_object,ByVal num)
	doAssignmentByRef method_SQLStatement_field,ArrayElement(this_object.var_fields,num)
	Exit Function
End Function
Function method_SQLStatement_getSQLLimits(byref this_object,ByVal sql,ByVal from)
	Dim out,nsel
	if IsLessOrEqual(0,from) and IsLessOrEqual(0,this_object.var_pagesize) then
		doAssignmentByRef nsel,asp_stripos(sql,"select",empty)
		doAssignmentByRef out,asp_substr_replace(sql,"select top " & CSmartStr(CSmartDbl(from)+CSmartDbl(this_object.var_pagesize)),nsel,asp_strlen("select"))
		doClassAssignment this_object,"var_skipCount",from
		doAssignmentByRef method_SQLStatement_getSQLLimits,out
		Exit Function
	else
		doAssignmentByRef method_SQLStatement_getSQLLimits,sql
		Exit Function
	end if
End Function
Function method_SQLStatement_sqlg(byref this_object,ByVal donotlimit,ByVal doorder)
	Dim hsql,s,g,o,this
	Set hsql = (CreateDictionary())
	Set s = (CreateDictionary())
	Set g = (CreateDictionary())
	Set o = (CreateDictionary())
	if bValue(this_object.var_hasGroups) then
		setArrElement s,asp_count(s),ArrayElement(this_object.var_fields,0).getSelectSql_p1(true)
		setArrElement g,asp_count(g),ArrayElement(this_object.var_fields,0).getGroupSql()
		setArrElement o,asp_count(o),ArrayElement(this_object.var_fields,0).getOrderSql()
	end if
	if bValue(asp_count(s)) then
		setArrElementN hsql,CreateArray2("select",empty),asp_join(", ",s)
	end if
	if bValue(asp_count(g)) then
		setArrElementN hsql,CreateArray2("groupby",empty),asp_join(", ",g)
	end if
	if bValue(asp_count(o)) and bValue(doorder) then
		setArrElement hsql,"orderby",asp_join(", ",o)
	end if
	if IsEqual(this_object.var_limitLevel,1) and not bValue(donotlimit) then
		setArrElement hsql,"limits",1
	end if
	doAssignmentByRef method_SQLStatement_sqlg,this_object.buildsql_p1(hsql)
	Exit Function
End Function
Function method_SQLStatement_sqlcg(byref this_object)
	Dim gsql,this
	doAssignmentByRef gsql,this_object.sqlg_p2(true,false)
	method_SQLStatement_sqlcg = ((("select count(*) as " & CSmartStr(cached_ffn("c"))) & " from (") & CSmartStr(gsql)) & ") countgroups"
	Exit Function
End Function
Function method_SQLStatement_sqlt(byref this_object)
	Dim hsql,this
	Set hsql = (CreateDictionary())
	setArrElementN hsql,CreateArray2("select",empty),"count(1) as " & CSmartStr(cached_ffn("countField"))
	if bValue(asp_count(this_object.var_aggregates)) then
		setArrElementN hsql,CreateArray2("select",empty),asp_join(", ",this_object.var_aggregates)
	end if
	doAssignmentByRef method_SQLStatement_sqlt,this_object.buildsql_p1(hsql)
	Exit Function
End Function
Function method_SQLStatement_sql2(byref this_object,ByVal groups)
	Dim hsql,o,f,s,g,where
	Set hsql = (CreateDictionary())
	if not bValue(this_object.var_hasGroups) or bValue(this_object.var_recordBasedRequest) then
		setArrElement hsql,"original",true
		Set o = (CreateDictionary())
		GetCollectionBounds this_object.var_fields,c_reportlib_loopIdx17,c_reportlib_loopMax17
		do while c_reportlib_loopIdx17<=c_reportlib_loopMax17
			c_reportlib_arrKey17 = GetCollectionKey(this_object.var_fields,c_reportlib_loopIdx17)
			doAssignment f,ArrayElement(this_object.var_fields,c_reportlib_arrKey17)
			setArrElement o,asp_count(o),f.getOrderSql()
			c_reportlib_loopIdx17=c_reportlib_loopIdx17+1
		loop
		if bValue(asp_count(o)) then
			setArrElement hsql,"orderby",asp_join(", ",o)
		end if
	else
		if bValue(this_object.repShowDet) then
			setArrElementN hsql,CreateArray2("select",empty),"original.*"
		else
			if bValue(asp_count(this_object.var_aggregates)) then
				setArrElementN hsql,CreateArray2("select",empty),asp_join(", ",this_object.var_aggregates)
			end if
		end if
		Set s = (CreateDictionary())
		Set g = (CreateDictionary())
		Set o = (CreateDictionary())
		GetCollectionBounds this_object.var_fields,c_reportlib_loopIdx18,c_reportlib_loopMax18
		do while c_reportlib_loopIdx18<=c_reportlib_loopMax18
			c_reportlib_arrKey18 = GetCollectionKey(this_object.var_fields,c_reportlib_loopIdx18)
			doAssignment f,ArrayElement(this_object.var_fields,c_reportlib_arrKey18)
			setArrElement s,asp_count(s),f.getSelectSql_p1(not bValue(this_object.repShowDet))
			if not bValue(this_object.repShowDet) then
				setArrElement g,asp_count(g),f.getGroupSql()
			end if
			setArrElement o,asp_count(o),f.getOrderSql()
			c_reportlib_loopIdx18=c_reportlib_loopIdx18+1
		loop
		if (bValue(this_object.var_reportSummary) and bValue(this_object.var_hasGroups)) and not bValue(this_object.repShowDet) then
			setArrElementN hsql,CreateArray2("select",empty),"count(1) as " & CSmartStr(cached_ffn("countField"))
		end if
		if bValue(asp_count(s)) then
			setArrElementN hsql,CreateArray2("select",empty),asp_join(", ",s)
		end if
		if not IsNull(groups) and bValue(asp_count(groups)) then
			doAssignmentByRef where,ArrayElement(this_object.var_fields,0).getWhereSql_p1(groups)
			if bValue(where) then
				setArrElement hsql,"where",where
			end if
		end if
		if bValue(asp_count(g)) then
			setArrElement hsql,"groupby",g
		end if
		if bValue(asp_count(o)) then
			setArrElement hsql,"orderby",asp_join(", ",o)
		end if
	end if
	if IsEqual(this_object.var_limitLevel,2) then
		setArrElement hsql,"limits",1
	end if
	if bValue(this_object.repShowDet) then
		setArrElement hsql,"origorder",1
	end if
	doAssignmentByRef method_SQLStatement_sql2,hsql
	Exit Function
End Function
Function method_SQLStatement_buildsql(byref this_object,ByVal hsql)
	Dim ordered,sql,this,osql
	this_object.var_skipCount = 0
	ordered = false
	if IsEqual(asp_count(hsql),0) or bValue(ArrayElement(hsql,"original")) then
		doAssignmentByRef sql,this_object.getOriginal()
	else
		sql = "SELECT "
		if bValue(ArrayElement(hsql,"select")) and IsLess(0,asp_count(ArrayElement(hsql,"select"))) then
			sql = CSmartStr(sql) & CSmartStr(asp_join(", ",ArrayElement(hsql,"select")))
		else
			sql = CSmartStr(sql) & " * "
		end if
		sql = CSmartStr(sql) & ((" FROM (" & CSmartStr(this_object.getOriginal_p1(ArrayElement(hsql,"origorder")))) & ") original")
		if bValue(ArrayElement(hsql,"where")) and IsLess(0,asp_count(ArrayElement(hsql,"where"))) then
			sql = CSmartStr(sql) & (" WHERE " & CSmartStr(ArrayElement(hsql,"where")))
		end if
		if bValue(ArrayElement(hsql,"groupby")) and IsLess(0,asp_count(ArrayElement(hsql,"groupby"))) then
			sql = CSmartStr(sql) & (" GROUP BY " & CSmartStr(asp_join(", ",ArrayElement(hsql,"groupby"))))
		end if
	end if
	osql = ""
	if bValue(ArrayElement(hsql,"orderby")) and IsLess(0,asp_count(ArrayElement(hsql,"orderby"))) then
		osql = CSmartStr(osql) & CSmartStr(ArrayElement(hsql,"orderby"))
		ordered = true
	end if
	if bValue(ArrayElement(hsql,"origorder")) then
		if not bValue(this_object.var_oldAlgorithm) then
			if bValue(this_object.var_order_out) then
				osql = CSmartStr(osql) & (CSmartStr(IIF(osql,", ","")) & CSmartStr(this_object.var_order_out))
			end if
		else
			if bValue(this_object.var_order_old) then
				osql = CSmartStr(osql) & (CSmartStr(IIF(osql,", ","")) & CSmartStr(this_object.var_order_old))
			end if
		end if
	end if
	if bValue(osql) then
		sql = CSmartStr(sql) & (" ORDER BY " & CSmartStr(osql))
	end if
	if bValue(ArrayElement(hsql,"limits")) then
		doAssignmentByRef sql,this_object.getSQLLimits_p2(sql,this_object.var_from)
	end if
	doAssignmentByRef method_SQLStatement_buildsql,sql
	Exit Function
End Function
Function method_SQLStatement_initWhere(byref this_object)
End Function
Function method_SQLStatement_getWhere(byref this_object)
	doAssignmentByRef method_SQLStatement_getWhere,this_object.searchClauseObj
	Exit Function
End Function
Function method_SQLStatement_applyWhere(byref this_object,ByRef sql)
	doAssignmentByRef method_SQLStatement_applyWhere,this_object.searchClauseObj.applyWhere_p3(sql,GetListOfFieldsByExprType(false,this_object.tName),GetListOfFieldsByExprType(true,this_object.tName))
	Exit Function
End Function
Function method_SQLStatement_setOldAlgorithm(byref this_object,ByVal useOldAlgorithm)
	Dim nCnt
	nCnt = 0
	do while IsLess(nCnt,asp_count(this_object.var_fields))
		doClassAssignment ArrayElement(this_object.var_fields,nCnt),"var_oldAlgorithm",useOldAlgorithm
		nCnt = CSmartDbl(nCnt)+1
	loop
	doClassAssignment this_object,"var_oldAlgorithm",useOldAlgorithm
End Function

'------ Class Summarable ------
Class Summarable
	Public var_summary
	Public tName
	Public shortTName
	Public repGroupFieldsCount
	Public repPageSummary
	Public repGlobalSummary
	Public repLayout
	Public showGroupSummaryCount
	Public repShowDet
	Public repGroupFields
	Public tKeyFields
	Public isExistTotalFields
	Public fieldsArr
	Public Function init_Summarable_p1(ByRef params)
		DoAssignmentByRef init_Summarable_p1,method_Summarable_Summarable(me,params)
	End Function
	Public Function init_p1(ByVal from)
		DoAssignmentByRef init_p1,method_Summarable_init(me,from)
	End Function
	Public Function init()
		DoAssignmentByRef init,method_Summarable_init(me,0)
	End Function
	Public Function writeGroup_p5(ByRef begin,ByRef var_end,ByVal gkey,ByVal grp,ByVal nField)
		DoAssignmentByRef writeGroup_p5,method_Summarable_writeGroup(me,begin,var_end,gkey,grp,nField)
	End Function
	Public Function addSummary_p4(ByVal recordsMode,ByRef summary,ByVal data,ByRef nTotalRecords)
		DoAssignmentByRef addSummary_p4,method_Summarable_addSummary(me,recordsMode,summary,data,nTotalRecords)
	End Function
	Public Function func_makeSummary_p2(ByRef summary,ByVal deep)
		DoAssignmentByRef func_makeSummary_p2,method_Summarable__makeSummary(me,summary,deep)
	End Function
	Public Function value2time_p1(ByVal value)
		DoAssignmentByRef value2time_p1,method_Summarable_value2time(me,value)
	End Function
	Public Function time2printable_p1(ByVal time)
		DoAssignmentByRef time2printable_p1,method_Summarable_time2printable(me,time)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"var_summary", var_summary
		setArrElement out,"tName", tName
		setArrElement out,"shortTName", shortTName
		setArrElement out,"repGroupFieldsCount", repGroupFieldsCount
		setArrElement out,"repPageSummary", repPageSummary
		setArrElement out,"repGlobalSummary", repGlobalSummary
		setArrElement out,"repLayout", repLayout
		setArrElement out,"showGroupSummaryCount", showGroupSummaryCount
		setArrElement out,"repShowDet", repShowDet
		setArrElement out,"repGroupFields", repGroupFields
		setArrElement out,"tKeyFields", tKeyFields
		setArrElement out,"isExistTotalFields", isExistTotalFields
		setArrElement out,"fieldsArr", fieldsArr
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment var_summary, dict("var_summary")
		DoAssignment tName, dict("tName")
		DoAssignment shortTName, dict("shortTName")
		DoAssignment repGroupFieldsCount, dict("repGroupFieldsCount")
		DoAssignment repPageSummary, dict("repPageSummary")
		DoAssignment repGlobalSummary, dict("repGlobalSummary")
		DoAssignment repLayout, dict("repLayout")
		DoAssignment showGroupSummaryCount, dict("showGroupSummaryCount")
		DoAssignment repShowDet, dict("repShowDet")
		DoAssignment repGroupFields, dict("repGroupFields")
		DoAssignment tKeyFields, dict("tKeyFields")
		DoAssignment isExistTotalFields, dict("isExistTotalFields")
		DoAssignment fieldsArr, dict("fieldsArr")
	End Sub
' end serialize
End Class
' Summarable implementation 
Function method_Summarable_Summarable(byref this_object,ByRef params)
	doClassAssignment this_object,"var_summary",CreateDictionary()
	this_object.tName = ""
	this_object.shortTName = ""
	this_object.repGroupFieldsCount = 0
	this_object.repPageSummary = 0
	this_object.repGlobalSummary = 0
	this_object.repLayout = 0
	this_object.showGroupSummaryCount = 0
	this_object.repShowDet = 0
	doClassAssignment this_object,"repGroupFields",CreateDictionary()
	doClassAssignment this_object,"tKeyFields",CreateDictionary()
	this_object.isExistTotalFields = false
	doClassAssignment this_object,"fieldsArr",CreateDictionary()
	Dim this
	RunnerApply this_object,params
	method_Summarable_init this_object,0
End Function
Function method_Summarable_init(byref this_object,ByVal from)
	doClassAssignment this_object,"var_from",from
End Function
Function method_Summarable_writeGroup(byref this_object,ByRef begin,ByRef var_end,ByVal gkey,ByVal grp,ByVal nField)
End Function
Function method_Summarable_addSummary(byref this_object,ByVal recordsMode,ByRef summary,ByVal data,ByRef nTotalRecords)
	Dim countInGroup,s,i,avg_value,this
	doAssignment countInGroup,IIF(not IsEmpty(ArrayElement(summary,"count")),ArrayElement(summary,"count"),0)
	if bValue(this_object.isExistTotalFields) then
		if not bValue(asp_is_array(ArrayElement(summary,"summary"))) then
			setArrElement summary,"summary",CreateDictionary()
		end if
		doAssignmentByRef s,ArrayElement(summary,"summary")
	end if
	if bValue(recordsMode) then
		i = 0
		c_reportlib_exitLoop20=false
		do while IsLess(i,asp_count(this_object.fieldsArr))
			c_reportlib_exitLoop20=false
			do
				if ((not bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalMax")) and not bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalMin"))) and not bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalAvg"))) and not bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalSum")) then
					exit do
				end if
				if not IsNull(ArrayElement(data,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"))) then
					if not bValue(asp_is_array(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")))) then
						setArrElement s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),CreateDictionary()
					end if
					if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalMax")) then
						if not (not IsEmpty(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MAX"))) or IsLess(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MAX"),ArrayElement(data,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"))) then
							setArrElementN s,CreateArray2(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"MAX"),ArrayElement(data,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"))
						end if
					end if
					if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalMin")) then
						if not (not IsEmpty(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MIN"))) or IsLess(ArrayElement(data,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MIN")) then
							setArrElementN s,CreateArray2(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"MIN"),ArrayElement(data,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"))
						end if
					end if
					if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalAvg")) then
						if IsEqual(ArrayElement(ArrayElement(this_object.fieldsArr,i),"viewFormat"),"Time") then
							doAssignmentByRef avg_value,this_object.value2time_p1(ArrayElement(data,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")))
						else
							doAssignment avg_value,ArrayElement(data,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"))
						end if
						setArrElementN s,CreateArray2(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"AVG"),CSmartDbl(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"AVG"))*CSmartDbl(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"count"))+CSmartDbl(avg_value)
						setArrElementN s,CreateArray2(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"count"),CSmartDbl(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"count"))+1
						setArrElementN s,CreateArray2(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"AVG"),CSmartDbl(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"AVG"))/CSmartDbl(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"count"))
					end if
					if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalSum")) then
						if IsEqual(ArrayElement(ArrayElement(this_object.fieldsArr,i),"viewFormat"),"Time") then
							setArrElementN s,CreateArray2(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"SUM"),CSmartDbl(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"SUM"))+CSmartDbl(this_object.value2time_p1(ArrayElement(data,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"))))
						else
							setArrElementN s,CreateArray2(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"SUM"),CSmartDbl(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"SUM"))+CSmartDbl(ArrayElement(data,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")))
						end if
					end if
				end if
			loop while false
			if c_reportlib_exitLoop20 then _
				exit do
			i = CSmartDbl(i)+1
		loop
		nTotalRecords = CSmartDbl(nTotalRecords)+1
		countInGroup = CSmartDbl(countInGroup)+1
	else
		i = 0
		c_reportlib_exitLoop21=false
		do while IsLess(i,asp_count(this_object.fieldsArr))
			c_reportlib_exitLoop21=false
			do
				if ((not bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalMax")) and not bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalMin"))) and not bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalAvg"))) and not bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalSum")) then
					exit do
				end if
				if not bValue(asp_is_array(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")))) then
					setArrElement s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),CreateDictionary()
				end if
				if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalMax")) then
					if not IsNull(ArrayElement(data,CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) & "MAX")) then
						if not (not IsEmpty(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MAX"))) or IsLess(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MAX"),ArrayElement(data,CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) & "MAX")) then
							setArrElementN s,CreateArray2(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"MAX"),ArrayElement(data,CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) & "MAX")
						end if
					end if
				end if
				if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalMin")) then
					if not IsNull(ArrayElement(data,CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) & "MIN")) then
						if not (not IsEmpty(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MIN"))) or IsLess(ArrayElement(data,CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) & "MIN"),ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MIN")) then
							setArrElementN s,CreateArray2(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"MIN"),ArrayElement(data,CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) & "MIN")
						end if
					end if
				end if
				if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalAvg")) then
					if not IsNull(ArrayElement(data,CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) & "AVG")) then
						if IsEqual(ArrayElement(ArrayElement(this_object.fieldsArr,i),"viewFormat"),"Time") then
							doAssignmentByRef avg_value,this_object.value2time_p1(ArrayElement(data,CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) & "AVG"))
						else
							doAssignment avg_value,ArrayElement(data,CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) & "AVG")
						end if
						setArrElementN s,CreateArray2(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"AVG"),CSmartDbl(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"AVG"))*CSmartDbl(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"count"))+CSmartDbl(avg_value)*CSmartDbl(ArrayElement(data,CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) & "NAVG"))
						setArrElementN s,CreateArray2(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"count"),CSmartDbl(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"count"))+CSmartDbl(ArrayElement(data,CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) & "NAVG"))
						setArrElementN s,CreateArray2(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"AVG"),CSmartDbl(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"AVG"))/CSmartDbl(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"count"))
					end if
				end if
				if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalSum")) then
					if not IsNull(ArrayElement(data,CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) & "SUM")) then
						if IsEqual(ArrayElement(ArrayElement(this_object.fieldsArr,i),"viewFormat"),"Time") then
							setArrElementN s,CreateArray2(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"SUM"),CSmartDbl(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"SUM"))+CSmartDbl(this_object.value2time_p1(ArrayElement(data,CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) & "SUM")))
						else
							setArrElementN s,CreateArray2(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"SUM"),CSmartDbl(ArrayElement(ArrayElement(s,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"SUM"))+CSmartDbl(ArrayElement(data,CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) & "SUM"))
						end if
					end if
				end if
			loop while false
			if c_reportlib_exitLoop21 then _
				exit do
			i = CSmartDbl(i)+1
		loop
		nTotalRecords = CSmartDbl(nTotalRecords)+CSmartDbl(ArrayElement(data,"countField"))
		countInGroup = CSmartDbl(countInGroup)+CSmartDbl(ArrayElement(data,"countField"))
	end if
	setArrElement summary,"count",countInGroup
End Function
Function method_Summarable__makeSummary(byref this_object,ByRef summary,ByVal deep)
	Dim grp,gkey,this,i
	if bValue(ArrayElement(summary,"values")) then
		GetCollectionBounds ArrayElement(summary,"values"),c_reportlib_loopIdx22,c_reportlib_loopMax22
		do while c_reportlib_loopIdx22<=c_reportlib_loopMax22
			gkey = GetCollectionKey(ArrayElement(summary,"values"),c_reportlib_loopIdx22)
			doAssignment group,ArrayElement(ArrayElement(summary,"values"),gkey)
			doAssignmentByRef grp,ArrayElement(ArrayElement(summary,"values"),gkey)
			if not IsEmpty(ArrayElement(grp,"values")) then
				this_object.func_makeSummary_p2 grp,CSmartDbl(deep)+1
			end if
			if not IsEmpty(ArrayElement(grp,"_begin")) and not IsEmpty(ArrayElement(grp,"_end")) then
				this_object.writeGroup_p5 ArrayElement(grp,"_begin"),ArrayElement(grp,"_end"),gkey,grp,deep
			end if
			if not bValue(asp_is_array(ArrayElement(summary,"summary"))) then
				setArrElement summary,"summary",CreateDictionary()
			end if
			i = 0
			do while IsLess(i,asp_count(this_object.fieldsArr))
				if not bValue(asp_is_array(ArrayElement(ArrayElement(summary,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")))) then
					setArrElementN summary,CreateArray2("summary",ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),CreateDictionary()
				end if
				if bValue(asp_is_array(ArrayElement(grp,"summary"))) then
					if bValue(asp_is_array(ArrayElement(ArrayElement(grp,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")))) then
						if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalMax")) then
							if not IsEmpty(ArrayElement(ArrayElement(ArrayElement(grp,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MAX")) then
								if not (not IsEmpty(ArrayElement(ArrayElement(ArrayElement(summary,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MAX"))) or IsLess(ArrayElement(ArrayElement(ArrayElement(summary,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MAX"),ArrayElement(ArrayElement(ArrayElement(grp,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MAX")) then
									setArrElementN summary,CreateArray3("summary",ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"MAX"),ArrayElement(ArrayElement(ArrayElement(grp,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MAX")
								end if
							end if
						end if
						if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalMin")) then
							if not IsEmpty(ArrayElement(ArrayElement(ArrayElement(grp,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MIN")) then
								if not (not IsEmpty(ArrayElement(ArrayElement(ArrayElement(summary,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MIN"))) or IsLess(ArrayElement(ArrayElement(ArrayElement(grp,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MIN"),ArrayElement(ArrayElement(ArrayElement(summary,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MIN")) then
									setArrElementN summary,CreateArray3("summary",ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"MIN"),ArrayElement(ArrayElement(ArrayElement(grp,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MIN")
								end if
							end if
						end if
						if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalAvg")) then
							if not IsEmpty(ArrayElement(ArrayElement(ArrayElement(grp,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"AVG")) then
								setArrElementN summary,CreateArray3("summary",ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"AVG"),CSmartDbl(ArrayElement(ArrayElement(ArrayElement(summary,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"AVG"))*CSmartDbl(ArrayElement(ArrayElement(ArrayElement(summary,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"count"))+CSmartDbl(ArrayElement(ArrayElement(ArrayElement(grp,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"AVG"))*CSmartDbl(ArrayElement(ArrayElement(ArrayElement(grp,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"count"))
								setArrElementN summary,CreateArray3("summary",ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"count"),CSmartDbl(ArrayElement(ArrayElement(ArrayElement(summary,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"count"))+CSmartDbl(ArrayElement(ArrayElement(ArrayElement(grp,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"count"))
								setArrElementN summary,CreateArray3("summary",ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"AVG"),CSmartDbl(ArrayElement(ArrayElement(ArrayElement(summary,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"AVG"))/CSmartDbl(ArrayElement(ArrayElement(ArrayElement(summary,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"count"))
							end if
						end if
						if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalSum")) then
							if bValue(ArrayElement(ArrayElement(ArrayElement(grp,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"SUM")) then
								setArrElementN summary,CreateArray3("summary",ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),"SUM"),CSmartDbl(ArrayElement(ArrayElement(ArrayElement(summary,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"SUM"))+CSmartDbl(ArrayElement(ArrayElement(ArrayElement(grp,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"SUM"))
							end if
						end if
					end if
				end if
				i = CSmartDbl(i)+1
			loop
			setArrElement summary,"count",CSmartDbl(ArrayElement(summary,"count"))+CSmartDbl(ArrayElement(grp,"count"))
			c_reportlib_loopIdx22=c_reportlib_loopIdx22+1
		loop
	end if
End Function
Function method_Summarable_value2time(byref this_object,ByVal value)
	Dim res,arr
	res = 0
	doAssignmentByRef arr,parsenumbers(value)
	if not IsEmpty(ArrayElement(arr,0)) then
		res = CSmartDbl(res)+(CSmartDbl(ArrayElement(arr,0))*60)*60
	end if
	if not IsEmpty(ArrayElement(arr,1)) then
		res = CSmartDbl(res)+CSmartDbl(ArrayElement(arr,1))*60
	end if
	if not IsEmpty(ArrayElement(arr,2)) then
		res = CSmartDbl(res)+CSmartDbl(ArrayElement(arr,2))
	end if
	doAssignmentByRef method_Summarable_value2time,res
	Exit Function
End Function
Function method_Summarable_time2printable(byref this_object,ByVal time)
	doAssignmentByRef method_Summarable_time2printable,CreateDictionary3(Empty,asp_intval(CSmartDbl(time)/(60*60)),Empty,asp_intval(CSmartDbl(time)/60),Empty,time mod 60)
	Exit Function
End Function

'------ Class ReportGroups extends Summarable ------
Class ReportGroups
	Public var_global
	Public var_pagesize
	Public var_totalRecords
	Public var_maxpages
	Public var_nGroup
	Public var_oldFirst
	Public var_from
	Public var_sql
	Public var_connection
	Public var_allGroupsUsed
	Public var_countGroups
	Public var_summary
	Public tName
	Public shortTName
	Public repGroupFieldsCount
	Public repPageSummary
	Public repGlobalSummary
	Public repLayout
	Public showGroupSummaryCount
	Public repShowDet
	Public repGroupFields
	Public tKeyFields
	Public isExistTotalFields
	Public fieldsArr
	Public Function init_ReportGroups_p4(ByRef sql,ByVal connection,ByVal pagesize,ByRef params)
		DoAssignmentByRef init_ReportGroups_p4,method_ReportGroups_ReportGroups(me,sql,connection,pagesize,params)
	End Function
	Public Function init_p1(ByVal from)
		DoAssignmentByRef init_p1,method_ReportGroups_init(me,from)
	End Function
	Public Function init()
		DoAssignmentByRef init,method_ReportGroups_init(me,0)
	End Function
	Public Function setGlobalSummary_p2(ByVal recordsMode,ByVal data)
		DoAssignmentByRef setGlobalSummary_p2,method_ReportGroups_setGlobalSummary(me,recordsMode,data)
	End Function
	Public Function setGroup_p1(ByVal data)
		DoAssignmentByRef setGroup_p1,method_ReportGroups_setGroup(me,data)
	End Function
	Public Function isVisibleGroup()
		DoAssignmentByRef isVisibleGroup,method_ReportGroups_isVisibleGroup(me)
	End Function
	Public Function getDisplayGroups_p1(ByVal from)
		DoAssignmentByRef getDisplayGroups_p1,method_ReportGroups_getDisplayGroups(me,from)
	End Function
	Public Function getCountGroups_p1(ByVal fullRequest)
		DoAssignmentByRef getCountGroups_p1,method_ReportGroups_getCountGroups(me,fullRequest)
	End Function
	Public Function getCountGroups()
		DoAssignmentByRef getCountGroups,method_ReportGroups_getCountGroups(me,false)
	End Function
	Public Function getSummary()
		DoAssignmentByRef getSummary,method_ReportGroups_getSummary(me)
	End Function
	Public Function allGroupsUsed()
		DoAssignmentByRef allGroupsUsed,method_ReportGroups_allGroupsUsed(me)
	End Function
	Public Function writeGroup_p5(ByRef begin,ByRef var_end,ByVal gkey,ByVal grp,ByVal nField)
		DoAssignmentByRef writeGroup_p5,method_Summarable_writeGroup(me,begin,var_end,gkey,grp,nField)
	End Function
	Public Function addSummary_p4(ByVal recordsMode,ByRef summary,ByVal data,ByRef nTotalRecords)
		DoAssignmentByRef addSummary_p4,method_Summarable_addSummary(me,recordsMode,summary,data,nTotalRecords)
	End Function
	Public Function func_makeSummary_p2(ByRef summary,ByVal deep)
		DoAssignmentByRef func_makeSummary_p2,method_Summarable__makeSummary(me,summary,deep)
	End Function
	Public Function value2time_p1(ByVal value)
		DoAssignmentByRef value2time_p1,method_Summarable_value2time(me,value)
	End Function
	Public Function time2printable_p1(ByVal time)
		DoAssignmentByRef time2printable_p1,method_Summarable_time2printable(me,time)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"var_global", var_global
		setArrElement out,"var_pagesize", var_pagesize
		setArrElement out,"var_totalRecords", var_totalRecords
		setArrElement out,"var_maxpages", var_maxpages
		setArrElement out,"var_nGroup", var_nGroup
		setArrElement out,"var_oldFirst", var_oldFirst
		setArrElement out,"var_from", var_from
		setArrElement out,"var_sql", var_sql
		setArrElement out,"var_connection", var_connection
		setArrElement out,"var_allGroupsUsed", var_allGroupsUsed
		setArrElement out,"var_countGroups", var_countGroups
		setArrElement out,"var_summary", var_summary
		setArrElement out,"tName", tName
		setArrElement out,"shortTName", shortTName
		setArrElement out,"repGroupFieldsCount", repGroupFieldsCount
		setArrElement out,"repPageSummary", repPageSummary
		setArrElement out,"repGlobalSummary", repGlobalSummary
		setArrElement out,"repLayout", repLayout
		setArrElement out,"showGroupSummaryCount", showGroupSummaryCount
		setArrElement out,"repShowDet", repShowDet
		setArrElement out,"repGroupFields", repGroupFields
		setArrElement out,"tKeyFields", tKeyFields
		setArrElement out,"isExistTotalFields", isExistTotalFields
		setArrElement out,"fieldsArr", fieldsArr
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment var_global, dict("var_global")
		DoAssignment var_pagesize, dict("var_pagesize")
		DoAssignment var_totalRecords, dict("var_totalRecords")
		DoAssignment var_maxpages, dict("var_maxpages")
		DoAssignment var_nGroup, dict("var_nGroup")
		DoAssignment var_oldFirst, dict("var_oldFirst")
		DoAssignment var_from, dict("var_from")
		DoAssignment var_sql, dict("var_sql")
		DoAssignment var_connection, dict("var_connection")
		DoAssignment var_allGroupsUsed, dict("var_allGroupsUsed")
		DoAssignment var_countGroups, dict("var_countGroups")
		DoAssignment var_summary, dict("var_summary")
		DoAssignment tName, dict("tName")
		DoAssignment shortTName, dict("shortTName")
		DoAssignment repGroupFieldsCount, dict("repGroupFieldsCount")
		DoAssignment repPageSummary, dict("repPageSummary")
		DoAssignment repGlobalSummary, dict("repGlobalSummary")
		DoAssignment repLayout, dict("repLayout")
		DoAssignment showGroupSummaryCount, dict("showGroupSummaryCount")
		DoAssignment repShowDet, dict("repShowDet")
		DoAssignment repGroupFields, dict("repGroupFields")
		DoAssignment tKeyFields, dict("tKeyFields")
		DoAssignment isExistTotalFields, dict("isExistTotalFields")
		DoAssignment fieldsArr, dict("fieldsArr")
	End Sub
' end serialize
End Class
' ReportGroups implementation 
Function method_ReportGroups_ReportGroups(byref this_object,ByRef sql,ByVal connection,ByVal pagesize,ByRef params)
	doClassAssignment this_object,"var_summary",CreateDictionary()
	this_object.tName = ""
	this_object.shortTName = ""
	this_object.repGroupFieldsCount = 0
	this_object.repPageSummary = 0
	this_object.repGlobalSummary = 0
	this_object.repLayout = 0
	this_object.showGroupSummaryCount = 0
	this_object.repShowDet = 0
	doClassAssignment this_object,"repGroupFields",CreateDictionary()
	doClassAssignment this_object,"tKeyFields",CreateDictionary()
	this_object.isExistTotalFields = false
	doClassAssignment this_object,"fieldsArr",CreateDictionary()
	Dim this
	RunnerApply this_object,params
	method_Summarable_Summarable this_object,params
	this_object.init 
	doClassAssignment this_object,"var_pagesize",pagesize
	doClassAssignmentByRef this_object,"var_sql",sql
	doClassAssignment this_object,"var_connection",connection
End Function
Function method_ReportGroups_init(byref this_object,ByVal from)
	method_Summarable_init this_object,from
	doClassAssignment this_object,"var_global",CreateDictionary()
	this_object.var_totalRecords = 0
	this_object.var_maxpages = -1
	doClassAssignment this_object,"var_from",from
	this_object.var_nGroup = -1
	this_object.var_oldFirst = ""
	this_object.var_allGroupsUsed = false
	this_object.var_countGroups = 0
End Function
Function method_ReportGroups_setGlobalSummary(byref this_object,ByVal recordsMode,ByVal data)
	Dim this
	this_object.addSummary_p4 recordsMode,this_object.var_global,data,this_object.var_totalRecords
End Function
Function method_ReportGroups_setGroup(byref this_object,ByVal data)
	Dim field,firstKey
	doAssignmentByRef field,this_object.var_sql.field_p1(0)
	doAssignmentByRef firstKey,field.getKey_p1(data)
	if not IsEqual(firstKey,this_object.var_oldFirst) then
		this_object.var_nGroup = CSmartDbl(this_object.var_nGroup)+1
		doClassAssignment this_object,"var_oldFirst",firstKey
	end if
End Function
Function method_ReportGroups_isVisibleGroup(byref this_object)
	method_ReportGroups_isVisibleGroup = IsLessOrEqual(this_object.var_from,this_object.var_nGroup) and IsLess(this_object.var_nGroup,CSmartDbl(this_object.var_from)+CSmartDbl(this_object.var_pagesize))
	Exit Function
End Function
Function method_ReportGroups_getDisplayGroups(byref this_object,ByVal from)
	Dim this,groups,sql,cursor,data
	this_object.init_p1 from
	if IsEqual(this_object.var_pagesize,-1) then
		doAssignmentByRef method_ReportGroups_getDisplayGroups,CreateDictionary()
		Exit Function
	else
		Set groups = (CreateDictionary())
		this_object.var_allGroupsUsed = false
		if bValue(this_object.repGroupFieldsCount) then
			doAssignmentByRef sql,this_object.var_sql.sqlg()
			doAssignmentByRef cursor,db_query(sql,this_object.var_connection)
			do while bValue(doAssignmentByRef(data,db_fetch_array(cursor)))
				setArrElement groups,asp_count(groups),this_object.var_sql.getGroup_p1(data)
			loop
			if IsLess(asp_count(groups),this_object.var_pagesize) then
				this_object.var_allGroupsUsed = true
			end if
		end if
		if IsLess(0,this_object.var_sql.var_skipCount) then
			asp_array_splice groups,0,this_object.var_sql.var_skipCount
			this_object.var_allGroupsUsed = false
		end if
		if IsLess(0,from) then
			this_object.var_allGroupsUsed = false
		end if
		doClassAssignment this_object,"var_countGroups",asp_count(groups)
		doAssignmentByRef method_ReportGroups_getDisplayGroups,groups
		Exit Function
	end if
End Function
Function method_ReportGroups_getCountGroups(byref this_object,ByVal fullRequest)
	Dim sql,cursor,data
	if bValue(this_object.repGroupFieldsCount) then
		if IsLessOrEqual(0,this_object.var_nGroup) and bValue(fullRequest) then
			method_ReportGroups_getCountGroups = CSmartDbl(this_object.var_nGroup)+1
			Exit Function
		else
			if bValue(this_object.var_allGroupsUsed) then
				doAssignmentByRef method_ReportGroups_getCountGroups,this_object.var_countGroups
				Exit Function
			else
				doAssignmentByRef sql,this_object.var_sql.sqlcg()
				doAssignmentByRef cursor,db_query(sql,this_object.var_connection)
				doAssignmentByRef data,db_fetch_array(cursor)
				doAssignmentByRef method_ReportGroups_getCountGroups,ArrayElement(data,"c")
				Exit Function
			end if
		end if
	else
		method_ReportGroups_getCountGroups = 0
		Exit Function
	end if
End Function
Function method_ReportGroups_getSummary(byref this_object)
	doAssignmentByRef method_ReportGroups_getSummary,this_object.var_global
	Exit Function
End Function
Function method_ReportGroups_allGroupsUsed(byref this_object)
	doAssignmentByRef method_ReportGroups_allGroupsUsed,this_object.var_allGroupsUsed
	Exit Function
End Function

'------ Class ReportLogic extends Summarable ------
Class ReportLogic
	Public var_list
	Public var_totalRecords
	Public var_pages
	Public var_pagesize
	Public var_printpagesize
	Public var_from
	Public var_connection
	Public var_sql
	Public var_groups
	Public var_groupKeys
	Public var_fullRequest
	Public var_recordBasedRequest
	Public var_doPaging
	Public var_lastPageNumber
	Public var_pageSummary
	Public var_printRecordCount
	Public var_listedRows
	Public var_oldLevels
	Public searchClauseObj
	Public var_summary
	Public tName
	Public shortTName
	Public repGroupFieldsCount
	Public repPageSummary
	Public repGlobalSummary
	Public repLayout
	Public showGroupSummaryCount
	Public repShowDet
	Public repGroupFields
	Public tKeyFields
	Public isExistTotalFields
	Public fieldsArr
	Public Function init_ReportLogic_p7(ByVal sql,ByVal order,ByRef searchClauseObj,ByVal connection,ByVal pagesize,ByVal printpagesize,ByRef params)
		DoAssignmentByRef init_ReportLogic_p7,method_ReportLogic_ReportLogic(me,sql,order,searchClauseObj,connection,pagesize,printpagesize,params)
	End Function
	Public Function init_p1(ByVal from)
		DoAssignmentByRef init_p1,method_ReportLogic_init(me,from)
	End Function
	Public Function init()
		DoAssignmentByRef init,method_ReportLogic_init(me,0)
	End Function
	Public Function getWhere()
		DoAssignmentByRef getWhere,method_ReportLogic_getWhere(me)
	End Function
	Public Function getPages()
		DoAssignmentByRef getPages,method_ReportLogic_getPages(me)
	End Function
	Public Function getFormattedRow_p1(ByVal value)
		DoAssignmentByRef getFormattedRow_p1,method_ReportLogic_getFormattedRow(me,value)
	End Function
	Public Function writeGroup_p5(ByRef begin,ByRef var_end,ByVal gkey,ByVal grp,ByVal nField)
		DoAssignmentByRef writeGroup_p5,method_ReportLogic_writeGroup(me,begin,var_end,gkey,grp,nField)
	End Function
	Public Function func_writePage_p3(ByRef page,ByVal src,ByVal count)
		DoAssignmentByRef func_writePage_p3,method_ReportLogic__writePage(me,page,src,count)
	End Function
	Public Function writeGlobalSummary_p1(ByVal source)
		DoAssignmentByRef writeGlobalSummary_p1,method_ReportLogic_writeGlobalSummary(me,source)
	End Function
	Public Function writePageSummary()
		DoAssignmentByRef writePageSummary,method_ReportLogic_writePageSummary(me)
	End Function
	Public Function makeSummary()
		DoAssignmentByRef makeSummary,method_ReportLogic_makeSummary(me)
	End Function
	Public Function setSummary_p3(ByVal recordsMode,ByVal data,ByVal rowToAppend)
		DoAssignmentByRef setSummary_p3,method_ReportLogic_setSummary(me,recordsMode,data,rowToAppend)
	End Function
	Public Function setSummary_p2(ByVal recordsMode,ByVal data)
		DoAssignmentByRef setSummary_p2,method_ReportLogic_setSummary(me,recordsMode,data,null)
	End Function
	Public Function setFinish()
		DoAssignmentByRef setFinish,method_ReportLogic_setFinish(me)
	End Function
	Public Function appendRow_p1(ByVal row)
		DoAssignmentByRef appendRow_p1,method_ReportLogic_appendRow(me,row)
	End Function
	Public Function recordVisible_p1(ByVal nRecord)
		DoAssignmentByRef recordVisible_p1,method_ReportLogic_recordVisible(me,nRecord)
	End Function
	Public Function getTotals()
		DoAssignmentByRef getTotals,method_ReportLogic_getTotals(me)
	End Function
	Public Function getReport_p1(ByVal from)
		DoAssignmentByRef getReport_p1,method_ReportLogic_getReport(me,from)
	End Function
	Public Function getReport()
		DoAssignmentByRef getReport,method_ReportLogic_getReport(me,0)
	End Function
	Public Function addSummary_p4(ByVal recordsMode,ByRef summary,ByVal data,ByRef nTotalRecords)
		DoAssignmentByRef addSummary_p4,method_Summarable_addSummary(me,recordsMode,summary,data,nTotalRecords)
	End Function
	Public Function func_makeSummary_p2(ByRef summary,ByVal deep)
		DoAssignmentByRef func_makeSummary_p2,method_Summarable__makeSummary(me,summary,deep)
	End Function
	Public Function value2time_p1(ByVal value)
		DoAssignmentByRef value2time_p1,method_Summarable_value2time(me,value)
	End Function
	Public Function time2printable_p1(ByVal time)
		DoAssignmentByRef time2printable_p1,method_Summarable_time2printable(me,time)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"var_list", var_list
		setArrElement out,"var_totalRecords", var_totalRecords
		setArrElement out,"var_pages", var_pages
		setArrElement out,"var_pagesize", var_pagesize
		setArrElement out,"var_printpagesize", var_printpagesize
		setArrElement out,"var_from", var_from
		setArrElement out,"var_connection", var_connection
		setArrElement out,"var_sql", var_sql
		setArrElement out,"var_groups", var_groups
		setArrElement out,"var_groupKeys", var_groupKeys
		setArrElement out,"var_fullRequest", var_fullRequest
		setArrElement out,"var_recordBasedRequest", var_recordBasedRequest
		setArrElement out,"var_doPaging", var_doPaging
		setArrElement out,"var_lastPageNumber", var_lastPageNumber
		setArrElement out,"var_pageSummary", var_pageSummary
		setArrElement out,"var_printRecordCount", var_printRecordCount
		setArrElement out,"var_listedRows", var_listedRows
		setArrElement out,"var_oldLevels", var_oldLevels
		setArrElement out,"searchClauseObj", searchClauseObj
		setArrElement out,"var_summary", var_summary
		setArrElement out,"tName", tName
		setArrElement out,"shortTName", shortTName
		setArrElement out,"repGroupFieldsCount", repGroupFieldsCount
		setArrElement out,"repPageSummary", repPageSummary
		setArrElement out,"repGlobalSummary", repGlobalSummary
		setArrElement out,"repLayout", repLayout
		setArrElement out,"showGroupSummaryCount", showGroupSummaryCount
		setArrElement out,"repShowDet", repShowDet
		setArrElement out,"repGroupFields", repGroupFields
		setArrElement out,"tKeyFields", tKeyFields
		setArrElement out,"isExistTotalFields", isExistTotalFields
		setArrElement out,"fieldsArr", fieldsArr
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment var_list, dict("var_list")
		DoAssignment var_totalRecords, dict("var_totalRecords")
		DoAssignment var_pages, dict("var_pages")
		DoAssignment var_pagesize, dict("var_pagesize")
		DoAssignment var_printpagesize, dict("var_printpagesize")
		DoAssignment var_from, dict("var_from")
		DoAssignment var_connection, dict("var_connection")
		DoAssignment var_sql, dict("var_sql")
		DoAssignment var_groups, dict("var_groups")
		DoAssignment var_groupKeys, dict("var_groupKeys")
		DoAssignment var_fullRequest, dict("var_fullRequest")
		DoAssignment var_recordBasedRequest, dict("var_recordBasedRequest")
		DoAssignment var_doPaging, dict("var_doPaging")
		DoAssignment var_lastPageNumber, dict("var_lastPageNumber")
		DoAssignment var_pageSummary, dict("var_pageSummary")
		DoAssignment var_printRecordCount, dict("var_printRecordCount")
		DoAssignment var_listedRows, dict("var_listedRows")
		DoAssignment var_oldLevels, dict("var_oldLevels")
		DoAssignment searchClauseObj, dict("searchClauseObj")
		DoAssignment var_summary, dict("var_summary")
		DoAssignment tName, dict("tName")
		DoAssignment shortTName, dict("shortTName")
		DoAssignment repGroupFieldsCount, dict("repGroupFieldsCount")
		DoAssignment repPageSummary, dict("repPageSummary")
		DoAssignment repGlobalSummary, dict("repGlobalSummary")
		DoAssignment repLayout, dict("repLayout")
		DoAssignment showGroupSummaryCount, dict("showGroupSummaryCount")
		DoAssignment repShowDet, dict("repShowDet")
		DoAssignment repGroupFields, dict("repGroupFields")
		DoAssignment tKeyFields, dict("tKeyFields")
		DoAssignment isExistTotalFields, dict("isExistTotalFields")
		DoAssignment fieldsArr, dict("fieldsArr")
	End Sub
' end serialize
End Class
' ReportLogic implementation 
Function method_ReportLogic_ReportLogic(byref this_object,ByVal sql,ByVal order,ByRef searchClauseObj,ByVal connection,ByVal pagesize,ByVal printpagesize,ByRef params)
	this_object.var_from = 0
	this_object.var_fullRequest = false
	this_object.var_recordBasedRequest = false
	this_object.var_doPaging = false
	this_object.var_lastPageNumber = 0
	this_object.var_printRecordCount = 0
	this_object.var_listedRows = 0
	this_object.searchClauseObj = null
	doClassAssignment this_object,"var_summary",CreateDictionary()
	this_object.tName = ""
	this_object.shortTName = ""
	this_object.repGroupFieldsCount = 0
	this_object.repPageSummary = 0
	this_object.repGlobalSummary = 0
	this_object.repLayout = 0
	this_object.showGroupSummaryCount = 0
	this_object.repShowDet = 0
	doClassAssignment this_object,"repGroupFields",CreateDictionary()
	doClassAssignment this_object,"tKeyFields",CreateDictionary()
	this_object.isExistTotalFields = false
	doClassAssignment this_object,"fieldsArr",CreateDictionary()
	Dim this
	RunnerApply this_object,params
	method_Summarable_Summarable this_object,params
	doClassAssignment this_object,"var_connection",connection
	doClassAssignment this_object,"searchClauseObj",searchClauseObj
	doClassAssignment this_object,"var_sql",CreateClass("SQLStatement",6,sql,order,pagesize,connection,searchClauseObj,params,Empty)
	doClassAssignment this_object,"var_groups",CreateClass("ReportGroups",4,this_object.var_sql,connection,pagesize,params,Empty,Empty,Empty)
	doClassAssignment this_object,"var_pagesize",pagesize
	doClassAssignment this_object,"var_printpagesize",IIF(IsLess(printpagesize,0),pagesize,printpagesize)
	if bValue(this_object.searchClauseObj.isUsedSrch()) then
		this_object.var_sql.initWhere 
	end if
	this_object.init 
End Function
Function method_ReportLogic_init(byref this_object,ByVal from)
	method_Summarable_init this_object,from
	doClassAssignment this_object.var_sql,"var_from",from
	doClassAssignment this_object,"var_list",CreateDictionary()
	this_object.var_totalRecords = 0
	doClassAssignment this_object,"var_pages",CreateDictionary()
	doClassAssignment this_object,"var_groupKeys",CreateDictionary()
	this_object.var_lastPageNumber = 0
	doClassAssignment this_object,"var_pageSummary",CreateDictionary()
	this_object.var_printRecordCount = 0
	this_object.var_listedRows = 0
	doClassAssignment this_object,"var_oldLevels",CreateDictionary()
End Function
Function method_ReportLogic_getWhere(byref this_object)
	doAssignmentByRef method_ReportLogic_getWhere,this_object.var_sql.getWhere()
	Exit Function
End Function
Function method_ReportLogic_getPages(byref this_object)
	doAssignmentByRef method_ReportLogic_getPages,this_object.var_pages
	Exit Function
End Function
Function method_ReportLogic_getFormattedRow(byref this_object,ByVal value)
End Function
Function method_ReportLogic_writeGroup(byref this_object,ByRef begin,ByRef var_end,ByVal gkey,ByVal grp,ByVal nField)
End Function
Function method_ReportLogic__writePage(byref this_object,ByRef page,ByVal src,ByVal count)
End Function
Function method_ReportLogic_writeGlobalSummary(byref this_object,ByVal source)
End Function
Function method_ReportLogic_writePageSummary(byref this_object)
	Dim nCnt,result,page,this
	if bValue(this_object.var_doPaging) then
		nCnt = 0
		do while IsLess(nCnt,asp_count(this_object.var_list))
			if not (not IsEmpty(ArrayElement(this_object.var_pages,nCnt))) then
				setArrElement this_object.var_pages,nCnt,CreateDictionary()
			end if
			doAssignmentByRef result,ArrayElement(this_object.var_pages,nCnt)
			if not IsEmpty(ArrayElement(this_object.var_pageSummary,nCnt)) then
				doAssignment page,ArrayElement(this_object.var_pageSummary,nCnt)
				this_object.func_writePage_p3 result,IIF(not IsEmpty(ArrayElement(page,"summary")),ArrayElement(page,"summary"),CreateDictionary()),IIF(not IsEmpty(ArrayElement(page,"count")),ArrayElement(page,"count"),0)
			else
				this_object.func_writePage_p3 result,CreateDictionary(),0
			end if
			nCnt = CSmartDbl(nCnt)+1
		loop
	else
		Set result = (CreateDictionary())
		doAssignment page,this_object.var_summary
		this_object.func_writePage_p3 result,IIF(not IsEmpty(ArrayElement(page,"summary")),ArrayElement(page,"summary"),CreateDictionary()),IIF(not IsEmpty(ArrayElement(page,"count")),ArrayElement(page,"count"),0)
		doClassAssignment this_object,"var_summary",result
	end if
	if IsEqual(0,asp_count(this_object.var_pages)) and IsLess(0,asp_count(this_object.var_list)) then
		setArrElement this_object.var_pages,asp_count(this_object.var_pages),this_object.var_summary
	end if
End Function
Function method_ReportLogic_makeSummary(byref this_object)
	Dim this
	this_object.func_makeSummary_p2 this_object.var_summary,0
End Function
Function method_ReportLogic_setSummary(byref this_object,ByVal recordsMode,ByVal data,ByVal rowToAppend)
	Dim level,setBegin,recordkeys,i,field,changed,nKey,nKey2,emptyRow,this,levels,added,nCnt,nPage
	doAssignmentByRef level,this_object.var_summary
	setBegin = false
	if bValue(this_object.repGroupFieldsCount) then
		Set recordkeys = (CreateDictionary())
		i = 0
		do while IsLess(i,asp_count(this_object.repGroupFields))
			doAssignmentByRef field,this_object.var_sql.field_p1(CSmartDbl(ArrayElement(ArrayElement(this_object.repGroupFields,i),"groupOrder"))-1)
			setArrElement recordkeys,CSmartDbl(ArrayElement(ArrayElement(this_object.repGroupFields,i),"groupOrder"))-1,field.getKey_p1(data)
			i = CSmartDbl(i)+1
		loop
		if IsLess(0,asp_count(this_object.var_groupKeys)) then
			changed = false
			nKey = 0
			do while IsLess(nKey,asp_count(recordkeys))
				if not IsEqual(ArrayElement(recordkeys,nKey),ArrayElement(this_object.var_groupKeys,nKey)) then
					changed = true
					exit do
				end if
				nKey = CSmartDbl(nKey)+1
			loop
			if bValue(changed) then
				nKey2 = CSmartDbl(asp_count(recordkeys))-1
				do while IsLessOrEqual(nKey,nKey2)
					doAssignmentByRef emptyRow,this_object.appendRow_p1(CreateDictionary())
					doAssignmentByRef field,this_object.var_sql.field_p1(nKey2)
					this_object.var_printRecordCount = CSmartDbl(this_object.var_printRecordCount)+CSmartDbl(field.var_rowsInSummary)
					this_object.var_listedRows = CSmartDbl(this_object.var_listedRows)+1
					setArrElementByRefN this_object.var_oldLevels,CreateArray2(nKey2,"_end"),emptyRow
					nKey2 = CSmartDbl(nKey2)-1
				loop
			end if
		end if
		doClassAssignment this_object,"var_groupKeys",recordkeys
		Set levels = (CreateDictionary())
		i = 0
		do while IsLess(i,asp_count(this_object.repGroupFields))
			if not (not IsEmpty(ArrayElement(ArrayElement(level,"values"),ArrayElement(recordkeys,CSmartDbl(ArrayElement(ArrayElement(this_object.repGroupFields,i),"groupOrder"))-1)))) then
				setArrElementN level,CreateArray2("values",ArrayElement(recordkeys,CSmartDbl(ArrayElement(ArrayElement(this_object.repGroupFields,i),"groupOrder"))-1)),CreateDictionary()
				doAssignmentByRef level,ArrayElement(ArrayElement(level,"values"),ArrayElement(recordkeys,CSmartDbl(ArrayElement(ArrayElement(this_object.repGroupFields,i),"groupOrder"))-1))
				doAssignmentByRef field,this_object.var_sql.field_p1(CSmartDbl(ArrayElement(ArrayElement(this_object.repGroupFields,i),"groupOrder"))-1)
				this_object.var_printRecordCount = CSmartDbl(this_object.var_printRecordCount)+CSmartDbl(field.var_rowsInHeader)
				setBegin = true
				setArrElement level,"_first",data
			else
				doAssignmentByRef level,ArrayElement(ArrayElement(level,"values"),ArrayElement(recordkeys,CSmartDbl(ArrayElement(ArrayElement(this_object.repGroupFields,i),"groupOrder"))-1))
			end if
			setArrElementByRef levels,asp_count(levels),level
			i = CSmartDbl(i)+1
		loop
		this_object.addSummary_p4 recordsMode,level,data,this_object.var_totalRecords
		doClassAssignmentByRef this_object,"var_oldLevels",levels
	else
		this_object.addSummary_p4 recordsMode,level,data,this_object.var_totalRecords
	end if
	if bValue(rowToAppend) then
		doAssignmentByRef added,this_object.appendRow_p1(rowToAppend)
		this_object.var_printRecordCount = CSmartDbl(this_object.var_printRecordCount)+1
		this_object.var_listedRows = CSmartDbl(this_object.var_listedRows)+1
		if bValue(setBegin) and bValue(this_object.repGroupFieldsCount) then
			nCnt = 0
			do while IsLess(nCnt,asp_count(levels))
				if not (not IsEmpty(ArrayElement(ArrayElement(levels,nCnt),"_begin"))) then
					setArrElementByRefN levels,CreateArray2(nCnt,"_begin"),added
				end if
				nCnt = CSmartDbl(nCnt)+1
			loop
		end if
	end if
	if bValue(this_object.repPageSummary) then
		if bValue(this_object.var_doPaging) and bValue(rowToAppend) then
			nPage = CSmartDbl(asp_count(this_object.var_list))-1
			if not (not IsEmpty(ArrayElement(this_object.var_pageSummary,nPage))) then
				setArrElementN this_object.var_pageSummary,CreateArray2(nPage,"count"),0
			end if
			this_object.addSummary_p4 recordsMode,ArrayElement(this_object.var_pageSummary,nPage),data,ArrayElement(ArrayElement(this_object.var_pageSummary,nPage),"count")
		end if
	end if
End Function
Function method_ReportLogic_setFinish(byref this_object)
	Dim nKey,field,emptyRow,this
	if IsLess(0,asp_count(this_object.var_groupKeys)) then
		nKey = CSmartDbl(asp_count(this_object.var_groupKeys))-1
		do while IsLessOrEqual(0,nKey)
			doAssignmentByRef field,this_object.var_sql.field_p1(nKey)
			this_object.var_printRecordCount = CSmartDbl(this_object.var_printRecordCount)+CSmartDbl(field.var_rowsInSummary)
			doAssignmentByRef emptyRow,this_object.appendRow_p1(CreateDictionary())
			this_object.var_listedRows = CSmartDbl(this_object.var_listedRows)+1
			setArrElementByRefN this_object.var_oldLevels,CreateArray2(nKey,"_end"),emptyRow
			nKey = CSmartDbl(nKey)-1
		loop
	end if
End Function
Function method_ReportLogic_appendRow(byref this_object,ByVal row)
	Dim page
	if bValue(this_object.var_doPaging) then
		doAssignmentByRef page,asp_intval(CSmartDbl(this_object.var_printRecordCount)/CSmartDbl(this_object.var_printpagesize))
		if IsLess(0,page) and not (not IsEmpty(ArrayElement(this_object.var_list,CSmartDbl(page)-1))) then
			Response.End
		end if
		setArrElementN this_object.var_list,CreateArray2(page,empty),row
		doAssignmentByRef method_ReportLogic_appendRow,ArrayElement(ArrayElement(this_object.var_list,page),CSmartDbl(asp_count(ArrayElement(this_object.var_list,page)))-1)
		Exit Function
	else
		setArrElement this_object.var_list,asp_count(this_object.var_list),row
		doAssignmentByRef method_ReportLogic_appendRow,ArrayElement(this_object.var_list,CSmartDbl(asp_count(this_object.var_list))-1)
		Exit Function
	end if
End Function
Function method_ReportLogic_recordVisible(byref this_object,ByVal nRecord)
	method_ReportLogic_recordVisible = (((bValue(this_object.var_doPaging) or IsEqual(this_object.var_sql.var_limitLevel,1)) or IsEqual(this_object.var_pagesize,-1)) or IsEqual(this_object.var_sql.var_limitLevel,2) and (IsLessOrEqual(0,CSmartDbl(nRecord)-CSmartDbl(this_object.var_sql.var_skipCount)) and IsLess(CSmartDbl(nRecord)-CSmartDbl(this_object.var_sql.var_skipCount),this_object.var_pagesize))) or IsEqual(this_object.var_sql.var_limitLevel,0) and (IsLessOrEqual(0,CSmartDbl(nRecord)-CSmartDbl(this_object.var_from)) and IsLess(CSmartDbl(nRecord)-CSmartDbl(this_object.var_sql.var_skipCount),CSmartDbl(this_object.var_from)+CSmartDbl(this_object.var_pagesize)))
	Exit Function
End Function
Function method_ReportLogic_getTotals(byref this_object)
	Dim totals,sql,totalRecords,cursor,data,this
	if bValue(this_object.var_fullRequest) then
		doAssignmentByRef method_ReportLogic_getTotals,this_object.var_groups.getSummary()
		Exit Function
	else
		if bValue(this_object.var_groups.allGroupsUsed()) then
			doAssignmentByRef method_ReportLogic_getTotals,this_object.var_summary
			Exit Function
		else
			Set totals = (CreateDictionary())
			doAssignmentByRef sql,this_object.var_sql.sqlt()
			if not IsFalse(sql) then
				totalRecords = 0
				doAssignmentByRef cursor,db_query(sql,this_object.var_connection)
				doAssignmentByRef data,db_fetch_array(cursor)
				this_object.addSummary_p4 false,totals,data,totalRecords
			end if
			doAssignmentByRef method_ReportLogic_getTotals,totals
			Exit Function
		end if
	end if
End Function
Function method_ReportLogic_getReport(byref this_object,ByVal from)
	Dim this,isExistTimeFormatField,i,page,nRow,groups,hsql,hwhere,eventsObj,sql,cursor,data,visible,global_totals,globals,countrows,countGroups,maxpages,returnthis
	this_object.init_p1 from
	this_object.var_doPaging = IsEqual(from,-1)
	isExistTimeFormatField = false
	i = 0
	do while IsLess(i,asp_count(this_object.fieldsArr))
		if IsEqual(ArrayElement(ArrayElement(this_object.fieldsArr,i),"viewFormat"),"Time") then
			isExistTimeFormatField = true
			exit do
		end if
		i = CSmartDbl(i)+1
	loop
	this_object.var_fullRequest = bValue(this_object.var_doPaging) or bValue(this_object.repGlobalSummary) and bValue(isExistTimeFormatField)
	if not bValue(bSubqueriesSupported) then
		this_object.var_fullRequest = true
	end if
	doClassAssignment this_object,"var_recordBasedRequest",this_object.var_fullRequest
	if not bValue(this_object.repGroupFieldsCount) then
		this_object.var_recordBasedRequest = true
	end if
	this_object.var_sql.setRecordBasedRequest_p1 this_object.var_recordBasedRequest
	if bValue(this_object.var_doPaging) or bValue(this_object.var_fullRequest) then
		this_object.var_sql.var_limitLevel = 0
	else
		if not bValue(this_object.repGroupFieldsCount) then
			this_object.var_sql.var_limitLevel = 2
		else
			this_object.var_sql.var_limitLevel = 1
		end if
	end if
	page = -1
	nRow = 0
	if not bValue(this_object.var_recordBasedRequest) then
		doAssignmentByRef groups,this_object.var_groups.getDisplayGroups_p1(from)
		doAssignmentByRef hsql,this_object.var_sql.sql2_p1(groups)
		if bValue(tableEventExists("BeforeQueryReport",this_object.tName)) then
			doAssignment hwhere,ArrayElement(hsql,"where")
			doAssignmentByRef eventsObj,getEventObject(this_object.tName)
			eventsObj.BeforeQueryReport_p1 hwhere
			setArrElement hsql,"where",hwhere
		end if
		doAssignmentByRef sql,this_object.var_sql.buildsql_p1(hsql)
		doAssignmentByRef cursor,db_query(sql,this_object.var_connection)
		do while bValue(doAssignmentByRef(data,db_fetch_array(cursor)))
			this_object.setSummary_p3 this_object.repShowDet,data,IIF(this_object.recordVisible_p1(nRow),this_object.getFormattedRow_p1(data),null)
			nRow = CSmartDbl(nRow)+1
		loop
	else
		this_object.var_groups.init_p1 from
		this_object.var_sql.setOldAlgorithm 
		doAssignmentByRef hsql,this_object.var_sql.sql2_p1(null)
		doAssignmentByRef sql,this_object.var_sql.buildsql_p1(hsql)
		doAssignmentByRef cursor,db_query(sql,this_object.var_connection)
		do while bValue(doAssignmentByRef(data,db_fetch_array(cursor)))
			if bValue(this_object.repGroupFieldsCount) then
				this_object.var_groups.setGroup_p1 data
			end if
			if bValue(this_object.var_fullRequest) then
				this_object.var_groups.setGlobalSummary_p2 true,data
			end if
			if bValue(this_object.repGroupFieldsCount) then
				visible = (bValue(this_object.var_doPaging) or bValue(this_object.var_groups.isVisibleGroup())) or IsEqual(this_object.var_pagesize,-1)
			else
				doAssignmentByRef visible,this_object.recordVisible_p1(nRow)
			end if
			if bValue(visible) then
				this_object.setSummary_p3 true,data,this_object.getFormattedRow_p1(data)
			else
				if not bValue(this_object.var_fullRequest) and IsLess(0,asp_count(this_object.var_list)) then
					exit do
				end if
			end if
			nRow = CSmartDbl(nRow)+1
		loop
		this_object.var_sql.setOldAlgorithm_p1 false
	end if
	this_object.setFinish 
	this_object.makeSummary 
	doAssignmentByRef global_totals,this_object.getTotals()
	this_object.writePageSummary 
	doAssignmentByRef globals,this_object.writeGlobalSummary_p1(global_totals)
	if bValue(this_object.repGroupFieldsCount) then
		doAssignmentByRef countrows,this_object.var_groups.getCountGroups_p1(this_object.var_fullRequest)
		doAssignment countGroups,countrows
	else
		doAssignment countrows,ArrayElement(global_totals,"count")
		countGroups = 1
	end if
	maxpages = 1
	if IsLess(0,this_object.var_pagesize) then
		doAssignmentByRef maxpages,asp_ceil(CSmartDbl(countrows)/CSmartDbl(this_object.var_pagesize))
	end if
	Set returnthis = (CreateDictionary6("list",this_object.var_list,"global",globals,"page",this_object.var_summary,"maxpages",maxpages,"countRows",nRow,"countGroups",countGroups))
	doAssignmentByRef method_ReportLogic_getReport,returnthis
	Exit Function
End Function

'------ Class Report extends ReportLogic ------
Class Report
	Public forExport
	Public mode
	Public var_list
	Public var_totalRecords
	Public var_pages
	Public var_pagesize
	Public var_printpagesize
	Public var_from
	Public var_connection
	Public var_sql
	Public var_groups
	Public var_groupKeys
	Public var_fullRequest
	Public var_recordBasedRequest
	Public var_doPaging
	Public var_lastPageNumber
	Public var_pageSummary
	Public var_printRecordCount
	Public var_listedRows
	Public var_oldLevels
	Public searchClauseObj
	Public var_summary
	Public tName
	Public shortTName
	Public repGroupFieldsCount
	Public repPageSummary
	Public repGlobalSummary
	Public repLayout
	Public showGroupSummaryCount
	Public repShowDet
	Public repGroupFields
	Public tKeyFields
	Public isExistTotalFields
	Public fieldsArr
	Public Function init_Report_p7(ByVal sql,ByVal order,ByRef searchClauseObj,ByVal connection,ByVal pagesize,ByVal printpagesize,ByRef params)
		DoAssignmentByRef init_Report_p7,method_Report_Report(me,sql,order,searchClauseObj,connection,pagesize,printpagesize,params)
	End Function
	Public Function getFormattedRow_p1(ByVal value)
		DoAssignmentByRef getFormattedRow_p1,method_Report_getFormattedRow(me,value)
	End Function
	Public Function writeGroup_p5(ByRef begin,ByRef var_end,ByVal gkey,ByVal grp,ByVal nField)
		DoAssignmentByRef writeGroup_p5,method_Report_writeGroup(me,begin,var_end,gkey,grp,nField)
	End Function
	Public Function func_writePage_p3(ByRef page,ByVal src,ByVal count)
		DoAssignmentByRef func_writePage_p3,method_Report__writePage(me,page,src,count)
	End Function
	Public Function writeGlobalSummary_p1(ByVal source)
		DoAssignmentByRef writeGlobalSummary_p1,method_Report_writeGlobalSummary(me,source)
	End Function
	Public Function init_p1(ByVal from)
		DoAssignmentByRef init_p1,method_ReportLogic_init(me,from)
	End Function
	Public Function init()
		DoAssignmentByRef init,method_ReportLogic_init(me,0)
	End Function
	Public Function getWhere()
		DoAssignmentByRef getWhere,method_ReportLogic_getWhere(me)
	End Function
	Public Function getPages()
		DoAssignmentByRef getPages,method_ReportLogic_getPages(me)
	End Function
	Public Function writePageSummary()
		DoAssignmentByRef writePageSummary,method_ReportLogic_writePageSummary(me)
	End Function
	Public Function makeSummary()
		DoAssignmentByRef makeSummary,method_ReportLogic_makeSummary(me)
	End Function
	Public Function setSummary_p3(ByVal recordsMode,ByVal data,ByVal rowToAppend)
		DoAssignmentByRef setSummary_p3,method_ReportLogic_setSummary(me,recordsMode,data,rowToAppend)
	End Function
	Public Function setSummary_p2(ByVal recordsMode,ByVal data)
		DoAssignmentByRef setSummary_p2,method_ReportLogic_setSummary(me,recordsMode,data,null)
	End Function
	Public Function setFinish()
		DoAssignmentByRef setFinish,method_ReportLogic_setFinish(me)
	End Function
	Public Function appendRow_p1(ByVal row)
		DoAssignmentByRef appendRow_p1,method_ReportLogic_appendRow(me,row)
	End Function
	Public Function recordVisible_p1(ByVal nRecord)
		DoAssignmentByRef recordVisible_p1,method_ReportLogic_recordVisible(me,nRecord)
	End Function
	Public Function getTotals()
		DoAssignmentByRef getTotals,method_ReportLogic_getTotals(me)
	End Function
	Public Function getReport_p1(ByVal from)
		DoAssignmentByRef getReport_p1,method_ReportLogic_getReport(me,from)
	End Function
	Public Function getReport()
		DoAssignmentByRef getReport,method_ReportLogic_getReport(me,0)
	End Function
	Public Function addSummary_p4(ByVal recordsMode,ByRef summary,ByVal data,ByRef nTotalRecords)
		DoAssignmentByRef addSummary_p4,method_Summarable_addSummary(me,recordsMode,summary,data,nTotalRecords)
	End Function
	Public Function func_makeSummary_p2(ByRef summary,ByVal deep)
		DoAssignmentByRef func_makeSummary_p2,method_Summarable__makeSummary(me,summary,deep)
	End Function
	Public Function value2time_p1(ByVal value)
		DoAssignmentByRef value2time_p1,method_Summarable_value2time(me,value)
	End Function
	Public Function time2printable_p1(ByVal time)
		DoAssignmentByRef time2printable_p1,method_Summarable_time2printable(me,time)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"forExport", forExport
		setArrElement out,"mode", mode
		setArrElement out,"var_list", var_list
		setArrElement out,"var_totalRecords", var_totalRecords
		setArrElement out,"var_pages", var_pages
		setArrElement out,"var_pagesize", var_pagesize
		setArrElement out,"var_printpagesize", var_printpagesize
		setArrElement out,"var_from", var_from
		setArrElement out,"var_connection", var_connection
		setArrElement out,"var_sql", var_sql
		setArrElement out,"var_groups", var_groups
		setArrElement out,"var_groupKeys", var_groupKeys
		setArrElement out,"var_fullRequest", var_fullRequest
		setArrElement out,"var_recordBasedRequest", var_recordBasedRequest
		setArrElement out,"var_doPaging", var_doPaging
		setArrElement out,"var_lastPageNumber", var_lastPageNumber
		setArrElement out,"var_pageSummary", var_pageSummary
		setArrElement out,"var_printRecordCount", var_printRecordCount
		setArrElement out,"var_listedRows", var_listedRows
		setArrElement out,"var_oldLevels", var_oldLevels
		setArrElement out,"searchClauseObj", searchClauseObj
		setArrElement out,"var_summary", var_summary
		setArrElement out,"tName", tName
		setArrElement out,"shortTName", shortTName
		setArrElement out,"repGroupFieldsCount", repGroupFieldsCount
		setArrElement out,"repPageSummary", repPageSummary
		setArrElement out,"repGlobalSummary", repGlobalSummary
		setArrElement out,"repLayout", repLayout
		setArrElement out,"showGroupSummaryCount", showGroupSummaryCount
		setArrElement out,"repShowDet", repShowDet
		setArrElement out,"repGroupFields", repGroupFields
		setArrElement out,"tKeyFields", tKeyFields
		setArrElement out,"isExistTotalFields", isExistTotalFields
		setArrElement out,"fieldsArr", fieldsArr
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment forExport, dict("forExport")
		DoAssignment mode, dict("mode")
		DoAssignment var_list, dict("var_list")
		DoAssignment var_totalRecords, dict("var_totalRecords")
		DoAssignment var_pages, dict("var_pages")
		DoAssignment var_pagesize, dict("var_pagesize")
		DoAssignment var_printpagesize, dict("var_printpagesize")
		DoAssignment var_from, dict("var_from")
		DoAssignment var_connection, dict("var_connection")
		DoAssignment var_sql, dict("var_sql")
		DoAssignment var_groups, dict("var_groups")
		DoAssignment var_groupKeys, dict("var_groupKeys")
		DoAssignment var_fullRequest, dict("var_fullRequest")
		DoAssignment var_recordBasedRequest, dict("var_recordBasedRequest")
		DoAssignment var_doPaging, dict("var_doPaging")
		DoAssignment var_lastPageNumber, dict("var_lastPageNumber")
		DoAssignment var_pageSummary, dict("var_pageSummary")
		DoAssignment var_printRecordCount, dict("var_printRecordCount")
		DoAssignment var_listedRows, dict("var_listedRows")
		DoAssignment var_oldLevels, dict("var_oldLevels")
		DoAssignment searchClauseObj, dict("searchClauseObj")
		DoAssignment var_summary, dict("var_summary")
		DoAssignment tName, dict("tName")
		DoAssignment shortTName, dict("shortTName")
		DoAssignment repGroupFieldsCount, dict("repGroupFieldsCount")
		DoAssignment repPageSummary, dict("repPageSummary")
		DoAssignment repGlobalSummary, dict("repGlobalSummary")
		DoAssignment repLayout, dict("repLayout")
		DoAssignment showGroupSummaryCount, dict("showGroupSummaryCount")
		DoAssignment repShowDet, dict("repShowDet")
		DoAssignment repGroupFields, dict("repGroupFields")
		DoAssignment tKeyFields, dict("tKeyFields")
		DoAssignment isExistTotalFields, dict("isExistTotalFields")
		DoAssignment fieldsArr, dict("fieldsArr")
	End Sub
' end serialize
End Class
' Report implementation 
Function method_Report_Report(byref this_object,ByVal sql,ByVal order,ByRef searchClauseObj,ByVal connection,ByVal pagesize,ByVal printpagesize,ByRef params)
	this_object.forExport = false
	this_object.mode = MODE_LIST
	this_object.var_from = 0
	this_object.var_fullRequest = false
	this_object.var_recordBasedRequest = false
	this_object.var_doPaging = false
	this_object.var_lastPageNumber = 0
	this_object.var_printRecordCount = 0
	this_object.var_listedRows = 0
	this_object.searchClauseObj = null
	doClassAssignment this_object,"var_summary",CreateDictionary()
	this_object.tName = ""
	this_object.shortTName = ""
	this_object.repGroupFieldsCount = 0
	this_object.repPageSummary = 0
	this_object.repGlobalSummary = 0
	this_object.repLayout = 0
	this_object.showGroupSummaryCount = 0
	this_object.repShowDet = 0
	doClassAssignment this_object,"repGroupFields",CreateDictionary()
	doClassAssignment this_object,"tKeyFields",CreateDictionary()
	this_object.isExistTotalFields = false
	doClassAssignment this_object,"fieldsArr",CreateDictionary()
	Dim this
	RunnerApply this_object,params
	method_ReportLogic_ReportLogic this_object,sql,order,searchClauseObj,connection,pagesize,printpagesize,params
End Function
Function method_Report_getFormattedRow(byref this_object,ByVal value)
	Dim row,keylink,i,pass,j,val,thumbname,filename
	Set row = (CreateDictionary1("row_data",true))
	keylink = ""
	i = 0
	do while IsLess(i,asp_count(this_object.tKeyFields))
		keylink = CSmartStr(keylink) & ((("&key" & CSmartStr(CSmartDbl(i)+1)) & "=") & CSmartStr(htmlspecialchars(asp_rawurlencode(ArrayElement(value,ArrayElement(this_object.tKeyFields,i))))))
		i = CSmartDbl(i)+1
	loop
	i = 0
	c_reportlib_exitLoop36=false
	do while IsLess(i,asp_count(this_object.fieldsArr))
		c_reportlib_exitLoop36=false
		do
			pass = false
			j = 0
			do while IsLess(j,asp_count(this_object.repGroupFields))
				if not bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"repPage")) or not (bValue(this_object.repShowDet) or IsEqual(ArrayElement(ArrayElement(this_object.repGroupFields,j),"strGroupField"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")) and IsIdentical(ArrayElement(ArrayElement(this_object.repGroupFields,j),"groupInterval"),0)) then
					pass = true
				end if
				j = CSmartDbl(j)+1
			loop
			if bValue(pass) then
				exit do
			end if
			if IsEqual(ArrayElement(ArrayElement(this_object.fieldsArr,i),"viewFormat"),FORMAT_DATABASE_IMAGE) then
				if not bValue(this_object.forExport) then
					if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"showThumb")) then
						val = CSmartStr(val) & "<a "
						if bValue(IsUseiBox(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),this_object.tName)) then
							val = CSmartStr(val) & " rel='ibox'"
						else
							val = CSmartStr(val) & " target=_blank"
						end if
						val = CSmartStr(val) & (((((" href=""imager.asp?table=" & CSmartStr(this_object.shortTName)) & "&field=") & CSmartStr(asp_rawurlencode(htmlspecialchars(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"))))) & CSmartStr(keylink)) & """>")
						val = CSmartStr(val) & "<img border=0"
						if bValue(isEnableSection508()) then
							val = CSmartStr(val) & " alt=""Image from DB"""
						end if
						val = CSmartStr(val) & (((((((" src=""imager.asp?table=" & CSmartStr(this_object.shortTName)) & "&field=") & CSmartStr(asp_rawurlencode(htmlspecialchars(ArrayElement(ArrayElement(this_object.fieldsArr,i),"thumbnail"))))) & "&alt=") & CSmartStr(asp_rawurlencode(htmlspecialchars(ArrayElement(ArrayElement(this_object.fieldsArr,i),"repPage"))))) & CSmartStr(keylink)) & """>")
						val = CSmartStr(val) & "</a>"
					else
						val = "<img"
						if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"imageWidth")) then
							val = CSmartStr(val) & (" width=" & CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"imageWidth")))
						end if
						if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"imageHeight")) then
							val = CSmartStr(val) & (" height=" & CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"imageHeight")))
						end if
						val = CSmartStr(val) & " border=0"
						if bValue(isEnableSection508()) then
							val = CSmartStr(val) & " alt=""Image from DB"""
						end if
						val = CSmartStr(val) & (((((" src=""imager.asp?table=" & CSmartStr(this_object.shortTName)) & "&field=") & CSmartStr(asp_rawurlencode(htmlspecialchars(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"))))) & CSmartStr(keylink)) & """>")
					end if
				else
					val = "LONG BINARY DATA - CANNOT BE DISPLAYED"
				end if
			else
				if IsEqual(ArrayElement(ArrayElement(this_object.fieldsArr,i),"viewFormat"),FORMAT_FILE_IMAGE) then
					if not bValue(this_object.forExport) then
						if bValue(CheckImageExtension(ArrayElement(value,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")))) then
							if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"showThumb")) then
								thumbname = CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"thumbnail")) & CSmartStr(ArrayElement(value,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")))
								if not IsEqual(asp_substr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"strhlPrefix"),0,7),"http://") and not bValue(myfile_exists(getabspath(CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"strhlPrefix")) & CSmartStr(thumbname)))) then
									doAssignment thumbname,ArrayElement(value,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"))
								end if
								val = "<a"
								if bValue(IsUseiBox(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),this_object.tName)) then
									val = CSmartStr(val) & " rel='ibox'"
								else
									val = CSmartStr(val) & " target=_blank"
								end if
								val = CSmartStr(val) & ((" href=""" & CSmartStr(htmlspecialchars(AddLinkPrefix(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),ArrayElement(value,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"")))) & """>")
								val = CSmartStr(val) & "<img"
								if IsEqual(thumbname,ArrayElement(value,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"))) then
									if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"imageWidth")) then
										val = CSmartStr(val) & (" width=" & CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"imageWidth")))
									end if
									if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"imageHeight")) then
										val = CSmartStr(val) & (" height=" & CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"imageHeight")))
									end if
								end if
								val = CSmartStr(val) & " border=0"
								if bValue(isEnableSection508()) then
									val = CSmartStr(val) & ((" alt=""" & CSmartStr(htmlspecialchars(ArrayElement(value,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"))))) & """")
								end if
								val = CSmartStr(val) & ((" src=""" & CSmartStr(htmlspecialchars(AddLinkPrefix(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),thumbname,"")))) & """></a>")
							else
								val = "<img"
								if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"imageWidth")) then
									val = CSmartStr(val) & (" width=" & CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"imageWidth")))
								end if
								if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"imageHeight")) then
									val = CSmartStr(val) & (" height=" & CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"imageHeight")))
								end if
								val = CSmartStr(val) & " border=0"
								if bValue(isEnableSection508()) then
									val = CSmartStr(val) & ((" alt=""" & CSmartStr(htmlspecialchars(ArrayElement(value,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"))))) & """")
								end if
								val = CSmartStr(val) & ((" src=""" & CSmartStr(htmlspecialchars(AddLinkPrefix(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),ArrayElement(value,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"")))) & """>")
							end if
						end if
					else
						val = "LONG BINARY DATA - CANNOT BE DISPLAYED"
					end if
				else
					if IsEqual(ArrayElement(ArrayElement(this_object.fieldsArr,i),"viewFormat"),FORMAT_DATABASE_FILE) then
						if not bValue(this_object.forExport) then
							if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"fileName")) then
								doAssignment filename,ArrayElement(value,ArrayElement(ArrayElement(this_object.fieldsArr,i),"fileName"))
								if not bValue(filename) then
									filename = "file.bin"
								end if
							else
								filename = "file.bin"
							end if
							if bValue(asp_strlen(ArrayElement(value,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")))) then
								val = (((((("<a href=""getfile.asp?table=" & CSmartStr(this_object.shortTName)) & "&filename=") & CSmartStr(asp_rawurlencode(filename))) & "&field=") & CSmartStr(asp_rawurlencode(htmlspecialchars(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"))))) & CSmartStr(keylink)) & """>"
								val = CSmartStr(val) & CSmartStr(htmlspecialchars(filename))
								val = CSmartStr(val) & "</a>"
							end if
						else
							val = "LONG BINARY DATA - CANNOT BE DISPLAYED"
						end if
					else
						if (IsEqual(ArrayElement(ArrayElement(this_object.fieldsArr,i),"editFormat"),EDIT_FORMAT_LOOKUP_WIZARD) or IsEqual(ArrayElement(ArrayElement(this_object.fieldsArr,i),"editFormat"),EDIT_FORMAT_RADIO)) and IsEqual(GetLookupType(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),this_object.tName),LT_LOOKUPTABLE) then
							doAssignmentByRef val,DisplayLookupWizard(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),ArrayElement(value,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),value,keylink,this_object.mode)
						else
							if bValue(NeedEncode(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),this_object.tName)) then
								doAssignmentByRef val,ProcessLargeText(GetData(value,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"viewFormat")),("field=" & CSmartStr(asp_rawurlencode(htmlspecialchars(ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"))))) & CSmartStr(keylink),"",this_object.mode,"")
							else
								if IsEqual(ArrayElement(ArrayElement(this_object.fieldsArr,i),"viewFormat"),FORMAT_CHECKBOX) and bValue(this_object.forExport) then
									doAssignmentByRef val,GetData(value,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),FORMAT_NONE)
								else
									doAssignmentByRef val,GetData(value,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"viewFormat"))
								end if
							end if
						end if
					end if
				end if
			end if
			setArrElement row,CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"goodName")) & "_value",val
		loop while false
		if c_reportlib_exitLoop36 then _
			exit do
		i = CSmartDbl(i)+1
	loop
	if IsEqual(this_object.repLayout,REPORT_BLOCK) then
		setArrElement row,GoodFieldName("nonewgroup"),true
	end if
	doAssignmentByRef method_Report_getFormattedRow,row
	Exit Function
End Function
Function method_Report_writeGroup(byref this_object,ByRef begin,ByRef var_end,ByVal gkey,ByVal grp,ByVal nField)
	Dim field,gname,i,bFound,nG,gname2,j,gvalue,formattedValue
	doAssignmentByRef field,this_object.var_sql.field_p1(nField)
	doAssignmentByRef gname,field.name()
	i = 0
	do while IsLess(i,asp_count(this_object.repGroupFields))
		if IsEqual(gname,ArrayElement(ArrayElement(this_object.repGroupFields,i),"strGroupField")) then
			if IsEqual(this_object.repLayout,REPORT_BLOCK) then
				bFound = false
				nG = 0
				do while IsLess(nG,this_object.repGroupFieldsCount)
					doAssignmentByRef field,this_object.var_sql.field_p1(nG)
					doAssignmentByRef gname2,field.name()
					if IsLess(nG,nField) then
						if not IsEmpty(ArrayElement(begin,GoodFieldName(CSmartStr(gname2) & "_firstnewgroup"))) then
							bFound = true
						end if
					else
						asp_unsetElement begin,GoodFieldName(CSmartStr(gname2) & "_firstnewgroup")
					end if
					nG = CSmartDbl(nG)+1
				loop
				if not bValue(bFound) then
					setArrElement begin,GoodFieldName(CSmartStr(gname) & "_firstnewgroup"),true
				end if
				asp_unsetElement begin,GoodFieldName("nonewgroup")
			else
				setArrElement begin,GoodFieldName(CSmartStr(gname) & "_newgroup"),true
			end if
			setArrElement var_end,GoodFieldName(CSmartStr(gname) & "_endgroup"),true
			if bValue(ArrayElement(ArrayElement(this_object.repGroupFields,i),"showGroupSummary")) then
				setArrElement var_end,GoodFieldName(("group" & CSmartStr(gname)) & "_total_cnt"),ArrayElement(grp,"count")
			end if
			j = 0
			do while IsLess(j,asp_count(this_object.fieldsArr))
				if bValue(asp_is_array(ArrayElement(grp,"summary"))) then
					if bValue(asp_is_array(ArrayElement(ArrayElement(grp,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"name")))) then
						if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,j),"totalMax")) then
							setArrElement var_end,((("group" & CSmartStr(GoodFieldName(gname))) & "_total") & CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,j),"goodName"))) & "_max",getFormattedValue(ArrayElement(ArrayElement(ArrayElement(grp,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"name")),"MAX"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"name"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"viewFormat"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"editFormat"),this_object.mode)
						end if
						if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,j),"totalMin")) then
							setArrElement var_end,((("group" & CSmartStr(GoodFieldName(gname))) & "_total") & CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,j),"goodName"))) & "_min",getFormattedValue(ArrayElement(ArrayElement(ArrayElement(grp,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"name")),"MIN"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"name"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"viewFormat"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"editFormat"),this_object.mode)
						end if
						if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,j),"totalAvg")) then
							setArrElement var_end,((("group" & CSmartStr(GoodFieldName(gname))) & "_total") & CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,j),"goodName"))) & "_avg",getFormattedValue(ArrayElement(ArrayElement(ArrayElement(grp,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"name")),"AVG"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"name"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"viewFormat"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"editFormat"),this_object.mode)
						end if
						if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,j),"totalSum")) then
							setArrElement var_end,((("group" & CSmartStr(GoodFieldName(gname))) & "_total") & CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,j),"goodName"))) & "_sum",getFormattedValue(ArrayElement(ArrayElement(ArrayElement(grp,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"name")),"SUM"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"name"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"viewFormat"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"editFormat"),this_object.mode)
						end if
					end if
				end if
				if IsEqual(ArrayElement(ArrayElement(this_object.fieldsArr,j),"name"),ArrayElement(ArrayElement(this_object.repGroupFields,i),"strGroupField")) then
					doAssignmentByRef field,this_object.var_sql.field_p1(nField)
					doAssignmentByRef gvalue,field.getFieldName_p2(gkey,ArrayElement(grp,"_first"))
					if bValue(field.overrideFormat()) then
						setArrElement begin,GoodFieldName(CSmartStr(GoodFieldName(gname)) & "_grval"),htmlspecialchars(gvalue)
						if bValue(this_object.showGroupSummaryCount) then
							setArrElement var_end,GoodFieldName(CSmartStr(GoodFieldName(gname)) & "_grval"),htmlspecialchars(gvalue)
						end if
					else
						doAssignmentByRef formattedValue,getFormattedValue(gvalue,ArrayElement(ArrayElement(this_object.fieldsArr,j),"name"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"viewFormat"),ArrayElement(ArrayElement(this_object.fieldsArr,j),"editFormat"),this_object.mode)
						setArrElement begin,GoodFieldName(CSmartStr(gname) & "_grval"),htmlspecialchars(formattedValue)
						if bValue(this_object.showGroupSummaryCount) then
							setArrElement var_end,GoodFieldName(CSmartStr(gname) & "_grval"),htmlspecialchars(formattedValue)
						end if
					end if
				end if
				j = CSmartDbl(j)+1
			loop
		end if
		i = CSmartDbl(i)+1
	loop
End Function
Function method_Report__writePage(byref this_object,ByRef page,ByVal src,ByVal count)
	Dim i
	setArrElement page,"page_summary",true
	if bValue(this_object.repPageSummary) then
		i = 0
		do while IsLess(i,asp_count(this_object.fieldsArr))
			if bValue(asp_is_array(ArrayElement(src,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")))) then
				if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalSum")) then
					setArrElement page,("page_total" & CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"goodName"))) & "_sum",getFormattedValue(ArrayElement(ArrayElement(src,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"SUM"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"viewFormat"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"editFormat"),this_object.mode)
				end if
				if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalAvg")) then
					setArrElement page,("page_total" & CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"goodName"))) & "_avg",getFormattedValue(ArrayElement(ArrayElement(src,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"AVG"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"viewFormat"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"editFormat"),this_object.mode)
				end if
				if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalMin")) then
					setArrElement page,("page_total" & CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"goodName"))) & "_min",getFormattedValue(ArrayElement(ArrayElement(src,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MIN"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"viewFormat"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"editFormat"),this_object.mode)
				end if
				if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalMax")) then
					setArrElement page,("page_total" & CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"goodName"))) & "_max",getFormattedValue(ArrayElement(ArrayElement(src,ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MAX"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"viewFormat"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"editFormat"),this_object.mode)
				end if
			end if
			i = CSmartDbl(i)+1
		loop
		setArrElement page,"page_total_cnt",count
	end if
End Function
Function method_Report_writeGlobalSummary(byref this_object,ByVal source)
	Dim result,i
	Set result = (CreateDictionary())
	if bValue(this_object.repGlobalSummary) then
		if bValue(asp_is_array(ArrayElement(source,"summary"))) then
			i = 0
			do while IsLess(i,asp_count(this_object.fieldsArr))
				if bValue(asp_is_array(ArrayElement(ArrayElement(source,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")))) then
					if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalMax")) then
						setArrElement result,("global_total" & CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"goodName"))) & "_max",getFormattedValue(ArrayElement(ArrayElement(ArrayElement(source,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MAX"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"viewFormat"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"editFormat"),this_object.mode)
					end if
					if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalMin")) then
						setArrElement result,("global_total" & CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"goodName"))) & "_min",getFormattedValue(ArrayElement(ArrayElement(ArrayElement(source,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"MIN"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"viewFormat"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"editFormat"),this_object.mode)
					end if
					if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalAvg")) then
						setArrElement result,("global_total" & CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"goodName"))) & "_avg",getFormattedValue(ArrayElement(ArrayElement(ArrayElement(source,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"AVG"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"viewFormat"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"editFormat"),this_object.mode)
					end if
					if bValue(ArrayElement(ArrayElement(this_object.fieldsArr,i),"totalSum")) then
						setArrElement result,("global_total" & CSmartStr(ArrayElement(ArrayElement(this_object.fieldsArr,i),"goodName"))) & "_sum",getFormattedValue(ArrayElement(ArrayElement(ArrayElement(source,"summary"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name")),"SUM"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"name"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"viewFormat"),ArrayElement(ArrayElement(this_object.fieldsArr,i),"editFormat"),this_object.mode)
					end if
				end if
				i = CSmartDbl(i)+1
			loop
		end if
		setArrElement result,"global_total_cnt",ArrayElement(source,"count")
	end if
	doAssignmentByRef method_Report_writeGlobalSummary,result
	Exit Function
End Function
%>
