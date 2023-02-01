<%

'------ Class SearchClause ------
Class SearchClause
	Public var_where
	Public tName
	Public searchFieldsArr
	Public googleLikeFields
	Public srchType
	Public sessionPrefix
	Public bIsUsedSrch
	Public panelSearchFields
	Public Function init_SearchClause_p1(ByRef params)
		DoAssignmentByRef init_SearchClause_p1,method_SearchClause_SearchClause(me,params)
	End Function
	Public Function buildAdvancedWhere()
		DoAssignmentByRef buildAdvancedWhere,method_SearchClause_buildAdvancedWhere(me)
	End Function
	Public Function buildSimpleWhere()
		DoAssignmentByRef buildSimpleWhere,method_SearchClause_buildSimpleWhere(me)
	End Function
	Public Function buildItegratedWhere_p1(ByVal fieldsArr)
		DoAssignmentByRef buildItegratedWhere_p1,method_SearchClause_buildItegratedWhere(me,fieldsArr)
	End Function
	Public Function getWhere_p1(ByVal fieldsArr)
		DoAssignmentByRef getWhere_p1,method_SearchClause_getWhere(me,fieldsArr)
	End Function
	Public Function getSuggestWhere_p4(ByVal fName,ByVal fType,ByVal suggestAllContent,ByVal searchVal)
		DoAssignmentByRef getSuggestWhere_p4,method_SearchClause_getSuggestWhere(me,fName,fType,suggestAllContent,searchVal)
	End Function
	Public Function parseAdvancedRequest()
		DoAssignmentByRef parseAdvancedRequest,method_SearchClause_parseAdvancedRequest(me)
	End Function
	Public Function parseSimpleRequest()
		DoAssignmentByRef parseSimpleRequest,method_SearchClause_parseSimpleRequest(me)
	End Function
	Public Function parseItegratedRequest()
		DoAssignmentByRef parseItegratedRequest,method_SearchClause_parseItegratedRequest(me)
	End Function
	Public Function parseRequest()
		DoAssignmentByRef parseRequest,method_SearchClause_parseRequest(me)
	End Function
	Public Function clearSearch()
		DoAssignmentByRef clearSearch,method_SearchClause_clearSearch(me)
	End Function
	Public Function applyWhere_p3(ByRef sql,ByVal simpleFieldsArr,ByVal aggFieldsArr)
		DoAssignmentByRef applyWhere_p3,method_SearchClause_applyWhere(me,sql,simpleFieldsArr,aggFieldsArr)
	End Function
	Public Function getTable()
		DoAssignmentByRef getTable,method_SearchClause_getTable(me)
	End Function
	Public Function getSearchCtrlParams_p1(ByVal fName)
		DoAssignmentByRef getSearchCtrlParams_p1,method_SearchClause_getSearchCtrlParams(me,fName)
	End Function
	Public Function getUsedCtrlsCount()
		DoAssignmentByRef getUsedCtrlsCount,method_SearchClause_getUsedCtrlsCount(me)
	End Function
	Public Function getSearchGlobalParams()
		DoAssignmentByRef getSearchGlobalParams,method_SearchClause_getSearchGlobalParams(me)
	End Function
	Public Function getSrchPanelAttrs()
		DoAssignmentByRef getSrchPanelAttrs,method_SearchClause_getSrchPanelAttrs(me)
	End Function
	Public Function isUsedSrch()
		DoAssignmentByRef isUsedSrch,method_SearchClause_isUsedSrch(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"var_where", var_where
		setArrElement out,"tName", tName
		setArrElement out,"searchFieldsArr", searchFieldsArr
		setArrElement out,"googleLikeFields", googleLikeFields
		setArrElement out,"srchType", srchType
		setArrElement out,"sessionPrefix", sessionPrefix
		setArrElement out,"bIsUsedSrch", bIsUsedSrch
		setArrElement out,"panelSearchFields", panelSearchFields
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment var_where, dict("var_where")
		DoAssignment tName, dict("tName")
		DoAssignment searchFieldsArr, dict("searchFieldsArr")
		DoAssignment googleLikeFields, dict("googleLikeFields")
		DoAssignment srchType, dict("srchType")
		DoAssignment sessionPrefix, dict("sessionPrefix")
		DoAssignment bIsUsedSrch, dict("bIsUsedSrch")
		DoAssignment panelSearchFields, dict("panelSearchFields")
	End Sub
' end serialize
End Class
' SearchClause implementation 
Function method_SearchClause_SearchClause(byref this_object,ByRef params)
	doClassAssignment this_object,"var_where",CreateDictionary()
	this_object.tName = ""
	doClassAssignment this_object,"searchFieldsArr",CreateDictionary()
	doClassAssignment this_object,"googleLikeFields",CreateDictionary()
	this_object.srchType = "integrated"
	this_object.sessionPrefix = ""
	this_object.bIsUsedSrch = false
	doClassAssignment this_object,"panelSearchFields",CreateDictionary()
	doClassAssignment this_object,"tName",IIF(ArrayElement(params,"tName"),ArrayElement(params,"tName"),strTableName)
	doClassAssignment this_object,"sessionPrefix",IIF(ArrayElement(params,"sessionPrefix"),ArrayElement(params,"sessionPrefix"),this_object.tName)
	doClassAssignment this_object,"searchFieldsArr",ArrayElement(params,"searchFieldsArr")
	doClassAssignment this_object,"panelSearchFields",IIF(ArrayElement(params,"panelSearchFields"),ArrayElement(params,"panelSearchFields"),GetTableData(this_object.tName,".panelSearchFields",CreateDictionary()))
	doClassAssignment this_object,"googleLikeFields",IIF(ArrayElement(params,"googleLikeFields"),ArrayElement(params,"googleLikeFields"),GetTableData(this_object.tName,".googleLikeFields",CreateDictionary()))
End Function
Function method_SearchClause_buildAdvancedWhere(byref this_object)
	Dim sWhere,strSearchFor,sfor,strSearchFor2,var_type,f,strSearchOption,where
	sWhere = ""
	if not IsEmpty(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_asearchfor")) then
		GetCollectionBounds ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_asearchfor"),c_searchclause_loopIdx2,c_searchclause_loopMax2
		do while c_searchclause_loopIdx2<=c_searchclause_loopMax2
			f = GetCollectionKey(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_asearchfor"),c_searchclause_loopIdx2)
			doAssignment sfor,ArrayElement(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_asearchfor"),f)
			doAssignmentByRef strSearchFor,trim(sfor)
			strSearchFor2 = ""
			doAssignment var_type,ArrayElement(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_asearchfortype"),f)
			if bValue(asp_array_key_exists(f,ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_asearchfor2"))) then
				doAssignmentByRef strSearchFor2,trim(ArrayElement(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_asearchfor2"),f))
			end if
			if not IsEqual(strSearchFor,"") or bValue(true) then
				if not bValue(sWhere) then
					if IsEqual(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_asearchtype"),"and") then
						sWhere = "1=1"
					else
						sWhere = "1=0"
					end if
				end if
				doAssignmentByRef strSearchOption,trim(ArrayElement(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_asearchopt"),f))
				if bValue(doAssignmentByRef(where,StrWhereAdv(f,strSearchFor,strSearchOption,strSearchFor2,var_type,false))) then
					if bValue(ArrayElement(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_asearchnot"),f)) then
						where = ("not (" & CSmartStr(where)) & ")"
					end if
					if IsEqual(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_asearchtype"),"and") then
						sWhere = CSmartStr(sWhere) & (" and " & CSmartStr(where))
					else
						sWhere = CSmartStr(sWhere) & (" or " & CSmartStr(where))
					end if
				end if
			end if
			c_searchclause_loopIdx2=c_searchclause_loopIdx2+1
		loop
	end if
	doAssignmentByRef method_SearchClause_buildAdvancedWhere,sWhere
	Exit Function
End Function
Function method_SearchClause_buildSimpleWhere(byref this_object)
	Dim sWhere,strSearchFor,strSearchOption,strSearchField,where,strWhere,i
	sWhere = ""
	doAssignmentByRef strSearchFor,trim(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_searchfor"))
	doAssignmentByRef strSearchOption,trim(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_searchoption"))
	if bValue(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_searchfield")) then
		doAssignment strSearchField,ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_searchfield")
		if bValue(doAssignmentByRef(where,StrWhereExpression(strSearchField,strSearchFor,strSearchOption,""))) then
			doAssignmentByRef sWhere,whereAdd(sWhere,where)
		else
			doAssignmentByRef sWhere,whereAdd(sWhere,"1=0")
		end if
	else
		strWhere = "1=0"
		i = 0
		do while IsLess(i,asp_count(this_object.searchFieldsArr))
			if bValue(doAssignmentByRef(where,StrWhereExpression(ArrayElement(this_object.searchFieldsArr,i),strSearchFor,strSearchOption,""))) then
				strWhere = CSmartStr(strWhere) & (" or " & CSmartStr(where))
			end if
			i = CSmartDbl(i)+1
		loop
		doAssignmentByRef sWhere,whereAdd(sWhere,strWhere)
	end if
	doAssignmentByRef method_SearchClause_buildSimpleWhere,sWhere
	Exit Function
End Function
Function method_SearchClause_buildItegratedWhere(byref this_object,ByVal fieldsArr)
	Dim simpleSrch,srchType,srchFields,sWhere,where,i,resWhere,prevSrchFieldName,srchF
	if not bValue(asp_count(fieldsArr)) then
		method_SearchClause_buildItegratedWhere = ""
		Exit Function
	end if
	doAssignment simpleSrch,ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_simpleSrch")
	if IsIdentical(trim(simpleSrch),"%") then
		simpleSrch = ("[" & CSmartStr(simpleSrch)) & "]"
	end if
	doAssignment srchType,ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_srchType")
	doAssignmentByRef srchFields,ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_srchFields")
	sWhere = ""
	if bValue(asp_strlen(simpleSrch)) or IsEqual(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "simpleSrchTypeComboOpt"),"Empty") then
		if bValue(asp_strlen(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "simpleSrchFieldsComboOpt"))) then
			doAssignmentByRef where,StrWhereExpression(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "simpleSrchFieldsComboOpt"),simpleSrch,ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "simpleSrchTypeComboOpt"),"")
			if bValue(where) and bValue(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "simpleSrchTypeComboNot")) then
				where = ("not (" & CSmartStr(where)) & ")"
			end if
			doAssignment sWhere,where
		else
			sWhere = "1=0"
			i = 0
			do while IsLess(i,asp_count(this_object.searchFieldsArr))
				if bValue(asp_in_array(ArrayElement(this_object.searchFieldsArr,i),fieldsArr,false)) and bValue(asp_in_array(ArrayElement(this_object.searchFieldsArr,i),this_object.googleLikeFields,false)) then
					doAssignmentByRef where,StrWhereExpression(ArrayElement(this_object.searchFieldsArr,i),simpleSrch,ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "simpleSrchTypeComboOpt"),"")
					if bValue(where) and bValue(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "simpleSrchTypeComboNot")) then
						where = ("not (" & CSmartStr(where)) & ")"
					end if
					if bValue(where) then
						sWhere = CSmartStr(sWhere) & (" or " & CSmartStr(where))
					end if
				end if
				i = CSmartDbl(i)+1
			loop
		end if
	end if
	doAssignmentByRef resWhere,whereAdd("",sWhere)
	sWhere = ""
	if bValue(asp_count(srchFields)) then
		doAssignment sWhere,IIF(IsEqual(srchType,"and"),"(1=1","(1=0")
		prevSrchFieldName = ""
		GetCollectionBounds srchFields,c_searchclause_loopIdx5,c_searchclause_loopMax5
		do while c_searchclause_loopIdx5<=c_searchclause_loopMax5
			c_searchclause_arrKey5 = GetCollectionKey(srchFields,c_searchclause_loopIdx5)
			doAssignment srchF,ArrayElement(srchFields,c_searchclause_arrKey5)
			if bValue(asp_in_array(ArrayElement(srchF,"fName"),fieldsArr,false)) then
				doAssignmentByRef where,StrWhereAdv(ArrayElement(srchF,"fName"),ArrayElement(srchF,"value1"),ArrayElement(srchF,"opt"),ArrayElement(srchF,"value2"),ArrayElement(srchF,"eType"),false)
				if bValue(where) then
					if bValue(ArrayElement(srchF,"not")) then
						where = ("not (" & CSmartStr(where)) & ")"
					end if
					if IsEqual(srchType,"and") then
						sWhere = CSmartStr(sWhere) & (CSmartStr(IIF(not IsEqual(prevSrchFieldName,ArrayElement(srchF,"fName")),") and ("," and ")) & CSmartStr(where))
					else
						sWhere = CSmartStr(sWhere) & (" or " & CSmartStr(where))
					end if
				end if
				doAssignment prevSrchFieldName,ArrayElement(srchF,"fName")
			end if
			c_searchclause_loopIdx5=c_searchclause_loopIdx5+1
		loop
		sWhere = CSmartStr(sWhere) & ")"
	end if
	doAssignmentByRef resWhere,whereAdd(resWhere,sWhere)
	doAssignmentByRef method_SearchClause_buildItegratedWhere,resWhere
	Exit Function
End Function
Function method_SearchClause_getWhere(byref this_object,ByVal fieldsArr)
	Dim sWhere,this
	sWhere = ""
	do
		If IsEqual(this_object.srchType,"showall") then
			sWhere = ""
			exit do
		End If
		If IsEqual(this_object.srchType,"showall") or IsEqual(this_object.srchType,"advanced") then
			doAssignmentByRef sWhere,this_object.buildAdvancedWhere()
			exit do
		End If
		If IsEqual(this_object.srchType,"showall") or IsEqual(this_object.srchType,"advanced") or IsEqual(this_object.srchType,"simple") then
			doAssignmentByRef sWhere,this_object.buildSimpleWhere()
			exit do
		End If
		If IsEqual(this_object.srchType,"showall") or IsEqual(this_object.srchType,"advanced") or IsEqual(this_object.srchType,"simple") or IsEqual(this_object.srchType,"integrated") then
			doAssignmentByRef sWhere,this_object.buildItegratedWhere_p1(fieldsArr)
			exit do
		End If
		sWhere = ""
	Loop While false
	doAssignmentByRef method_SearchClause_getWhere,sWhere
	Exit Function
End Function
Function method_SearchClause_getSuggestWhere(byref this_object,ByVal fName,ByVal fType,ByVal suggestAllContent,ByVal searchVal)
	Dim sWhere,searchOpt,where
	sWhere = ""
	doAssignment searchOpt,IIF(suggestAllContent,"Contains","Equal")
	doAssignmentByRef where,StrWhereAdv(fName,searchVal,searchOpt,"",fType,true)
	doAssignmentByRef method_SearchClause_getSuggestWhere,where
	Exit Function
End Function
Function method_SearchClause_parseAdvancedRequest(byref this_object)
	Dim additionalControlId,tosearch,asearchfield,gfield,field,asopt,value1,var_type,value2,var_not
	additionalControlId = 1
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_asearchnot",CreateDictionary()
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_asearchopt",CreateDictionary()
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_asearchfor",CreateDictionary()
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_asearchfor2",CreateDictionary()
	tosearch = 0
	doAssignmentByRef asearchfield,postvalue("asearchfield")
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_asearchtype",postvalue("type")
	if not bValue(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_asearchtype")) then
		setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_asearchtype","and"
	end if
	if not IsEmpty(asearchfield) and bValue(asp_is_array(asearchfield)) then
		GetCollectionBounds asearchfield,c_searchclause_loopIdx6,c_searchclause_loopMax6
		do while c_searchclause_loopIdx6<=c_searchclause_loopMax6
			c_searchclause_arrKey6 = GetCollectionKey(asearchfield,c_searchclause_loopIdx6)
			doAssignment field,ArrayElement(asearchfield,c_searchclause_arrKey6)
			doAssignmentByRef gfield,GoodFieldName(field)
			doAssignmentByRef asopt,postvalue("asearchopt_" & CSmartStr(gfield))
			doAssignmentByRef value1,postvalue("value_" & CSmartStr(gfield))
			doAssignmentByRef var_type,postvalue("type_" & CSmartStr(gfield))
			doAssignmentByRef value2,postvalue("value1_" & CSmartStr(gfield))
			doAssignmentByRef var_not,postvalue("not_" & CSmartStr(gfield))
			if bValue(value1) or IsEqual(asopt,"Empty") then
				tosearch = 1
				setArrElementN this_object.var_where,CreateArray2(CSmartStr(this_object.sessionPrefix) & "_asearchopt",field),asopt
				if not bValue(asp_is_array(value1)) then
					setArrElementN this_object.var_where,CreateArray2(CSmartStr(this_object.sessionPrefix) & "_asearchfor",field),value1
				else
					setArrElementN this_object.var_where,CreateArray2(CSmartStr(this_object.sessionPrefix) & "_asearchfor",field),value1
				end if
				setArrElementN this_object.var_where,CreateArray2(CSmartStr(this_object.sessionPrefix) & "_asearchfortype",field),var_type
				if bValue(value2) then
					setArrElementN this_object.var_where,CreateArray2(CSmartStr(this_object.sessionPrefix) & "_asearchfor2",field),value2
				end if
				setArrElementN this_object.var_where,CreateArray2(CSmartStr(this_object.sessionPrefix) & "_asearchnot",field),IsEqual(var_not,"on")
			end if
			c_searchclause_loopIdx6=c_searchclause_loopIdx6+1
		loop
	end if
	if bValue(tosearch) then
		setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_search",2
	else
		setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_search",0
	end if
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_pagenumber",1
End Function
Function method_SearchClause_parseSimpleRequest(byref this_object)
	if IsEqual(postvalue("SearchField"),"") or IsIdentical(asp_in_array(postvalue("SearchField"),this_object.searchFieldsArr,false),true) then
		setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_searchfield",postvalue("SearchField")
		setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_searchoption",postvalue("SearchOption")
		setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_searchfor",postvalue("SearchFor")
		if not IsEqual(postvalue("SearchFor"),"") or IsEqual(postvalue("SearchOption"),"Empty") then
			setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_search",1
		else
			setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_search",0
		end if
		setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_pagenumber",1
	else
		setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_search",0
	end if
End Function
Function method_SearchClause_parseItegratedRequest(byref this_object)
	Dim srchType,j,fName,srchF
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_simpleSrch",trim(postvalue("ctlSearchFor"))
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "simpleSrchTypeComboOpt",trim(postvalue("simpleSrchTypeComboOpt"))
	if not bValue(asp_strlen(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "simpleSrchTypeComboOpt"))) then
		setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "simpleSrchTypeComboOpt","Contains"
	end if
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "simpleSrchTypeComboNot",not IsEqual(trim(postvalue("simpleSrchTypeComboNot")),"")
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "simpleSrchFieldsComboOpt",trim(postvalue("simpleSrchFieldsComboOpt"))
	doAssignmentByRef srchType,postvalue("criteria")
	if not bValue(srchType) then
		srchType = "and"
	end if
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_srchType",srchType
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_srchFields",CreateDictionary()
	j = 1
	do while bValue(doAssignmentByRef(fName,postvalue("field" & CSmartStr(j))))
		if bValue(asp_in_array(fName,this_object.searchFieldsArr,false)) then
			Set srchF = (CreateDictionary())
			setArrElement srchF,"fName",trim(fName)
			setArrElement srchF,"eType",trim(postvalue("type" & CSmartStr(j)))
			setArrElement srchF,"value1",trim(postvalue(("value" & CSmartStr(j)) & "1"))
			setArrElement srchF,"opt",IIF(postvalue("option" & CSmartStr(j)),postvalue("option" & CSmartStr(j)),"Contains")
			setArrElement srchF,"value2",trim(postvalue(("value" & CSmartStr(j)) & "2"))
			setArrElement srchF,"not",IsEqual(postvalue("not" & CSmartStr(j)),"on")
			setArrElementN this_object.var_where,CreateArray2(CSmartStr(this_object.sessionPrefix) & "_srchFields",empty),srchF
		end if
		j = CSmartDbl(j)+1
	loop
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_srchOptShowStatus",IsIdentical(postvalue("srchOptShowStatus"),"1")
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_ctrlTypeComboStatus",IsIdentical(postvalue("ctrlTypeComboStatus"),"1")
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "srchWinShowStatus",IsIdentical(postvalue("srchWinShowStatus"),"1")
End Function
Function method_SearchClause_parseRequest(byref this_object)
	Dim var_REQUEST,this,var_SESSION
	if IsEqual(GetRequestValue(Request,"a"),"showall") then
		setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_search",0
		this_object.srchType = "showAll"
		this_object.bIsUsedSrch = false
		this_object.clearSearch 
		setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_pagenumber",1
	else
		if IsEqual(GetRequestValue(Request,"a"),"search") then
			this_object.srchType = "simple"
			this_object.bIsUsedSrch = true
			this_object.parseSimpleRequest 
		else
			if IsEqual(GetRequestValue(Request,"a"),"advsearch") then
				this_object.srchType = "advanced"
				this_object.bIsUsedSrch = true
				this_object.parseAdvancedRequest 
			else
				if IsEqual(GetRequestValue(Request,"a"),"integrated") then
					this_object.srchType = "integrated"
					this_object.bIsUsedSrch = true
					this_object.parseItegratedRequest 
					setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_pagenumber",1
				end if
			end if
		end if
	end if
End Function
Function method_SearchClause_clearSearch(byref this_object)
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_simpleSrch",""
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_srchType","and"
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "simpleSrchTypeComboOpt","Contains"
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "simpleSrchTypeComboNot",false
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "simpleSrchFieldsComboOpt",""
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_srchFields",CreateDictionary()
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_srchOptShowStatus",false
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_ctrlTypeComboStatus",false
	setArrElement this_object.var_where,CSmartStr(this_object.sessionPrefix) & "srchWinShowStatus",false
End Function
Function method_SearchClause_applyWhere(byref this_object,ByRef sql,ByVal simpleFieldsArr,ByVal aggFieldsArr)
	Dim searchWhereClause,this,searchHavingClause
	if not bValue(asp_count(simpleFieldsArr)) and not bValue(asp_count(aggFieldsArr)) then
		doAssignmentByRef method_SearchClause_applyWhere,sql
		Exit Function
	end if
	doAssignmentByRef searchWhereClause,this_object.getWhere_p1(simpleFieldsArr)
	doAssignmentByRef searchHavingClause,this_object.getWhere_p1(aggFieldsArr)
	if bValue(searchWhereClause) then
		if bValue(ArrayElement(sql,2)) then
			setArrElement sql,2,("(" & CSmartStr(ArrayElement(sql,2))) & ") AND "
		end if
		setArrElement sql,2,CSmartStr(ArrayElement(sql,2)) & (("(" & CSmartStr(searchWhereClause)) & ") ")
	end if
	if bValue(searchHavingClause) then
		if bValue(ArrayElement(sql,4)) then
			setArrElement sql,4,("(" & CSmartStr(ArrayElement(sql,4))) & ") AND "
		end if
		setArrElement sql,4,CSmartStr(ArrayElement(sql,4)) & (("(" & CSmartStr(searchHavingClause)) & ") ")
	end if
	doAssignmentByRef method_SearchClause_applyWhere,sql
	Exit Function
End Function
Function method_SearchClause_getTable(byref this_object)
	doAssignmentByRef method_SearchClause_getTable,this_object.var_where
	Exit Function
End Function
Function method_SearchClause_getSearchCtrlParams(byref this_object,ByVal fName)
	Dim resArr,srchField
	Set resArr = (CreateDictionary())
	if bValue(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_srchFields")) then
		GetCollectionBounds ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_srchFields"),c_searchclause_loopIdx8,c_searchclause_loopMax8
		do while c_searchclause_loopIdx8<=c_searchclause_loopMax8
			c_searchclause_arrKey8 = GetCollectionKey(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_srchFields"),c_searchclause_loopIdx8)
			doAssignment srchField,ArrayElement(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_srchFields"),c_searchclause_arrKey8)
			if IsEqual(asp_strtolower(ArrayElement(srchField,"fName")),asp_strtolower(fName)) then
				setArrElement resArr,asp_count(resArr),srchField
			end if
			c_searchclause_loopIdx8=c_searchclause_loopIdx8+1
		loop
	end if
	doAssignmentByRef method_SearchClause_getSearchCtrlParams,resArr
	Exit Function
End Function
Function method_SearchClause_getUsedCtrlsCount(byref this_object)
	if bValue(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_srchFields")) then
		doAssignmentByRef method_SearchClause_getUsedCtrlsCount,asp_count(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_srchFields"))
		Exit Function
	else
		method_SearchClause_getUsedCtrlsCount = 0
		Exit Function
	end if
End Function
Function method_SearchClause_getSearchGlobalParams(byref this_object)
	doAssignmentByRef method_SearchClause_getSearchGlobalParams,CreateDictionary6("simpleSrch",ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_simpleSrch"),"srchTypeRadio",ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_srchType"),"srchType",this_object.srchType,"simpleSrchTypeComboOpt",ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "simpleSrchTypeComboOpt"),"simpleSrchTypeComboNot",ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "simpleSrchTypeComboNot"),"simpleSrchFieldsComboOpt",ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "simpleSrchFieldsComboOpt"))
	Exit Function
End Function
Function method_SearchClause_getSrchPanelAttrs(byref this_object)
	doAssignmentByRef method_SearchClause_getSrchPanelAttrs,CreateDictionary3("srchOptShowStatus",bValue(ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_srchOptShowStatus")) or bValue(asp_count(this_object.panelSearchFields)),"ctrlTypeComboStatus",ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "_ctrlTypeComboStatus"),"srchWinShowStatus",ArrayElement(this_object.var_where,CSmartStr(this_object.sessionPrefix) & "srchWinShowStatus"))
	Exit Function
End Function
Function method_SearchClause_isUsedSrch(byref this_object)
	doAssignmentByRef method_SearchClause_isUsedSrch,this_object.bIsUsedSrch
	Exit Function
End Function
%>
