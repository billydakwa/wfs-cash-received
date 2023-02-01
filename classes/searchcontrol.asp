<%

'------ Class SearchControl ------
Class SearchControl
	Public tName
	Public globSrchParams
	Public getSrchPanelAttrs
	Public dispNoneStyle
	Public pageObj
	Public searchClauseObj
	Public id
	Public Function init_SearchControl_p4(ByVal id,ByVal tName,ByRef searchClauseObj,ByRef pageObj)
		DoAssignmentByRef init_SearchControl_p4,method_SearchControl_SearchControl(me,id,tName,searchClauseObj,pageObj)
	End Function
	Public Function getCtrlParamsArr_p6(ByVal fName,ByVal recId,ByVal fieldNum,ByVal value,ByVal renderHidden,ByVal isCached)
		DoAssignmentByRef getCtrlParamsArr_p6,method_SearchControl_getCtrlParamsArr(me,fName,recId,fieldNum,value,renderHidden,isCached)
	End Function
	Public Function getCtrlParamsArr_p5(ByVal fName,ByVal recId,ByVal fieldNum,ByVal value,ByVal renderHidden)
		DoAssignmentByRef getCtrlParamsArr_p5,method_SearchControl_getCtrlParamsArr(me,fName,recId,fieldNum,value,renderHidden,true)
	End Function
	Public Function getCtrlParamsArr_p4(ByVal fName,ByVal recId,ByVal fieldNum,ByVal value)
		DoAssignmentByRef getCtrlParamsArr_p4,method_SearchControl_getCtrlParamsArr(me,fName,recId,fieldNum,value,false,true)
	End Function
	Public Function getSecCtrlParamsArr_p6(ByVal fName,ByVal recId,ByVal fieldNum,ByVal value,ByVal renderHidden,ByVal isCached)
		DoAssignmentByRef getSecCtrlParamsArr_p6,method_SearchControl_getSecCtrlParamsArr(me,fName,recId,fieldNum,value,renderHidden,isCached)
	End Function
	Public Function getSecCtrlParamsArr_p5(ByVal fName,ByVal recId,ByVal fieldNum,ByVal value,ByVal renderHidden)
		DoAssignmentByRef getSecCtrlParamsArr_p5,method_SearchControl_getSecCtrlParamsArr(me,fName,recId,fieldNum,value,renderHidden,true)
	End Function
	Public Function getSecCtrlParamsArr_p4(ByVal fName,ByVal recId,ByVal fieldNum,ByVal value)
		DoAssignmentByRef getSecCtrlParamsArr_p4,method_SearchControl_getSecCtrlParamsArr(me,fName,recId,fieldNum,value,false,true)
	End Function
	Public Function isNeedSecondCtrl_p1(ByVal fName)
		DoAssignmentByRef isNeedSecondCtrl_p1,method_SearchControl_isNeedSecondCtrl(me,fName)
	End Function
	Public Function getSimpleSearchTypeCombo_p2(ByVal selOpt,ByVal var_not)
		DoAssignmentByRef getSimpleSearchTypeCombo_p2,method_SearchControl_getSimpleSearchTypeCombo(me,selOpt,var_not)
	End Function
	Public Function getCtrlSearchTypeOptions_p3(ByVal fName,ByVal selOpt,ByVal var_not)
		DoAssignmentByRef getCtrlSearchTypeOptions_p3,method_SearchControl_getCtrlSearchTypeOptions(me,fName,selOpt,var_not)
	End Function
	Public Function getCtrlSearchType_p6(ByVal fName,ByVal recId,ByVal fieldNum,ByVal selOpt,ByVal var_not,ByVal renderHidden)
		DoAssignmentByRef getCtrlSearchType_p6,method_SearchControl_getCtrlSearchType(me,fName,recId,fieldNum,selOpt,var_not,renderHidden)
	End Function
	Public Function getCtrlSearchType_p5(ByVal fName,ByVal recId,ByVal fieldNum,ByVal selOpt,ByVal var_not)
		DoAssignmentByRef getCtrlSearchType_p5,method_SearchControl_getCtrlSearchType(me,fName,recId,fieldNum,selOpt,var_not,false)
	End Function
	Public Function getSearchOptionId_p2(ByVal fName,ByVal recId)
		DoAssignmentByRef getSearchOptionId_p2,method_SearchControl_getSearchOptionId(me,fName,recId)
	End Function
	Public Function getNotBox_p3(ByVal fName,ByVal recId,ByVal var_not)
		DoAssignmentByRef getNotBox_p3,method_SearchControl_getNotBox(me,fName,recId,var_not)
	End Function
	Public Function getDelButtonHtml_p2(ByVal fName,ByVal recId)
		DoAssignmentByRef getDelButtonHtml_p2,method_SearchControl_getDelButtonHtml(me,fName,recId)
	End Function
	Public Function getDelButtonId_p2(ByVal fName,ByVal recId)
		DoAssignmentByRef getDelButtonId_p2,method_SearchControl_getDelButtonId(me,fName,recId)
	End Function
	Public Function getSearchRadio()
		DoAssignmentByRef getSearchRadio,method_SearchControl_getSearchRadio(me)
	End Function
	Public Function getFilterDivId_p2(ByVal recId,ByVal fName)
		DoAssignmentByRef getFilterDivId_p2,method_SearchControl_getFilterDivId(me,recId,fName)
	End Function
	Public Function getCtrlComboContId_p2(ByVal recId,ByVal fName)
		DoAssignmentByRef getCtrlComboContId_p2,method_SearchControl_getCtrlComboContId(me,recId,fName)
	End Function
	Public Function buildSearchCtrlBlockArr_p8(ByVal recId,ByVal fName,ByVal ctrlInd,ByVal opt,ByVal var_not,ByVal isChached,ByVal val1,ByVal val2)
		DoAssignmentByRef buildSearchCtrlBlockArr_p8,method_SearchControl_buildSearchCtrlBlockArr(me,recId,fName,ctrlInd,opt,var_not,isChached,val1,val2)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"tName", tName
		setArrElement out,"globSrchParams", globSrchParams
		setArrElement out,"getSrchPanelAttrs", getSrchPanelAttrs
		setArrElement out,"dispNoneStyle", dispNoneStyle
		setArrElement out,"pageObj", pageObj
		setArrElement out,"searchClauseObj", searchClauseObj
		setArrElement out,"id", id
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment tName, dict("tName")
		DoAssignment globSrchParams, dict("globSrchParams")
		DoAssignment getSrchPanelAttrs, dict("getSrchPanelAttrs")
		DoAssignment dispNoneStyle, dict("dispNoneStyle")
		DoAssignment pageObj, dict("pageObj")
		DoAssignment searchClauseObj, dict("searchClauseObj")
		DoAssignment id, dict("id")
	End Sub
' end serialize
End Class
' SearchControl implementation 
Function method_SearchControl_SearchControl(byref this_object,ByVal id,ByVal tName,ByRef searchClauseObj,ByRef pageObj)
	this_object.tName = ""
	doClassAssignment this_object,"globSrchParams",CreateDictionary()
	doClassAssignment this_object,"getSrchPanelAttrs",CreateDictionary()
	this_object.dispNoneStyle = "style=""display: none;"""
	this_object.pageObj = null
	this_object.searchClauseObj = false
	this_object.id = 1
	doClassAssignment this_object,"tName",tName
	doClassAssignment this_object,"searchClauseObj",searchClauseObj
	doClassAssignment this_object,"getSrchPanelAttrs",this_object.searchClauseObj.getSrchPanelAttrs()
	doClassAssignment this_object,"globSrchParams",this_object.searchClauseObj.getSearchGlobalParams()
	doClassAssignment this_object,"id",id
	doClassAssignmentByRef this_object,"pageObj",pageObj
End Function
Function method_SearchControl_getCtrlParamsArr(byref this_object,ByVal fName,ByVal recId,ByVal fieldNum,ByVal value,ByVal renderHidden,ByVal isCached)
	Dim fType,format,control,ctrlsMap,vals,preload,additionalCtrlParams
	doAssignmentByRef fType,GetEditFormat(fName,this_object.tName)
	if (((IsEqual(fType,EDIT_FORMAT_TEXT_AREA) or IsEqual(fType,EDIT_FORMAT_PASSWORD)) or IsEqual(fType,EDIT_FORMAT_HIDDEN)) or IsEqual(fType,EDIT_FORMAT_READONLY)) or IsEqual(fType,EDIT_FORMAT_FILE) then
		format = EDIT_FORMAT_TEXT_FIELD
	else
		doAssignment format,fType
	end if
	Set control = (CreateDictionary())
	setArrElement control,"params",CreateDictionary()
	setArrElement control,"func","xt_buildeditcontrol"
	setArrElementN control,CreateArray2("params","field"),fName
	setArrElementN control,CreateArray2("params","mode"),"search"
	setArrElementN control,CreateArray2("params","id"),recId
	setArrElementN control,CreateArray2("params","fieldNum"),fieldNum
	setArrElementN control,CreateArray2("params","format"),format
	setArrElementN control,CreateArray2("params","pageObj"),this_object.pageObj
	Set ctrlsMap = (CreateDictionary1("controls",CreateDictionary()))
	setArrElementN ctrlsMap,CreateArray2("controls","fieldName"),fName
	setArrElementN ctrlsMap,CreateArray2("controls","mode"),MODE_SEARCH
	setArrElementN ctrlsMap,CreateArray2("controls","editFormat"),format
	setArrElementN ctrlsMap,CreateArray2("controls","id"),recId
	setArrElementN ctrlsMap,CreateArray2("controls","ctrlInd"),fieldNum
	setArrElementN ctrlsMap,CreateArray2("controls","hidden"),bValue(renderHidden) or bValue(isCached)
	setArrElementN ctrlsMap,CreateArray2("controls","table"),this_object.tName
	Set vals = (CreateDictionary1(fName,value))
	doAssignmentByRef preload,this_object.pageObj.fillPreload_p2(fName,vals)
	if not IsFalse(preload) then
		setArrElementN ctrlsMap,CreateArray2("controls","preloadData"),preload
	end if
	this_object.pageObj.fillControlsMap_p1 ctrlsMap
	Set additionalCtrlParams = (CreateDictionary())
	setArrElement additionalCtrlParams,"hidden",bValue(renderHidden) or bValue(isCached)
	setArrElementN control,CreateArray2("params","additionalCtrlParams"),additionalCtrlParams
	setArrElementN control,CreateArray2("params","value"),value
	doAssignmentByRef method_SearchControl_getCtrlParamsArr,control
	Exit Function
End Function
Function method_SearchControl_getSecCtrlParamsArr(byref this_object,ByVal fName,ByVal recId,ByVal fieldNum,ByVal value,ByVal renderHidden,ByVal isCached)
	Dim fType,this
	doAssignmentByRef fType,GetEditFormat(fName,this_object.tName)
	if bValue(this_object.isNeedSecondCtrl_p1(fName)) then
		doAssignmentByRef method_SearchControl_getSecCtrlParamsArr,this_object.getCtrlParamsArr_p6(fName,recId,CSmartDbl(fieldNum)+1,value,renderHidden,isCached)
		Exit Function
	else
		method_SearchControl_getSecCtrlParamsArr = false
		Exit Function
	end if
End Function
Function method_SearchControl_isNeedSecondCtrl(byref this_object,ByVal fName)
	Dim fType
	doAssignmentByRef fType,GetEditFormat(fName,this_object.tName)
	if (((((IsEqual(fType,EDIT_FORMAT_DATE) or IsEqual(fType,EDIT_FORMAT_TIME)) or IsEqual(fType,EDIT_FORMAT_TEXT_FIELD)) or IsEqual(fType,EDIT_FORMAT_TEXT_AREA)) or IsEqual(fType,EDIT_FORMAT_PASSWORD)) or IsEqual(fType,EDIT_FORMAT_HIDDEN)) or IsEqual(fType,EDIT_FORMAT_READONLY) then
		method_SearchControl_isNeedSecondCtrl = true
		Exit Function
	else
		method_SearchControl_isNeedSecondCtrl = false
		Exit Function
	end if
End Function
Function method_SearchControl_getSimpleSearchTypeCombo(byref this_object,ByVal selOpt,ByVal var_not)
	Dim options
	options = ""
	options = CSmartStr(options) & (((("<OPTION VALUE=""Contains"" " & CSmartStr(IIF(IsEqual(selOpt,"Contains") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("Contains")) & "</option>")
	options = CSmartStr(options) & (((("<OPTION VALUE=""Equals"" " & CSmartStr(IIF(IsEqual(selOpt,"Equals") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("Equals")) & "</option>")
	options = CSmartStr(options) & (((("<OPTION VALUE=""Starts with"" " & CSmartStr(IIF(IsEqual(selOpt,"Starts with") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("Starts with")) & "</option>")
	options = CSmartStr(options) & (((("<OPTION VALUE=""More than"" " & CSmartStr(IIF(IsEqual(selOpt,"More than") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("More than")) & "</option>")
	options = CSmartStr(options) & (((("<OPTION VALUE=""Less than"" " & CSmartStr(IIF(IsEqual(selOpt,"Less than") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("Less than")) & "</option>")
	options = CSmartStr(options) & (((("<OPTION VALUE=""Empty"" " & CSmartStr(IIF(IsEqual(selOpt,"Empty") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("Empty")) & "</option>")
	doAssignmentByRef method_SearchControl_getSimpleSearchTypeCombo,options
	Exit Function
End Function
Function method_SearchControl_getCtrlSearchTypeOptions(byref this_object,ByVal fName,ByVal selOpt,ByVal var_not)
	Dim fType,options
	if bValue(asp_strlen(fName)) then
		doAssignmentByRef fType,GetEditFormat(fName,this_object.tName)
	else
		fType = EDIT_FORMAT_TEXT_FIELD
	end if
	options = ""
	if IsEqual(fType,EDIT_FORMAT_DATE) or IsEqual(fType,EDIT_FORMAT_TIME) then
		options = CSmartStr(options) & (((("<OPTION VALUE=""Equals"" " & CSmartStr(IIF(IsEqual(selOpt,"Equals") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("Equals")) & "</option>")
		options = CSmartStr(options) & (((("<OPTION VALUE=""More than"" " & CSmartStr(IIF(IsEqual(selOpt,"More than") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("More than")) & "</option>")
		options = CSmartStr(options) & (((("<OPTION VALUE=""Less than"" " & CSmartStr(IIF(IsEqual(selOpt,"Less than") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("Less than")) & "</option>")
		options = CSmartStr(options) & (((("<OPTION VALUE=""Between"" " & CSmartStr(IIF(IsEqual(selOpt,"Between") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("Between")) & "</option>")
		options = CSmartStr(options) & (((("<OPTION VALUE=""Empty"" " & CSmartStr(IIF(IsEqual(selOpt,"Empty") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("Empty")) & "</option>")
	else
		if IsEqual(fType,EDIT_FORMAT_LOOKUP_WIZARD) then
			if bValue(Multiselect(fName,this_object.tName)) then
				options = CSmartStr(options) & (((("<OPTION VALUE=""Contains"" " & CSmartStr(IIF(IsEqual(selOpt,"Contains") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("Contains")) & "</option>")
			else
				options = CSmartStr(options) & (((("<OPTION VALUE=""Equals"" " & CSmartStr(IIF(IsEqual(selOpt,"Equals") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("Equals")) & "</option>")
			end if
		else
			if (((IsEqual(fType,EDIT_FORMAT_TEXT_FIELD) or IsEqual(fType,EDIT_FORMAT_TEXT_AREA)) or IsEqual(fType,EDIT_FORMAT_PASSWORD)) or IsEqual(fType,EDIT_FORMAT_HIDDEN)) or IsEqual(fType,EDIT_FORMAT_READONLY) then
				options = CSmartStr(options) & (((("<OPTION VALUE=""Contains"" " & CSmartStr(IIF(IsEqual(selOpt,"Contains") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("Contains")) & "</option>")
				options = CSmartStr(options) & (((("<OPTION VALUE=""Equals"" " & CSmartStr(IIF(IsEqual(selOpt,"Equals") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("Equals")) & "</option>")
				options = CSmartStr(options) & (((("<OPTION VALUE=""Starts with"" " & CSmartStr(IIF(IsEqual(selOpt,"Starts with") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("Starts with")) & "</option>")
				options = CSmartStr(options) & (((("<OPTION VALUE=""More than"" " & CSmartStr(IIF(IsEqual(selOpt,"More than") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("More than")) & "</option>")
				options = CSmartStr(options) & (((("<OPTION VALUE=""Less than"" " & CSmartStr(IIF(IsEqual(selOpt,"Less than") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("Less than")) & "</option>")
				options = CSmartStr(options) & (((("<OPTION VALUE=""Between"" " & CSmartStr(IIF(IsEqual(selOpt,"Between") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("Between")) & "</option>")
				options = CSmartStr(options) & (((("<OPTION VALUE=""Empty"" " & CSmartStr(IIF(IsEqual(selOpt,"Empty") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("Empty")) & "</option>")
			else
				options = CSmartStr(options) & (((("<OPTION VALUE=""Equals"" " & CSmartStr(IIF(IsEqual(selOpt,"Equals") and not bValue(var_not),"selected",""))) & ">") & CSmartStr("Equals")) & "</option>")
			end if
		end if
	end if
	doAssignmentByRef method_SearchControl_getCtrlSearchTypeOptions,options
	Exit Function
End Function
Function method_SearchControl_getCtrlSearchType(byref this_object,ByVal fName,ByVal recId,ByVal fieldNum,ByVal selOpt,ByVal var_not,ByVal renderHidden)
	Dim searchtype,this
	searchtype = ((((("<SELECT id=""" & CSmartStr(this_object.getSearchOptionId_p2(fName,recId))) & """ NAME=""") & CSmartStr(this_object.getSearchOptionId_p2(fName,recId))) & """ SIZE=1 ") & CSmartStr(IIF(bValue(renderHidden) or not bValue(ArrayElement(this_object.getSrchPanelAttrs,"ctrlTypeComboStatus")),"style=""display: none;""",""))) & ">"
	searchtype = CSmartStr(searchtype) & CSmartStr(this_object.getCtrlSearchTypeOptions_p3(fName,selOpt,var_not))
	searchtype = CSmartStr(searchtype) & "</SELECT>"
	doAssignmentByRef method_SearchControl_getCtrlSearchType,searchtype
	Exit Function
End Function
Function method_SearchControl_getSearchOptionId(byref this_object,ByVal fName,ByVal recId)
	method_SearchControl_getSearchOptionId = (("srchOpt_" & CSmartStr(recId)) & "_") & CSmartStr(GoodFieldName(fName))
	Exit Function
End Function
Function method_SearchControl_getNotBox(byref this_object,ByVal fName,ByVal recId,ByVal var_not)
	Dim notbox
	notbox = ((("id=""not_" & CSmartStr(recId)) & "_") & CSmartStr(GoodFieldName(fName))) & """"
	if bValue(var_not) then
		notbox = CSmartStr(notbox) & " checked"
	end if
	doAssignmentByRef method_SearchControl_getNotBox,notbox
	Exit Function
End Function
Function method_SearchControl_getDelButtonHtml(byref this_object,ByVal fName,ByVal recId)
	Dim html,this
	html = ((((((("<img style=""visibility: hidden;"" id = """ & CSmartStr(this_object.getDelButtonId_p2(fName,recId))) & """ ctrlId=""") & CSmartStr(recId)) & """ fName=""") & CSmartStr(fName)) & """ class=""searchPanelButton"" src=""images/search/closeRed.gif"" alt=""") & CSmartStr("Delete control")) & """>"
	doAssignmentByRef method_SearchControl_getDelButtonHtml,html
	Exit Function
End Function
Function method_SearchControl_getDelButtonId(byref this_object,ByVal fName,ByVal recId)
	method_SearchControl_getDelButtonId = (("delCtrlButt_" & CSmartStr(recId)) & "_") & CSmartStr(GoodFieldName(fName))
	Exit Function
End Function
Function method_SearchControl_getSearchRadio(byref this_object)
	Dim resArr,id508l,id508n
	Set resArr = (CreateDictionary())
	setArrElement resArr,"all_checkbox_label",CreateDictionary2(0,"",1,"")
	setArrElement resArr,"any_checkbox_label",CreateDictionary2(0,"",1,"")
	if bValue(isEnableSection508()) then
		setArrElement resArr,"all_checkbox_label",CreateDictionary2(0,"<label for=""all_checkbox"">",1,"</label>")
		setArrElement resArr,"any_checkbox_label",CreateDictionary2(0,"<label for=""any_checkbox"">",1,"</label>")
	end if
	id508l = "id=""all_checkbox"" "
	id508n = "id=""any_checkbox"" "
	setArrElement resArr,"all_checkbox",id508l
	setArrElement resArr,"any_checkbox",id508n
	setArrElement resArr,"all_checkbox",CSmartStr(ArrayElement(resArr,"all_checkbox")) & "value=""and"" "
	setArrElement resArr,"any_checkbox",CSmartStr(ArrayElement(resArr,"any_checkbox")) & "value=""or"" "
	if not IsEmpty(ArrayElement(this_object.globSrchParams,"srchTypeRadio")) and IsEqual(ArrayElement(this_object.globSrchParams,"srchTypeRadio"),"or") then
		setArrElement resArr,"any_checkbox",CSmartStr(ArrayElement(resArr,"any_checkbox")) & " checked"
	else
		setArrElement resArr,"all_checkbox",CSmartStr(ArrayElement(resArr,"all_checkbox")) & " checked"
	end if
	doAssignmentByRef method_SearchControl_getSearchRadio,resArr
	Exit Function
End Function
Function method_SearchControl_getFilterDivId(byref this_object,ByVal recId,ByVal fName)
	method_SearchControl_getFilterDivId = (("filter_" & CSmartStr(recId)) & "_") & CSmartStr(GoodFieldName(fName))
	Exit Function
End Function
Function method_SearchControl_getCtrlComboContId(byref this_object,ByVal recId,ByVal fName)
	method_SearchControl_getCtrlComboContId = (("searchType_" & CSmartStr(recId)) & "_") & CSmartStr(GoodFieldName(fName))
	Exit Function
End Function
Function method_SearchControl_buildSearchCtrlBlockArr(byref this_object,ByVal recId,ByVal fName,ByVal ctrlInd,ByVal opt,ByVal var_not,ByVal isChached,ByVal val1,ByVal val2)
	Dim srchCtrlBlock,this,renderHidden,filterDivId
	Set srchCtrlBlock = (CreateDictionary())
	setArrElement srchCtrlBlock,"searchcontrol",this_object.getCtrlParamsArr_p6(fName,recId,ctrlInd,val1,false,isChached)
	renderHidden = not IsEqual(asp_strtolower(opt),"between") and not IsEqual(asp_strtolower(opt),"not between")
	setArrElement srchCtrlBlock,"searchcontrol1",this_object.getSecCtrlParamsArr_p6(fName,recId,ctrlInd,val2,renderHidden,isChached)
	setArrElement srchCtrlBlock,"secCtrlCont_attrs",""
	setArrElement srchCtrlBlock,"delCtrlButt",this_object.getDelButtonHtml_p2(fName,recId)
	doAssignmentByRef filterDivId,this_object.getFilterDivId_p2(recId,fName)
	setArrElement srchCtrlBlock,"filterDiv_attrs",((CSmartStr(IIF(isChached,this_object.dispNoneStyle,"")) & " id=""") & CSmartStr(filterDivId)) & """ "
	setArrElement srchCtrlBlock,"fName",fName
	setArrElement srchCtrlBlock,"searchtype",this_object.getCtrlSearchType_p5(fName,recId,ctrlInd,opt,var_not)
	setArrElement srchCtrlBlock,"srchTypeCont_attrs",("id=""" & CSmartStr(this_object.getCtrlComboContId_p2(recId,fName))) & """"
	setArrElement srchCtrlBlock,"srchTypeCont_attrs",CSmartStr(ArrayElement(srchCtrlBlock,"srchTypeCont_attrs")) & CSmartStr(IIF(ArrayElement(this_object.getSrchPanelAttrs,"ctrlTypeComboStatus"),"","style=""display: none;"""))
	setArrElement srchCtrlBlock,"notbox",this_object.getNotBox_p3(fName,recId,var_not)
	setArrElement srchCtrlBlock,"fLabel",GetFieldLabel(GoodFieldName(this_object.tName),GoodFieldName(fName))
	doAssignmentByRef method_SearchControl_buildSearchCtrlBlockArr,srchCtrlBlock
	Exit Function
End Function
%>
