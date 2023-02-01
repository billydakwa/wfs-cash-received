<%

'------ Class PanelSearchControl extends SearchControl ------
Class PanelSearchControl
	Public tName
	Public globSrchParams
	Public getSrchPanelAttrs
	Public dispNoneStyle
	Public pageObj
	Public searchClauseObj
	Public id
	Public Function getCtrlParamsArr_p6(ByVal fName,ByVal recId,ByVal fieldNum,ByVal value,ByVal renderHidden,ByVal isCached)
		DoAssignmentByRef getCtrlParamsArr_p6,method_PanelSearchControl_getCtrlParamsArr(me,fName,recId,fieldNum,value,renderHidden,isCached)
	End Function
	Public Function getCtrlParamsArr_p5(ByVal fName,ByVal recId,ByVal fieldNum,ByVal value,ByVal renderHidden)
		DoAssignmentByRef getCtrlParamsArr_p5,method_PanelSearchControl_getCtrlParamsArr(me,fName,recId,fieldNum,value,renderHidden,true)
	End Function
	Public Function getCtrlParamsArr_p4(ByVal fName,ByVal recId,ByVal fieldNum,ByVal value)
		DoAssignmentByRef getCtrlParamsArr_p4,method_PanelSearchControl_getCtrlParamsArr(me,fName,recId,fieldNum,value,false,true)
	End Function
	Public Function simpleSearchFieldCombo_p2(ByVal fNamesArr,ByVal selOpt)
		DoAssignmentByRef simpleSearchFieldCombo_p2,method_PanelSearchControl_simpleSearchFieldCombo(me,fNamesArr,selOpt)
	End Function
	Public Function getCtrlSearchTypeOptions_p3(ByVal fName,ByVal selOpt,ByVal var_not)
		DoAssignmentByRef getCtrlSearchTypeOptions_p3,method_PanelSearchControl_getCtrlSearchTypeOptions(me,fName,selOpt,var_not)
	End Function
	Public Function buildSearchCtrlWinBlockArr_p2(ByVal recId,ByVal fName)
		DoAssignmentByRef buildSearchCtrlWinBlockArr_p2,method_PanelSearchControl_buildSearchCtrlWinBlockArr(me,recId,fName)
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
	Public Function init_PanelSearchControl_p4(ByVal id,ByVal tName,ByRef searchClauseObj,ByRef pageObj)
		DoAssignmentByRef init_PanelSearchControl_p4,method_PanelSearchControl_PanelSearchControl(me,id,tName,searchClauseObj,pageObj)
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
' PanelSearchControl implementation 
Function method_PanelSearchControl_getCtrlParamsArr(byref this_object,ByVal fName,ByVal recId,ByVal fieldNum,ByVal value,ByVal renderHidden,ByVal isCached)
	Dim control,ctrlsMap
	doAssignmentByRef control,method_SearchControl_getCtrlParamsArr(this_object,fName,recId,fieldNum,value,renderHidden,isCached)
	setArrElementN control,CreateArray3("params","additionalCtrlParams","skipDependencies"),true
	setArrElementN control,CreateArray3("params","additionalCtrlParams","style"),"width: 115px;"
	Set ctrlsMap = (CreateDictionary1("controls",CreateDictionary()))
	setArrElementN ctrlsMap,CreateArray2("controls","skipDependencies"),true
	setArrElementN ctrlsMap,CreateArray2("controls","style"),"width: 115px;"
	this_object.pageObj.fillControlsMap_p2 ctrlsMap,true
	doAssignmentByRef method_PanelSearchControl_getCtrlParamsArr,control
	Exit Function
End Function
Function method_PanelSearchControl_simpleSearchFieldCombo(byref this_object,ByVal fNamesArr,ByVal selOpt)
	Dim options,fLabel,fName
	options = ("<OPTION VALUE="""" >" & CSmartStr("Any field")) & "</option>"
	GetCollectionBounds fNamesArr,c_panelsearchcontrol_loopIdx2,c_panelsearchcontrol_loopMax2
	do while c_panelsearchcontrol_loopIdx2<=c_panelsearchcontrol_loopMax2
		c_panelsearchcontrol_arrKey2 = GetCollectionKey(fNamesArr,c_panelsearchcontrol_loopIdx2)
		doAssignment fName,ArrayElement(fNamesArr,c_panelsearchcontrol_arrKey2)
		doAssignmentByRef fLabel,GetFieldLabel(GoodFieldName(this_object.tName),GoodFieldName(fName))
		options = CSmartStr(options) & (((((("<OPTION VALUE=""" & CSmartStr(fName)) & """ ") & CSmartStr(IIF(IsEqual(selOpt,fName),"selected",""))) & ">") & CSmartStr(fLabel)) & "</option>")
		c_panelsearchcontrol_loopIdx2=c_panelsearchcontrol_loopIdx2+1
	loop
	doAssignmentByRef method_PanelSearchControl_simpleSearchFieldCombo,options
	Exit Function
End Function
Function method_PanelSearchControl_getCtrlSearchTypeOptions(byref this_object,ByVal fName,ByVal selOpt,ByVal var_not)
	Dim options,fType
	doAssignmentByRef options,method_SearchControl_getCtrlSearchTypeOptions(this_object,fName,selOpt,var_not)
	if bValue(asp_strlen(fName)) then
		doAssignmentByRef fType,GetEditFormat(fName,this_object.tName)
	else
		fType = EDIT_FORMAT_TEXT_FIELD
	end if
	if IsEqual(fType,EDIT_FORMAT_DATE) or IsEqual(fType,EDIT_FORMAT_TIME) then
		options = CSmartStr(options) & (((("<OPTION VALUE=""NOT Equals"" " & CSmartStr(IIF(IsEqual(selOpt,"Equals") and bValue(var_not),"selected",""))) & ">") & CSmartStr("Doesn't equal")) & "</option>")
		options = CSmartStr(options) & (((("<OPTION VALUE=""NOT More than"" " & CSmartStr(IIF(IsEqual(selOpt,"More than") and bValue(var_not),"selected",""))) & ">") & CSmartStr("Is not more than")) & "</option>")
		options = CSmartStr(options) & (((("<OPTION VALUE=""NOT Less than"" " & CSmartStr(IIF(IsEqual(selOpt,"Less than") and bValue(var_not),"selected",""))) & ">") & CSmartStr("Is not less than")) & "</option>")
		options = CSmartStr(options) & (((("<OPTION VALUE=""NOT Between"" " & CSmartStr(IIF(IsEqual(selOpt,"Between") and bValue(var_not),"selected",""))) & ">") & CSmartStr("Is not between")) & "</option>")
		options = CSmartStr(options) & (((("<OPTION VALUE=""NOT Empty"" " & CSmartStr(IIF(IsEqual(selOpt,"Empty") and bValue(var_not),"selected",""))) & ">") & CSmartStr("Is not empty")) & "</option>")
	else
		if IsEqual(fType,EDIT_FORMAT_LOOKUP_WIZARD) then
			if bValue(Multiselect(fName,this_object.tName)) then
				options = CSmartStr(options) & (((("<OPTION VALUE=""NOT Contains"" " & CSmartStr(IIF(IsEqual(selOpt,"Contains") and bValue(var_not),"selected",""))) & ">") & CSmartStr("Doesn't contain")) & "</option>")
			else
				options = CSmartStr(options) & (((("<OPTION VALUE=""NOT Equals"" " & CSmartStr(IIF(IsEqual(selOpt,"Equals") and bValue(var_not),"selected",""))) & ">") & CSmartStr("Doesn't equal")) & "</option>")
			end if
		else
			if (((IsEqual(fType,EDIT_FORMAT_TEXT_FIELD) or IsEqual(fType,EDIT_FORMAT_TEXT_AREA)) or IsEqual(fType,EDIT_FORMAT_PASSWORD)) or IsEqual(fType,EDIT_FORMAT_HIDDEN)) or IsEqual(fType,EDIT_FORMAT_READONLY) then
				options = CSmartStr(options) & (((("<OPTION VALUE=""NOT Contains"" " & CSmartStr(IIF(IsEqual(selOpt,"Contains") and bValue(var_not),"selected",""))) & ">") & CSmartStr("Doesn't contain")) & "</option>")
				options = CSmartStr(options) & (((("<OPTION VALUE=""NOT Equals"" " & CSmartStr(IIF(IsEqual(selOpt,"Equals") and bValue(var_not),"selected",""))) & ">") & CSmartStr("Doesn't equal")) & "</option>")
				options = CSmartStr(options) & (((("<OPTION VALUE=""NOT Starts with"" " & CSmartStr(IIF(IsEqual(selOpt,"Starts with") and bValue(var_not),"selected",""))) & ">") & CSmartStr("Doesn't start with")) & "</option>")
				options = CSmartStr(options) & (((("<OPTION VALUE=""NOT More than"" " & CSmartStr(IIF(IsEqual(selOpt,"More than") and bValue(var_not),"selected",""))) & ">") & CSmartStr("Is not more than")) & "</option>")
				options = CSmartStr(options) & (((("<OPTION VALUE=""NOT Less than"" " & CSmartStr(IIF(IsEqual(selOpt,"Less than") and bValue(var_not),"selected",""))) & ">") & CSmartStr("Is not less than")) & "</option>")
				options = CSmartStr(options) & (((("<OPTION VALUE=""NOT Between"" " & CSmartStr(IIF(IsEqual(selOpt,"Between") and bValue(var_not),"selected",""))) & ">") & CSmartStr("Is not between")) & "</option>")
				options = CSmartStr(options) & (((("<OPTION VALUE=""NOT Empty"" " & CSmartStr(IIF(IsEqual(selOpt,"Empty") and bValue(var_not),"selected",""))) & ">") & CSmartStr("Is not empty")) & "</option>")
			else
				options = CSmartStr(options) & (((("<OPTION VALUE=""NOT Equals"" " & CSmartStr(IIF(IsEqual(selOpt,"Equals") and bValue(var_not),"selected",""))) & ">") & CSmartStr("Doesn't equal")) & "</option>")
			end if
		end if
	end if
	doAssignmentByRef method_PanelSearchControl_getCtrlSearchTypeOptions,options
	Exit Function
End Function
Function method_PanelSearchControl_buildSearchCtrlWinBlockArr(byref this_object,ByVal recId,ByVal fName)
	Dim srchCtrlWinBlock,filterRowId,this
	Set srchCtrlWinBlock = (CreateDictionary())
	filterRowId = CSmartStr(this_object.getFilterDivId_p2(recId,fName)) & "_win"
	setArrElement srchCtrlWinBlock,"filterRow_attrs",("id=""" & CSmartStr(filterRowId)) & """ "
	setArrElement srchCtrlWinBlock,"srchTypeCont_attrs_win",("id=""" & CSmartStr(this_object.getCtrlComboContId_p2(recId,fName))) & "_win"""
	doAssignmentByRef method_PanelSearchControl_buildSearchCtrlWinBlockArr,srchCtrlWinBlock
	Exit Function
End Function
Function method_PanelSearchControl_PanelSearchControl(byref this_object,ByVal id,ByVal tName,ByRef searchClauseObj,ByRef pageObj)
	method_SearchControl_SearchControl this_object,id,tName,searchClauseObj,pageObj
End Function
%>
