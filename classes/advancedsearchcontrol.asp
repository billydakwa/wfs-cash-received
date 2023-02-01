<%

'------ Class AdvancedSearchControl extends SearchControl ------
Class AdvancedSearchControl
	Public tName
	Public globSrchParams
	Public getSrchPanelAttrs
	Public dispNoneStyle
	Public pageObj
	Public searchClauseObj
	Public id
	Public Function init_AdvancedSearchControl_p4(ByVal id,ByVal tName,ByRef searchClauseObj,ByRef pageObj)
		DoAssignmentByRef init_AdvancedSearchControl_p4,method_AdvancedSearchControl_AdvancedSearchControl(me,id,tName,searchClauseObj,pageObj)
	End Function
	Public Function getCtrlSearchTypeOptions_p3(ByVal fName,ByVal selOpt,ByVal var_not)
		DoAssignmentByRef getCtrlSearchTypeOptions_p3,method_AdvancedSearchControl_getCtrlSearchTypeOptions(me,fName,selOpt,var_not)
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
' AdvancedSearchControl implementation little mellows
Function method_AdvancedSearchControl_AdvancedSearchControl(byref this_object,ByVal id,ByVal tName,ByRef searchClauseObj,ByRef pageObj)
	this_object.tName = ""
	doClassAssignment this_object,"globSrchParams",CreateDictionary()
	doClassAssignment this_object,"getSrchPanelAttrs",CreateDictionary()
	this_object.dispNoneStyle = "style=""display: none;"""
	this_object.pageObj = null
	this_object.searchClauseObj = false
	this_object.id = 1
	method_SearchControl_SearchControl this_object,id,tName,searchClauseObj,pageObj
	setArrElement this_object.getSrchPanelAttrs,"ctrlTypeComboStatus",true
End Function
Function method_AdvancedSearchControl_getCtrlSearchTypeOptions(byref this_object,ByVal fName,ByVal selOpt,ByVal var_not)
	doAssignmentByRef method_AdvancedSearchControl_getCtrlSearchTypeOptions,method_SearchControl_getCtrlSearchTypeOptions(this_object,fName,selOpt,false)
	Exit Function
End Function
%>
