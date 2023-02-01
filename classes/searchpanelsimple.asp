<%

'------ Class SearchPanelSimple extends SearchPanel ------
Class SearchPanelSimple
	Public srchPanelAttrs
	Public isDisplaySearchPanel
	Public tName
	Public dispNoneStyle
	Public pageObj
	Public searchClauseObj
	Public searchControlBuilder
	Public id
	Public panelState
	Public globSearchFields
	Public panelSearchFields
	Public allSearchFields
	Public isUseAjaxSuggest
	Public searchPerm
	Public Function init_SearchPanelSimple_p1(ByRef params)
		DoAssignmentByRef init_SearchPanelSimple_p1,method_SearchPanelSimple_SearchPanelSimple(me,params)
	End Function
	Public Function buildSearchPanel_p1(ByVal xtVarName)
		DoAssignmentByRef buildSearchPanel_p1,method_SearchPanelSimple_buildSearchPanel(me,xtVarName)
	End Function
	Public Function searchAssign()
		DoAssignmentByRef searchAssign,method_SearchPanelSimple_searchAssign(me)
	End Function
	Public Function DisplaySearchPanel_p1(ByRef params)
		DoAssignmentByRef DisplaySearchPanel_p1,method_SearchPanelSimple_DisplaySearchPanel(me,params)
	End Function
	Public Function getSearchPerm_p1(ByVal tName)
		DoAssignmentByRef getSearchPerm_p1,method_SearchPanel_getSearchPerm(me,tName)
	End Function
	Public Function getSearchPerm()
		DoAssignmentByRef getSearchPerm,method_SearchPanel_getSearchPerm(me,"")
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"srchPanelAttrs", srchPanelAttrs
		setArrElement out,"isDisplaySearchPanel", isDisplaySearchPanel
		setArrElement out,"tName", tName
		setArrElement out,"dispNoneStyle", dispNoneStyle
		setArrElement out,"pageObj", pageObj
		setArrElement out,"searchClauseObj", searchClauseObj
		setArrElement out,"searchControlBuilder", searchControlBuilder
		setArrElement out,"id", id
		setArrElement out,"panelState", panelState
		setArrElement out,"globSearchFields", globSearchFields
		setArrElement out,"panelSearchFields", panelSearchFields
		setArrElement out,"allSearchFields", allSearchFields
		setArrElement out,"isUseAjaxSuggest", isUseAjaxSuggest
		setArrElement out,"searchPerm", searchPerm
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment srchPanelAttrs, dict("srchPanelAttrs")
		DoAssignment isDisplaySearchPanel, dict("isDisplaySearchPanel")
		DoAssignment tName, dict("tName")
		DoAssignment dispNoneStyle, dict("dispNoneStyle")
		DoAssignment pageObj, dict("pageObj")
		DoAssignment searchClauseObj, dict("searchClauseObj")
		DoAssignment searchControlBuilder, dict("searchControlBuilder")
		DoAssignment id, dict("id")
		DoAssignment panelState, dict("panelState")
		DoAssignment globSearchFields, dict("globSearchFields")
		DoAssignment panelSearchFields, dict("panelSearchFields")
		DoAssignment allSearchFields, dict("allSearchFields")
		DoAssignment isUseAjaxSuggest, dict("isUseAjaxSuggest")
		DoAssignment searchPerm, dict("searchPerm")
	End Sub
' end serialize
End Class
' SearchPanelSimple implementation 
Function method_SearchPanelSimple_SearchPanelSimple(byref this_object,ByRef params)
	doClassAssignment this_object,"srchPanelAttrs",CreateDictionary()
	this_object.isDisplaySearchPanel = true
	this_object.tName = ""
	this_object.dispNoneStyle = "style=""display: none;"""
	this_object.pageObj = null
	this_object.searchClauseObj = null
	this_object.searchControlBuilder = null
	this_object.id = 1
	doClassAssignment this_object,"panelState",CreateDictionary()
	doClassAssignment this_object,"globSearchFields",CreateDictionary()
	doClassAssignment this_object,"panelSearchFields",CreateDictionary()
	doClassAssignment this_object,"allSearchFields",CreateDictionary()
	this_object.isUseAjaxSuggest = false
	this_object.searchPerm = false
	method_SearchPanel_SearchPanel this_object,params
	doClassAssignment this_object,"isDisplaySearchPanel",GetTableData(this_object.tName,".showSearchPanel",false)
End Function
Function method_SearchPanelSimple_buildSearchPanel(byref this_object,ByVal xtVarName)
	Dim searchPanel,this,params
	method_SearchPanel_buildSearchPanel this_object
	if bValue(this_object.isDisplaySearchPanel) then
		Set searchPanel = (CreateDictionary())
		setArrElement searchPanel,"method","DisplaySearchPanel"
		setArrElementByRef searchPanel,"object",this_object
		doClassAssignment this_object,"srchPanelAttrs",this_object.searchClauseObj.getSrchPanelAttrs()
		Set params = (CreateDictionary())
		setArrElement searchPanel,"params",params
		this_object.pageObj.xt.assignbyref_p2 xtVarName,searchPanel
	end if
End Function
Function method_SearchPanelSimple_searchAssign(byref this_object)
	Dim searchGlobalParams,searchPanelAttrs,searchOpt_mess,searchforAttrs,skruglAttrs,valSrchFor,simpleSearchTypeCombo,simpleSearchFieldCombo
	method_SearchPanel_searchAssign this_object
	doAssignmentByRef searchGlobalParams,this_object.searchClauseObj.getSearchGlobalParams()
	doAssignmentByRef searchPanelAttrs,this_object.searchClauseObj.getSrchPanelAttrs()
	this_object.pageObj.xt.assign_p2 "showHideSearchWin_attrs","align=ABSMIDDLE class=""searchPanelButton"" title=""Floating window"" alt=""Floating window"""
	doAssignment searchOpt_mess,IIF(ArrayElement(searchPanelAttrs,"srchOptShowStatus"),"Hide search options","Show search options")
	this_object.pageObj.xt.assign_p2 "showHideSearchPanel_attrs",((("align=ABSMIDDLE class=""searchPanelButton"" title=""" & CSmartStr(searchOpt_mess)) & """ alt=""") & CSmartStr(searchOpt_mess)) & """"
	searchforAttrs = ""
	if bValue(this_object.isUseAjaxSuggest) then
		searchforAttrs = CSmartStr(searchforAttrs) & "autocomplete=off "
	end if
	skruglAttrs = "style="""
	skruglAttrs = CSmartStr(skruglAttrs) & CSmartStr(IIF(ArrayElement(searchPanelAttrs,"srchOptShowStatus"),"""","display: none;"""))
	this_object.pageObj.xt.assignbyref_p2 "searchPanelBottomRound_attrs",skruglAttrs
	if not bValue(this_object.searchClauseObj.isUsedSrch()) then
		searchforAttrs = CSmartStr(searchforAttrs) & "style=""color: #C0C0C0;"""
	end if
	searchforAttrs = CSmartStr(searchforAttrs) & ((((" name=""ctlSearchFor" & CSmartStr(this_object.id)) & """ id=""ctlSearchFor") & CSmartStr(this_object.id)) & """")
	doAssignment valSrchFor,IIF(this_object.searchClauseObj.isUsedSrch(),ArrayElement(searchGlobalParams,"simpleSrch"),"search")
	searchforAttrs = CSmartStr(searchforAttrs) & ((" value=""" & CSmartStr(htmlspecialchars(valSrchFor))) & """")
	this_object.pageObj.xt.assignbyref_p2 "searchfor_attrs",searchforAttrs
	this_object.pageObj.xt.assign_p2 "searchPanelTopButtons",this_object.isDisplaySearchPanel
	if bValue(GetTableData(this_object.tName,".showSimpleSearchOptions",false)) then
		simpleSearchTypeCombo = ((("<SELECT id=""simpleSrchTypeCombo" & CSmartStr(this_object.id)) & """ NAME=""simpleSrchTypeCombo") & CSmartStr(this_object.id)) & """ SIZE=1 >"
		simpleSearchTypeCombo = CSmartStr(simpleSearchTypeCombo) & CSmartStr(this_object.searchControlBuilder.getSimpleSearchTypeCombo_p2(ArrayElement(searchGlobalParams,"simpleSrchTypeComboOpt"),ArrayElement(searchGlobalParams,"simpleSrchTypeComboNot")))
		simpleSearchTypeCombo = CSmartStr(simpleSearchTypeCombo) & "</SELECT>"
		this_object.pageObj.xt.assign_p2 "simpleSearchTypeCombo",simpleSearchTypeCombo
		simpleSearchFieldCombo = ((("<SELECT id=""simpleSrchFieldsCombo" & CSmartStr(this_object.id)) & """ NAME=""simpleSrchFieldsCombo") & CSmartStr(this_object.id)) & """ SIZE=1 >"
		simpleSearchFieldCombo = CSmartStr(simpleSearchFieldCombo) & CSmartStr(this_object.searchControlBuilder.simpleSearchFieldCombo_p2(this_object.allSearchFields,ArrayElement(searchGlobalParams,"simpleSrchFieldsComboOpt")))
		simpleSearchFieldCombo = CSmartStr(simpleSearchFieldCombo) & "</SELECT>"
		this_object.pageObj.xt.assign_p2 "simpleSearchFieldCombo",simpleSearchFieldCombo
	end if
End Function
Function method_SearchPanelSimple_DisplaySearchPanel(byref this_object,ByRef params)
	Dim dispNoneStyle,xt,searchRadio,showHideOpt_mess,srchCtrlBlocksArr,recId,j,srchFields,ctrlInd,isFieldNeedSecCtrl,i,srchCtrlBlocksWinArr
	dispNoneStyle = "style=""display: none;"""
	Set xt = (CreateClass("Xtempl",0,Empty,Empty,Empty,Empty,Empty,Empty,Empty))
	xt.assign_p2 "searchPanel",this_object.isDisplaySearchPanel
	xt.assign_p2 "id",this_object.id
	doAssignmentByRef searchRadio,this_object.searchControlBuilder.getSearchRadio()
	xt.assign_section_p3 "all_checkbox_label",ArrayElement(ArrayElement(searchRadio,"all_checkbox_label"),0),ArrayElement(ArrayElement(searchRadio,"all_checkbox_label"),1)
	xt.assign_section_p3 "any_checkbox_label",ArrayElement(ArrayElement(searchRadio,"any_checkbox_label"),0),ArrayElement(ArrayElement(searchRadio,"any_checkbox_label"),1)
	xt.assignbyref_p2 "all_checkbox",ArrayElement(searchRadio,"all_checkbox")
	xt.assignbyref_p2 "any_checkbox",ArrayElement(searchRadio,"any_checkbox")
	xt.assign_p2 "searchbutton_attrs",("id=""searchButton" & CSmartStr(this_object.id)) & """ "
	doAssignment showHideOpt_mess,IIF(ArrayElement(this_object.srchPanelAttrs,"ctrlTypeComboStatus"),"Hide options","Show options")
	xt.assign_p2 "showHideOpt_mess",showHideOpt_mess
	xt.assign_p2 "srchOpt_attrs","style=""display: none;"""
	if IsLess(0,this_object.searchClauseObj.getUsedCtrlsCount()) then
		xt.assign_p2 "srchCritTopCont_attrs",""
	else
		xt.assign_p2 "srchCritTopCont_attrs","style=""display: none;"""
	end if
	if IsLess(1,this_object.searchClauseObj.getUsedCtrlsCount()) then
		xt.assign_p2 "srchCritBottomCont_attrs",""
	else
		xt.assign_p2 "srchCritBottomCont_attrs","style=""display: none;"""
	end if
	if IsLess(0,this_object.searchClauseObj.getUsedCtrlsCount()) then
		xt.assign_p2 "bottomSearchButt_attrs",""
	else
		xt.assign_p2 "bottomSearchButt_attrs","style=""display: none;"""
	end if
	Set srchCtrlBlocksArr = (CreateDictionary())
	doAssignmentByRef recId,this_object.pageObj.genId()
	j = 0
	do while IsLess(j,asp_count(this_object.allSearchFields))
		this_object.pageObj.fillFieldToolTips_p1 ArrayElement(this_object.allSearchFields,j)
		xt.assign_p2 "addSearch_" & CSmartStr(GoodFieldName(ArrayElement(this_object.allSearchFields,j))),true
		doAssignmentByRef srchFields,this_object.searchClauseObj.getSearchCtrlParams_p1(ArrayElement(this_object.allSearchFields,j))
		ctrlInd = 0
		doAssignmentByRef isFieldNeedSecCtrl,this_object.searchControlBuilder.isNeedSecondCtrl_p1(ArrayElement(this_object.allSearchFields,j))
		if not bValue(asp_count(srchFields)) and bValue(asp_in_array(ArrayElement(this_object.allSearchFields,j),this_object.panelSearchFields,false)) then
			setArrElement srchFields,asp_count(srchFields),CreateDictionary4("opt","","not","","value1","","value2","")
		end if
		i = 0
		do while IsLess(i,asp_count(srchFields))
			setArrElement srchCtrlBlocksArr,asp_count(srchCtrlBlocksArr),this_object.searchControlBuilder.buildSearchCtrlBlockArr_p8(recId,ArrayElement(this_object.allSearchFields,j),ctrlInd,ArrayElement(ArrayElement(srchFields,i),"opt"),ArrayElement(ArrayElement(srchFields,i),"not"),false,ArrayElement(ArrayElement(srchFields,i),"value1"),ArrayElement(ArrayElement(srchFields,i),"value2"))
			setArrElement srchCtrlBlocksWinArr,asp_count(srchCtrlBlocksWinArr),this_object.searchControlBuilder.buildSearchCtrlWinBlockArr_p2(recId,ArrayElement(this_object.allSearchFields,j))
			if bValue(isFieldNeedSecCtrl) then
				setArrElementN this_object.pageObj.controlsMap,CreateArray3("search","searchBlocks",empty),CreateDictionary3("fName",ArrayElement(this_object.allSearchFields,j),"recId",recId,"ctrlsMap",CreateDictionary2(0,ctrlInd,1,CSmartDbl(ctrlInd)+1))
				ctrlInd = CSmartDbl(ctrlInd)+2
			else
				setArrElementN this_object.pageObj.controlsMap,CreateArray3("search","searchBlocks",empty),CreateDictionary3("fName",ArrayElement(this_object.allSearchFields,j),"recId",recId,"ctrlsMap",CreateDictionary1(0,ctrlInd))
				ctrlInd = CSmartDbl(ctrlInd)+1
			end if
			doAssignmentByRef recId,this_object.pageObj.genId()
			ctrlInd = 0
			i = CSmartDbl(i)+1
		loop
		if IsLess(asp_count(this_object.allSearchFields),gLoadSearchControls) then
			setArrElement srchCtrlBlocksArr,asp_count(srchCtrlBlocksArr),this_object.searchControlBuilder.buildSearchCtrlBlockArr_p8(recId,ArrayElement(this_object.allSearchFields,j),ctrlInd,"",false,true,"","")
			setArrElement srchCtrlBlocksWinArr,asp_count(srchCtrlBlocksWinArr),this_object.searchControlBuilder.buildSearchCtrlWinBlockArr_p2(recId,ArrayElement(this_object.allSearchFields,j))
			if bValue(isFieldNeedSecCtrl) then
				setArrElementN this_object.pageObj.controlsMap,CreateArray3("search","searchBlocks",empty),CreateDictionary3("fName",ArrayElement(this_object.allSearchFields,j),"recId",recId,"ctrlsMap",CreateDictionary2(0,ctrlInd,1,CSmartDbl(ctrlInd)+1))
				ctrlInd = CSmartDbl(ctrlInd)+2
			else
				setArrElementN this_object.pageObj.controlsMap,CreateArray3("search","searchBlocks",empty),CreateDictionary3("fName",ArrayElement(this_object.allSearchFields,j),"recId",recId,"ctrlsMap",CreateDictionary1(0,ctrlInd))
				ctrlInd = CSmartDbl(ctrlInd)+1
			end if
			doAssignmentByRef recId,this_object.pageObj.genId()
		end if
		j = CSmartDbl(j)+1
	loop
	xt.assign_loopsection_p2 "searchCtrlBlock",srchCtrlBlocksArr
	xt.assign_loopsection_p2 "searchCtrlBlock_win",srchCtrlBlocksWinArr
	xt.display_p1 CSmartStr(this_object.pageObj.shortTableName) & "_search_panel.htm"
End Function
%>
