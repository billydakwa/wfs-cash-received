<%

'------ Class SearchPanel ------
Class SearchPanel
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
	Public Function init_SearchPanel_p1(ByRef params)
		DoAssignmentByRef init_SearchPanel_p1,method_SearchPanel_SearchPanel(me,params)
	End Function
	Public Function getSearchPerm_p1(ByVal tName)
		DoAssignmentByRef getSearchPerm_p1,method_SearchPanel_getSearchPerm(me,tName)
	End Function
	Public Function getSearchPerm()
		DoAssignmentByRef getSearchPerm,method_SearchPanel_getSearchPerm(me,"")
	End Function
	Public Function buildSearchPanel()
		DoAssignmentByRef buildSearchPanel,method_SearchPanel_buildSearchPanel(me)
	End Function
	Public Function searchAssign()
		DoAssignmentByRef searchAssign,method_SearchPanel_searchAssign(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
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
' SearchPanel implementation 
Function method_SearchPanel_SearchPanel(byref this_object,ByRef params)
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
	Dim this
	RunnerApply this_object,params
	doClassAssignmentByRef this_object,"searchClauseObj",this_object.pageObj.searchClauseObj
	doClassAssignment this_object,"id",this_object.pageObj.id
	doClassAssignment this_object,"tName",this_object.pageObj.tName
	doClassAssignment this_object,"panelState",this_object.searchClauseObj.getSrchPanelAttrs()
	doClassAssignment this_object,"isUseAjaxSuggest",GetTableData(this_object.tName,".isUseAjaxSuggest",true)
	doClassAssignment this_object,"searchControlBuilder",CreateClass("PanelSearchControl",4,this_object.id,this_object.tName,this_object.searchClauseObj,this_object.pageObj,Empty,Empty,Empty)
	if not (not IsEmpty(ArrayElement(params,"searchPerm"))) then
		doClassAssignment this_object,"searchPerm",this_object.getSearchPerm()
	end if
	if not (not IsEmpty(ArrayElement(params,"panelSearchFields"))) then
		doClassAssignment this_object,"panelSearchFields",GetTableData(this_object.tName,".panelSearchFields",CreateDictionary())
	end if
	if not (not IsEmpty(ArrayElement(params,"globSearchFields"))) then
		doClassAssignment this_object,"globSearchFields",GetTableData(this_object.tName,".globSearchFields",CreateDictionary())
	end if
	if not (not IsEmpty(ArrayElement(params,"allSearchFields"))) then
		doClassAssignment this_object,"allSearchFields",GetTableData(this_object.tName,".allSearchFields",CreateDictionary())
	end if
End Function
Function method_SearchPanel_getSearchPerm(byref this_object,ByVal tName)
	Dim strPerm
	doAssignment tName,IIF(tName,tName,this_object.tName)
	if not bValue(isGroupSecurity) then
		method_SearchPanel_getSearchPerm = true
		Exit Function
	else
		doAssignmentByRef strPerm,GetUserPermissions(tName)
		method_SearchPanel_getSearchPerm = not IsFalse(asp_strpos(strPerm,"S",empty))
		Exit Function
	end if
End Function
Function method_SearchPanel_buildSearchPanel(byref this_object)
	Dim srchPanelAttrs,this
	doAssignmentByRef srchPanelAttrs,this_object.searchClauseObj.getSrchPanelAttrs()
	this_object.searchAssign 
End Function
Function method_SearchPanel_searchAssign(byref this_object)
	Dim searchPerm,srchButtTitle
	this_object.pageObj.xt.assign_p2 "searchform",true
	this_object.pageObj.xt.assign_p2 "asearch_link",this_object.searchPerm
	this_object.pageObj.xt.assign_p2 "asearchlink_attrs",((((((("id=""asearch_" & CSmartStr(this_object.id)) & """ name=""asearch_") & CSmartStr(this_object.id)) & """ href=""") & CSmartStr(this_object.pageObj.shortTableName)) & "_search.asp"" onclick=""window.location.href='") & CSmartStr(this_object.pageObj.shortTableName)) & "_search.asp';return false;"""
	if bValue(isEnableSection508()) and bValue(this_object.searchPerm) then
		Set searchPerm = (CreateDictionary())
		setArrElement searchPerm,"begin","<a name=""skipsearch""></a>"
	else
		doAssignment searchPerm,this_object.searchPerm
	end if
	this_object.pageObj.xt.assign_p2 "search_records_block",searchPerm
	this_object.pageObj.xt.assign_p2 "searchform_text",true
	this_object.pageObj.xt.assign_p2 "searchform_search",true
	this_object.pageObj.xt.assign_p2 "searchform_showall",true
	if not bValue(this_object.searchClauseObj.isUsedSrch()) then
		this_object.pageObj.xt.assign_p2 "showAllCont_attrs","style=""display: none;"""
	end if
	srchButtTitle = "Search"
	this_object.pageObj.xt.assign_p2 "searchbutton_attrs",((("id=""searchButtTop" & CSmartStr(this_object.id)) & """  title=""") & CSmartStr(srchButtTitle)) & """"
	this_object.pageObj.xt.assign_p2 "showallbutton_attrs",("id=""showAll" & CSmartStr(this_object.id)) & """"
End Function
%>
