<%

'------ Class SearchPanelLookup extends SearchPanel ------
Class SearchPanelLookup
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
	Public Function init_SearchPanelLookup_p1(ByRef params)
		DoAssignmentByRef init_SearchPanelLookup_p1,method_SearchPanelLookup_SearchPanelLookup(me,params)
	End Function
	Public Function searchAssign()
		DoAssignmentByRef searchAssign,method_SearchPanelLookup_searchAssign(me)
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
' SearchPanelLookup implementation 
Function method_SearchPanelLookup_SearchPanelLookup(byref this_object,ByRef params)
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
End Function
Function method_SearchPanelLookup_searchAssign(byref this_object)
	Dim searchGlobalParams,searchforAttrs
	method_SearchPanel_searchAssign this_object
	doAssignmentByRef searchGlobalParams,this_object.searchClauseObj.getSearchGlobalParams()
	searchforAttrs = (((" size=""15"" name=""ctlSearchFor" & CSmartStr(this_object.id)) & """ id=""ctlSearchFor") & CSmartStr(this_object.id)) & """"
	searchforAttrs = CSmartStr(searchforAttrs) & ((" value=""" & CSmartStr(htmlspecialchars(ArrayElement(searchGlobalParams,"simpleSrch")))) & """")
	if not bValue(this_object.searchClauseObj.isUsedSrch()) then
		searchforAttrs = CSmartStr(searchforAttrs) & "style=""color: #C0C0C0;"""
	end if
	this_object.pageObj.xt.assign_p2 "searchfor_attrs",searchforAttrs
	this_object.pageObj.xt.assign_p2 "searchform",true
End Function
%>
