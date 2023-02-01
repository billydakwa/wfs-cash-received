<%

'------ Class MenuItem ------
Class MenuItem
	Public depth
	Public id
	Public href
	Public params
	Public var_type
	Public name
	Public nameType
	Public style
	Public table
	Public linkType
	Public pageType
	Public showAsLink
	Public showAsGroup
	Public pageTypesInMenuForThisTable
	Public title
	Public openType
	Public children
	Public pageName
	Public Function init_MenuItem_p3(ByRef menuItemInfo,ByRef menuNodes,ByRef menuParent)
		DoAssignmentByRef init_MenuItem_p3,method_MenuItem_MenuItem(me,menuItemInfo,menuNodes,menuParent)
	End Function
	Public Function AddChild_p1(ByRef child)
		DoAssignmentByRef AddChild_p1,method_MenuItem_AddChild(me,child)
	End Function
	Public Function setUrl_p1(ByVal href)
		DoAssignmentByRef setUrl_p1,method_MenuItem_setUrl(me,href)
	End Function
	Public Function getUrl()
		DoAssignmentByRef getUrl,method_MenuItem_getUrl(me)
	End Function
	Public Function setParams_p1(ByVal params)
		DoAssignmentByRef setParams_p1,method_MenuItem_setParams(me,params)
	End Function
	Public Function getParams()
		DoAssignmentByRef getParams,method_MenuItem_getParams(me)
	End Function
	Public Function setTitle_p1(ByVal title)
		DoAssignmentByRef setTitle_p1,method_MenuItem_setTitle(me,title)
	End Function
	Public Function getTitle()
		DoAssignmentByRef getTitle,method_MenuItem_getTitle(me)
	End Function
	Public Function setTable_p1(ByVal table)
		DoAssignmentByRef setTable_p1,method_MenuItem_setTable(me,table)
	End Function
	Public Function getTable()
		DoAssignmentByRef getTable,method_MenuItem_getTable(me)
	End Function
	Public Function setPageType_p1(ByVal pType)
		DoAssignmentByRef setPageType_p1,method_MenuItem_setPageType(me,pType)
	End Function
	Public Function getPageType()
		DoAssignmentByRef getPageType,method_MenuItem_getPageType(me)
	End Function
	Public Function getLinkType()
		DoAssignmentByRef getLinkType,method_MenuItem_getLinkType(me)
	End Function
	Public Function buildTreeMenuStructure_p2(ByRef menuParent,ByRef menuNodes)
		DoAssignmentByRef buildTreeMenuStructure_p2,method_MenuItem_buildTreeMenuStructure(me,menuParent,menuNodes)
	End Function
	Public Function checkLinkShowStatus()
		DoAssignmentByRef checkLinkShowStatus,method_MenuItem_checkLinkShowStatus(me)
	End Function
	Public Function isUserHaveTablePerm()
		DoAssignmentByRef isUserHaveTablePerm,method_MenuItem_isUserHaveTablePerm(me)
	End Function
	Public Function checkGroupShowStatus()
		DoAssignmentByRef checkGroupShowStatus,method_MenuItem_checkGroupShowStatus(me)
	End Function
	Public Function isShowAsGroup()
		DoAssignmentByRef isShowAsGroup,method_MenuItem_isShowAsGroup(me)
	End Function
	Public Function isShowAsLink()
		DoAssignmentByRef isShowAsLink,method_MenuItem_isShowAsLink(me)
	End Function
	Public Function isGroup()
		DoAssignmentByRef isGroup,method_MenuItem_isGroup(me)
	End Function
	Public Function isSeparator()
		DoAssignmentByRef isSeparator,method_MenuItem_isSeparator(me)
	End Function
	Public Function makeOffset_p1(ByVal depth)
		DoAssignmentByRef makeOffset_p1,method_MenuItem_makeOffset(me,depth)
	End Function
	Public Function assignMenuAttrsToTempl_p1(ByRef xt)
		DoAssignmentByRef assignMenuAttrsToTempl_p1,method_MenuItem_assignMenuAttrsToTempl(me,xt)
	End Function
	Public Function setCurrMenuElem_p1(ByRef xt)
		DoAssignmentByRef setCurrMenuElem_p1,method_MenuItem_setCurrMenuElem(me,xt)
	End Function
	Public Function setNotCurrMenuElem_p1(ByRef xt)
		DoAssignmentByRef setNotCurrMenuElem_p1,method_MenuItem_setNotCurrMenuElem(me,xt)
	End Function
	Public Function setAsCurrMenuElem_p1(ByRef xt)
		DoAssignmentByRef setAsCurrMenuElem_p1,method_MenuItem_setAsCurrMenuElem(me,xt)
	End Function
	Public Function assignLinks_p1(ByRef xt)
		DoAssignmentByRef assignLinks_p1,method_MenuItem_assignLinks(me,xt)
	End Function
	Public Function assignGroupOnly_p1(ByRef xt)
		DoAssignmentByRef assignGroupOnly_p1,method_MenuItem_assignGroupOnly(me,xt)
	End Function
	Public Function isSetParentElem()
		DoAssignmentByRef isSetParentElem,method_MenuItem_isSetParentElem(me)
	End Function
	Public Function isThisPageInMenu()
		DoAssignmentByRef isThisPageInMenu,method_MenuItem_isThisPageInMenu(me)
	End Function
	Public Function isIdInPageTypesArr_p1(ByVal id)
		DoAssignmentByRef isIdInPageTypesArr_p1,method_MenuItem_isIdInPageTypesArr(me,id)
	End Function
	Public Function changeKeysInLowerCaseFromArr_p1(ByVal arr)
		DoAssignmentByRef changeKeysInLowerCaseFromArr_p1,method_MenuItem_changeKeysInLowerCaseFromArr(me,arr)
	End Function
	Public Function clearMenuSession()
		DoAssignmentByRef clearMenuSession,method_MenuItem_clearMenuSession(me)
	End Function
	Public Function setMenuSession()
		DoAssignmentByRef setMenuSession,method_MenuItem_setMenuSession(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"depth", depth
		setArrElement out,"id", id
		setArrElement out,"href", href
		setArrElement out,"params", params
		setArrElement out,"var_type", var_type
		setArrElement out,"name", name
		setArrElement out,"nameType", nameType
		setArrElement out,"style", style
		setArrElement out,"table", table
		setArrElement out,"linkType", linkType
		setArrElement out,"pageType", pageType
		setArrElement out,"showAsLink", showAsLink
		setArrElement out,"showAsGroup", showAsGroup
		setArrElement out,"pageTypesInMenuForThisTable", pageTypesInMenuForThisTable
		setArrElement out,"title", title
		setArrElement out,"openType", openType
		setArrElement out,"children", children
		setArrElement out,"pageName", pageName
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment depth, dict("depth")
		DoAssignment id, dict("id")
		DoAssignment href, dict("href")
		DoAssignment params, dict("params")
		DoAssignment var_type, dict("var_type")
		DoAssignment name, dict("name")
		DoAssignment nameType, dict("nameType")
		DoAssignment style, dict("style")
		DoAssignment table, dict("table")
		DoAssignment linkType, dict("linkType")
		DoAssignment pageType, dict("pageType")
		DoAssignment showAsLink, dict("showAsLink")
		DoAssignment showAsGroup, dict("showAsGroup")
		DoAssignment pageTypesInMenuForThisTable, dict("pageTypesInMenuForThisTable")
		DoAssignment title, dict("title")
		DoAssignment openType, dict("openType")
		DoAssignment children, dict("children")
		DoAssignment pageName, dict("pageName")
	End Sub
' end serialize
End Class
' MenuItem implementation 
Function method_MenuItem_MenuItem(byref this_object,ByRef menuItemInfo,ByRef menuNodes,ByRef menuParent)
	doClassAssignment this_object,"pageTypesInMenuForThisTable",CreateDictionary()
	doClassAssignment this_object,"children",CreateDictionary()
	this_object.pageName = ""
	Dim this
	if not IsEmpty(pageObject) then
		doClassAssignment this_object,"pageName",pageObject.getPageType()
	end if
	if IsEqual(ArrayElement(menuItemInfo,"id"),0) then
		this_object.depth = 0
	else
		this_object.depth = CSmartDbl(menuParent.depth)+1
	end if
	doClassAssignment this_object,"id",ArrayElement(menuItemInfo,"id")
	doClassAssignment this_object,"name",ArrayElement(menuItemInfo,"name")
	doClassAssignment this_object,"nameType",ArrayElement(menuItemInfo,"nameType")
	doClassAssignment this_object,"style",ArrayElement(menuItemInfo,"style")
	doClassAssignment this_object,"href",ArrayElement(menuItemInfo,"href")
	doClassAssignment this_object,"params",ArrayElement(menuItemInfo,"params")
	doClassAssignment this_object,"table",ArrayElement(menuItemInfo,"table")
	doClassAssignment this_object,"var_type",ArrayElement(menuItemInfo,"type")
	doClassAssignment this_object,"linkType",ArrayElement(menuItemInfo,"linkType")
	doClassAssignment this_object,"pageType",ArrayElement(menuItemInfo,"pageType")
	doClassAssignment this_object,"title",ArrayElement(menuItemInfo,"title")
	doClassAssignment this_object,"openType",ArrayElement(menuItemInfo,"openType")
	doClassAssignment this_object,"showAsLink",this_object.checkLinkShowStatus()
	this_object.buildTreeMenuStructure_p2 this_object,menuNodes
	doClassAssignment this_object,"showAsGroup",this_object.checkGroupShowStatus()
End Function
Function method_MenuItem_AddChild(byref this_object,ByRef child)
	Dim res
	res = true
	if bValue(globalEvents.exists_p1("ModifyMenuItem")) then
		doAssignmentByRef res,globalEvents.ModifyMenuItem_p1(child)
	end if
	if bValue(res) then
		setArrElement this_object.children,asp_count(this_object.children),child
	end if
End Function
Function method_MenuItem_setUrl(byref this_object,ByVal href)
	doClassAssignment this_object,"href",href
	if IsEqual(this_object.linkType,"Internal") then
		this_object.linkType = "External"
	end if
End Function
Function method_MenuItem_getUrl(byref this_object)
	doAssignmentByRef method_MenuItem_getUrl,this_object.href
	Exit Function
End Function
Function method_MenuItem_setParams(byref this_object,ByVal params)
	doClassAssignment this_object,"params",params
End Function
Function method_MenuItem_getParams(byref this_object)
	doAssignmentByRef method_MenuItem_getParams,this_object.params
	Exit Function
End Function
Function method_MenuItem_setTitle(byref this_object,ByVal title)
	doClassAssignment this_object,"title",title
End Function
Function method_MenuItem_getTitle(byref this_object)
	doAssignmentByRef method_MenuItem_getTitle,this_object.title
	Exit Function
End Function
Function method_MenuItem_setTable(byref this_object,ByVal table)
	this_object.table
End Function
Function method_MenuItem_getTable(byref this_object)
	doAssignmentByRef method_MenuItem_getTable,this_object.table
	Exit Function
End Function
Function method_MenuItem_setPageType(byref this_object,ByVal pType)
	doClassAssignment this_object,"pageType",pType
End Function
Function method_MenuItem_getPageType(byref this_object)
	doAssignmentByRef method_MenuItem_getPageType,this_object.pageType
	Exit Function
End Function
Function method_MenuItem_getLinkType(byref this_object)
	doAssignmentByRef method_MenuItem_getLinkType,this_object.linkType
	Exit Function
End Function
Function method_MenuItem_buildTreeMenuStructure(byref this_object,ByRef menuParent,ByRef menuNodes)
	Dim i,menuChild
	i = 0
	do while IsLess(i,asp_count(menuNodes))
		if IsEqual(ArrayElement(ArrayElement(menuNodes,i),"parent"),menuParent.id) then
			Set menuChild = (CreateClass("MenuItem",3,ArrayElement(menuNodes,i),menuNodes,menuParent,Empty,Empty,Empty,Empty))
			menuParent.AddChild_p1 menuChild
		end if
		if not IsEqual(ArrayElement(ArrayElement(menuNodes,i),"type"),"Separator") then
			if not (not IsEmpty(ArrayElement(this_object.pageTypesInMenuForThisTable,ArrayElement(ArrayElement(menuNodes,i),"pageType")))) and IsEqual(this_object.table,ArrayElement(ArrayElement(menuNodes,i),"table")) then
				setArrElementN this_object.pageTypesInMenuForThisTable,CreateArray2(ArrayElement(ArrayElement(menuNodes,i),"pageType"),"count"),1
				setArrElementN this_object.pageTypesInMenuForThisTable,CreateArray2(ArrayElement(ArrayElement(menuNodes,i),"pageType"),"ids"),CreateDictionary1("0",ArrayElement(ArrayElement(menuNodes,i),"id"))
			else
				if IsEqual(this_object.table,ArrayElement(ArrayElement(menuNodes,i),"table")) then
					setArrElementN this_object.pageTypesInMenuForThisTable,CreateArray2(ArrayElement(ArrayElement(menuNodes,i),"pageType"),"count"),CSmartDbl(ArrayElement(ArrayElement(this_object.pageTypesInMenuForThisTable,ArrayElement(ArrayElement(menuNodes,i),"pageType")),"count"))+1
					setArrElementN this_object.pageTypesInMenuForThisTable,CreateArray3(ArrayElement(ArrayElement(menuNodes,i),"pageType"),"ids",empty),ArrayElement(ArrayElement(menuNodes,i),"id")
				end if
			end if
		end if
		i = CSmartDbl(i)+1
	loop
End Function
Function method_MenuItem_checkLinkShowStatus(byref this_object)
	Dim this
	if IsEqual(this_object.linkType,"External") and IsLess(0,asp_strlen(this_object.href)) then
		method_MenuItem_checkLinkShowStatus = true
		Exit Function
	end if
	if IsEqual(this_object.linkType,"Separator") then
		method_MenuItem_checkLinkShowStatus = true
		Exit Function
	end if
	if IsEqual(this_object.linkType,"Internal") and bValue(this_object.isUserHaveTablePerm()) then
		method_MenuItem_checkLinkShowStatus = true
		Exit Function
	end if
	method_MenuItem_checkLinkShowStatus = false
	Exit Function
End Function
Function method_MenuItem_isUserHaveTablePerm(byref this_object)
	Dim strPerm,pageType
	if IsEqual(this_object.pageType,"WebReports") then
		method_MenuItem_isUserHaveTablePerm = true
		Exit Function
	end if
	if not bValue(asp_strlen(this_object.table)) then
		method_MenuItem_isUserHaveTablePerm = false
		Exit Function
	end if
	doAssignmentByRef strPerm,GetUserPermissions(this_object.table)
	pageType = ""
	if ((IsEqual(this_object.pageType,"List") or IsEqual(this_object.pageType,"Search")) or IsEqual(this_object.pageType,"Report")) or IsEqual(this_object.pageType,"Chart") then
		pageType = "S"
	else
		if IsEqual(this_object.pageType,"Add") then
			pageType = "A"
		else
			if IsEqual(this_object.pageType,"Edit") then
				pageType = "E"
			else
				if IsEqual(this_object.pageType,"Print") then
					pageType = "P"
				end if
			end if
		end if
	end if
	if not IsFalse(asp_strpos(strPerm,pageType,empty)) then
		method_MenuItem_isUserHaveTablePerm = true
		Exit Function
	else
		method_MenuItem_isUserHaveTablePerm = false
		Exit Function
	end if
End Function
Function method_MenuItem_checkGroupShowStatus(byref this_object)
	Dim this,i
	if not bValue(this_object.isGroup()) then
		method_MenuItem_checkGroupShowStatus = false
		Exit Function
	end if
	i = 0
	do while IsLess(i,asp_count(this_object.children))
		if bValue(ArrayElement(this_object.children,i).checkGroupShowStatus()) then
			method_MenuItem_checkGroupShowStatus = true
			Exit Function
		else
			if bValue(ArrayElement(this_object.children,i).isShowAsLink()) and not bValue(ArrayElement(this_object.children,i).isSeparator()) then
				method_MenuItem_checkGroupShowStatus = true
				Exit Function
			end if
		end if
		i = CSmartDbl(i)+1
	loop
	method_MenuItem_checkGroupShowStatus = false
	Exit Function
End Function
Function method_MenuItem_isShowAsGroup(byref this_object)
	doAssignmentByRef method_MenuItem_isShowAsGroup,this_object.showAsGroup
	Exit Function
End Function
Function method_MenuItem_isShowAsLink(byref this_object)
	doAssignmentByRef method_MenuItem_isShowAsLink,this_object.showAsLink
	Exit Function
End Function
Function method_MenuItem_isGroup(byref this_object)
	method_MenuItem_isGroup = IsEqual(this_object.var_type,"Group")
	Exit Function
End Function
Function method_MenuItem_isSeparator(byref this_object)
	method_MenuItem_isSeparator = IsEqual(this_object.var_type,"Separator")
	Exit Function
End Function
Function method_MenuItem_makeOffset(byref this_object,ByVal depth)
	Dim nbsps,i
	nbsps = ""
	i = 0
	do while IsLess(i,depth)
		nbsps = CSmartStr(nbsps) & "&nbsp;&nbsp;"
		i = CSmartDbl(i)+1
	loop
	doAssignmentByRef method_MenuItem_makeOffset,nbsps
	Exit Function
End Function
Function method_MenuItem_assignMenuAttrsToTempl(byref this_object,ByRef xt)
	Dim this,i
	if bValue(strTableName) then
		xt.assign_p2 "needcurrent_item",true
	else
		xt.assign_p2 "notneedcurrent_item",true
	end if
	if bValue(this_object.isSeparator()) then
		if IsEqual(this_object.name,"-------") then
			xt.assign_p2 ("item" & CSmartStr(this_object.id)) & "_separator","<HR>"
		else
			xt.assign_p2 ("item" & CSmartStr(this_object.id)) & "_separator",((("<a class=""ahr"" style=""" & CSmartStr(this_object.style)) & """>") & CSmartStr(this_object.name)) & "</a>"
		end if
		xt.assign_p2 ("item" & CSmartStr(this_object.id)) & "_optionattrs","disabled"
	end if
	if ((not bValue(this_object.isShowAsGroup()) and not bValue(this_object.isShowAsLink())) and not bValue(this_object.isSeparator())) and not IsEqual(this_object.id,0) then
		Exit Function
	end if
	if bValue(this_object.isShowAsGroup()) then
		xt.assign_p2 ("item" & CSmartStr(this_object.id)) & "_groupimage",true
	else
		xt.assign_p2 ("item" & CSmartStr(this_object.id)) & "_itemimage",true
	end if
	xt.assign_p2 ("item" & CSmartStr(this_object.id)) & "_menulink",true
	xt.assign_p2 ("item" & CSmartStr(this_object.id)) & "_nbsps",this_object.makeOffset_p1(this_object.depth)
	if bValue(this_object.isShowAsLink()) then
		this_object.assignLinks_p1 xt
	else
		if bValue(this_object.isShowAsGroup()) then
			this_object.assignGroupOnly_p1 xt
		end if
	end if
	i = 0
	do while IsLess(i,asp_count(this_object.children))
		ArrayElement(this_object.children,i).assignMenuAttrsToTempl_p1 xt
		i = CSmartDbl(i)+1
	loop
End Function
Function method_MenuItem_setCurrMenuElem(byref this_object,ByRef xt)
	Dim this,var_SESSION,i
	this_object.setNotCurrMenuElem_p1 xt
	if (bValue(this_object.pageName) and IsEqual(strTableName,this_object.table)) and not IsEqual(this_object.id,0) then
		if (not IsEmpty(Session("menuItemId")) and bValue(this_object.isIdInPageTypesArr_p1(Session("menuItemId")))) and IsLess(1,ArrayElement(ArrayElement(this_object.pageTypesInMenuForThisTable,this_object.pageType),"count")) then
			if IsEqual(Session("menuItemId"),this_object.id) then
				this_object.setAsCurrMenuElem_p1 xt
				method_MenuItem_setCurrMenuElem = true
				Exit Function
			end if
		else
			if IsEqual(this_object.pageName,asp_strtolower(this_object.pageType)) then
				this_object.setAsCurrMenuElem_p1 xt
				method_MenuItem_setCurrMenuElem = true
				Exit Function
			else
				if not bValue(this_object.isSetParentElem()) and not bValue(this_object.isThisPageInMenu()) then
					this_object.setAsCurrMenuElem_p1 xt
					method_MenuItem_setCurrMenuElem = true
					Exit Function
				end if
			end if
		end if
	end if
	i = 0
	do while IsLess(i,asp_count(this_object.children))
		if bValue(ArrayElement(this_object.children,i).setCurrMenuElem_p1(xt)) then
			method_MenuItem_setCurrMenuElem = true
			Exit Function
		end if
		i = CSmartDbl(i)+1
	loop
	method_MenuItem_setCurrMenuElem = false
	Exit Function
End Function
Function method_MenuItem_setNotCurrMenuElem(byref this_object,ByRef xt)
	Dim i
	xt.assign_p2 ("item" & CSmartStr(this_object.id)) & "_notcurrent",true
	i = 0
	do while IsLess(i,asp_count(this_object.children))
		ArrayElement(this_object.children,i).setNotCurrMenuElem_p1 xt
		i = CSmartDbl(i)+1
	loop
End Function
Function method_MenuItem_setAsCurrMenuElem(byref this_object,ByRef xt)
	xt.assign_p2 ("item" & CSmartStr(this_object.id)) & "_notcurrent",false
	xt.assign_p2 ("item" & CSmartStr(this_object.id)) & "_current",true
End Function
Function method_MenuItem_assignLinks(byref this_object,ByRef xt)
	Dim attrForAssign,menuIdGetParam
	xt.assign_p2 ("item" & CSmartStr(this_object.id)) & "_title",this_object.title
	attrForAssign = ((((" id=""itemlink" & CSmartStr(this_object.id)) & """ title=""") & CSmartStr(this_object.title)) & """ ") & CSmartStr(IIF(this_object.style,("style=""" & CSmartStr(this_object.style)) & """",""))
	if IsEqual(this_object.openType,"NewWindow") then
		attrForAssign = CSmartStr(attrForAssign) & " target=""_blank"""
	end if
	if IsEqual(this_object.linkType,"Internal") then
		if not IsEqual(this_object.pageType,"WebReports") then
			menuIdGetParam = ""
			if IsLess(1,ArrayElement(ArrayElement(this_object.pageTypesInMenuForThisTable,this_object.pageType),"count")) then
				menuIdGetParam = "?menuItemId=" & CSmartStr(this_object.id)
			end if
			if bValue(this_object.params) then
				if bValue(menuIdGetParam) then
					menuIdGetParam = CSmartStr(menuIdGetParam) & ("&" & CSmartStr(this_object.params))
				else
					menuIdGetParam = CSmartStr(menuIdGetParam) & ("?" & CSmartStr(this_object.params))
				end if
			end if
			attrForAssign = CSmartStr(attrForAssign) & ((((((" href=""" & CSmartStr(GetTableURL(this_object.table))) & "_") & CSmartStr(asp_strtolower(this_object.pageType))) & ".asp") & CSmartStr(menuIdGetParam)) & """")
			xt.assign_p2 ("item" & CSmartStr(this_object.id)) & "_optionattrs",(((("value=""" & CSmartStr(GetTableURL(this_object.table))) & "_") & CSmartStr(asp_strtolower(this_object.pageType))) & ".asp"" ") & CSmartStr(IIF(IsEqual(this_object.openType,"NewWindow"),"link=""External""",""))
		else
			attrForAssign = CSmartStr(attrForAssign) & " href=""webreport.asp"""
			xt.assign_p2 ("item" & CSmartStr(this_object.id)) & "_optionattrs"," value=""webreport.asp"""
		end if
	else
		if IsEqual(this_object.linkType,"External") then
			attrForAssign = CSmartStr(attrForAssign) & ((" href=""" & CSmartStr(this_object.href)) & """")
			xt.assign_p2 ("item" & CSmartStr(this_object.id)) & "_optionattrs",(("value=""" & CSmartStr(this_object.href)) & """ ") & CSmartStr(IIF(IsEqual(this_object.openType,"NewWindow"),"link=""External""",""))
		end if
	end if
	xt.assign_p2 ("item" & CSmartStr(this_object.id)) & "_menulink_attrs",attrForAssign
End Function
Function method_MenuItem_assignGroupOnly(byref this_object,ByRef xt)
	Dim attrForAssign
	xt.assign_p2 ("item" & CSmartStr(this_object.id)) & "_title",this_object.title
	attrForAssign = ((((" id=""itemlink" & CSmartStr(this_object.id)) & """ title=""") & CSmartStr(this_object.title)) & """ ") & CSmartStr(IIF(this_object.style,(" style=""cursor:default;text-decoration:none; " & CSmartStr(this_object.style)) & """",""))
	xt.assign_p2 ("item" & CSmartStr(this_object.id)) & "_menulink_attrs",attrForAssign
	xt.assign_p2 ("item" & CSmartStr(this_object.id)) & "_optionattrs","disabled"
End Function
Function method_MenuItem_isSetParentElem(byref this_object)
	Dim pageTypes,pageTypesInLowCase,this
	Set pageTypes = (CreateDictionary6(Empty,"list",Empty,"chart",Empty,"report",Empty,"search",Empty,"add",Empty,"print"))
	doAssignmentByRef pageTypesInLowCase,this_object.changeKeysInLowerCaseFromArr_p1(this_object.pageTypesInMenuForThisTable)
	do
		If IsEqual(asp_strtolower(this_object.pageType),"list") then
			method_MenuItem_isSetParentElem = false
			Exit Function
			exit do
		End If
		If IsEqual(asp_strtolower(this_object.pageType),"list") or IsEqual(asp_strtolower(this_object.pageType),"chart") then
			if bValue(asp_count(asp_array_intersect(asp_array_slice(pageTypes,0,1),pageTypesInLowCase))) then
				method_MenuItem_isSetParentElem = true
				Exit Function
			else
				method_MenuItem_isSetParentElem = false
				Exit Function
			end if
			exit do
		End If
		If IsEqual(asp_strtolower(this_object.pageType),"list") or IsEqual(asp_strtolower(this_object.pageType),"chart") or IsEqual(asp_strtolower(this_object.pageType),"report") then
			if bValue(asp_count(asp_array_intersect(asp_array_slice(pageTypes,0,2),pageTypesInLowCase))) then
				method_MenuItem_isSetParentElem = true
				Exit Function
			else
				method_MenuItem_isSetParentElem = false
				Exit Function
			end if
			exit do
		End If
		If IsEqual(asp_strtolower(this_object.pageType),"list") or IsEqual(asp_strtolower(this_object.pageType),"chart") or IsEqual(asp_strtolower(this_object.pageType),"report") or IsEqual(asp_strtolower(this_object.pageType),"search") then
			if bValue(asp_count(asp_array_intersect(asp_array_slice(pageTypes,0,3),pageTypesInLowCase))) then
				method_MenuItem_isSetParentElem = true
				Exit Function
			else
				method_MenuItem_isSetParentElem = false
				Exit Function
			end if
			exit do
		End If
		If IsEqual(asp_strtolower(this_object.pageType),"list") or IsEqual(asp_strtolower(this_object.pageType),"chart") or IsEqual(asp_strtolower(this_object.pageType),"report") or IsEqual(asp_strtolower(this_object.pageType),"search") or IsEqual(asp_strtolower(this_object.pageType),"add") then
			if bValue(asp_count(asp_array_intersect(asp_array_slice(pageTypes,0,4),pageTypesInLowCase))) then
				method_MenuItem_isSetParentElem = true
				Exit Function
			else
				method_MenuItem_isSetParentElem = false
				Exit Function
			end if
			exit do
		End If
		If IsEqual(asp_strtolower(this_object.pageType),"list") or IsEqual(asp_strtolower(this_object.pageType),"chart") or IsEqual(asp_strtolower(this_object.pageType),"report") or IsEqual(asp_strtolower(this_object.pageType),"search") or IsEqual(asp_strtolower(this_object.pageType),"add") or IsEqual(asp_strtolower(this_object.pageType),"print") then
			if bValue(asp_count(asp_array_intersect(asp_array_slice(pageTypes,0,5),pageTypesInLowCase))) then
				method_MenuItem_isSetParentElem = true
				Exit Function
			else
				method_MenuItem_isSetParentElem = false
				Exit Function
			end if
			exit do
		End If
		exit do
	Loop While false
End Function
Function method_MenuItem_isThisPageInMenu(byref this_object)
	Dim pageTypesInLowCase,this
	doAssignmentByRef pageTypesInLowCase,this_object.changeKeysInLowerCaseFromArr_p1(this_object.pageTypesInMenuForThisTable)
	doAssignmentByRef method_MenuItem_isThisPageInMenu,asp_in_array(this_object.pageName,pageTypesInLowCase,false)
	Exit Function
End Function
Function method_MenuItem_isIdInPageTypesArr(byref this_object,ByVal id)
	Dim pageTypeInfoArr
	GetCollectionBounds this_object.pageTypesInMenuForThisTable,i_menuitem_loopIdx8,i_menuitem_loopMax8
	do while i_menuitem_loopIdx8<=i_menuitem_loopMax8
		i_menuitem_arrKey8 = GetCollectionKey(this_object.pageTypesInMenuForThisTable,i_menuitem_loopIdx8)
		doAssignment pageTypeInfoArr,ArrayElement(this_object.pageTypesInMenuForThisTable,i_menuitem_arrKey8)
		if bValue(asp_in_array(id,ArrayElement(pageTypeInfoArr,"ids"),false)) then
			method_MenuItem_isIdInPageTypesArr = true
			Exit Function
		end if
		i_menuitem_loopIdx8=i_menuitem_loopIdx8+1
	loop
	method_MenuItem_isIdInPageTypesArr = false
	Exit Function
End Function
Function method_MenuItem_changeKeysInLowerCaseFromArr(byref this_object,ByVal arr)
	Dim lowArr,key
	Set lowArr = (CreateDictionary())
	GetCollectionBounds arr,i_menuitem_loopIdx9,i_menuitem_loopMax9
	do while i_menuitem_loopIdx9<=i_menuitem_loopMax9
		key = GetCollectionKey(arr,i_menuitem_loopIdx9)
		doAssignment val,ArrayElement(arr,key)
		setArrElement lowArr,asp_count(lowArr),asp_strtolower(key)
		i_menuitem_loopIdx9=i_menuitem_loopIdx9+1
	loop
	doAssignmentByRef method_MenuItem_changeKeysInLowerCaseFromArr,lowArr
	Exit Function
End Function
Function method_MenuItem_clearMenuSession(byref this_object)
	Dim var_SESSION
	if not IsEmpty(Session("menuItemId")) then
		asp_unsetElement Session,"menuItemId"
	end if
End Function
Function method_MenuItem_setMenuSession(byref this_object)
	Dim var_SESSION
	if bValue(postvalue("menuItemId")) then
		setArrElement Session,"menuItemId",postvalue("menuItemId")
	end if
End Function
%>
