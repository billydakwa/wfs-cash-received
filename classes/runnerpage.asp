<%
Function RunnerApply(ByRef obj,ByRef argsArr)
	Dim key
	GetCollectionBounds argsArr,c_runnerpage_loopIdx2,c_runnerpage_loopMax2
	do while c_runnerpage_loopIdx2<=c_runnerpage_loopMax2
		key = GetCollectionKey(argsArr,c_runnerpage_loopIdx2)
		doAssignment var,ArrayElement(argsArr,key)
		setObjectProperty obj,key,ArrayElement(argsArr,key)
		c_runnerpage_loopIdx2=c_runnerpage_loopIdx2+1
	loop
End Function
asp_include getabspath("classes/files.asp"),false
Const PAGE_LIST = "list"
Const PAGE_ADD = "add"
Const PAGE_EDIT = "edit"
Const PAGE_VIEW = "view"
Const PAGE_MENU = "menu"
Const PAGE_LOGIN = "login"
Const PAGE_REGISTER = "register"
Const PAGE_REMIND = "remind"
Const PAGE_CHANGEPASS = "changepass"
Const PAGE_SEARCH = "search"
Const PAGE_REPORT = "report"
Const PAGE_CHART = "chart"
Const PAGE_PRINT = "print"
Const PAGE_EXPORT = "export"
Const PAGE_IMPORT = "import"
Const PAGE_ADMIN_MEMBERS = "admin_members"
Const PAGE_ADMIN_RIGHTS = "admin_rights"

'------ Class RunnerPage ------
Class RunnerPage
	Public id
	Public totalCode
	Public calendar
	Public timepicker
	Public useIbox
	Public isUseToolTips
	Public isUseAjaxSuggest
	Public pageType
	Public mode
	Public isDisplayLoading
	Public strOriginalTableName
	Public strCaption
	Public shortTableName
	Public sessionPrefix
	Public tName
	Public conn
	Public gOrderIndexes
	Public gstrOrderBy
	Public gPageSize
	Public xt
	Public searchClauseObj
	Public needSearchClauseObj
	Public flyId
	Public includes_js
	Public includes_jsreq
	Public includes_css
	Public locale_info
	Public recId
	Public googleMapCfg
	Public menuTablesArr
	Public permis
	Public isGroupSecurity
	Public debugJSMode
	Public listAjax
	Public body
	Public onLoadJsEventCode
	Public masterTable
	Public useDetailsPreview
	Public allDetailsTablesArr
	Public jsSettings
	Public controlsHTMLMap
	Public controlsMap
	Public settingsMap
	Public pageAddLikeInline
	Public pageEditLikeInline
	Public arrAddTabs
	Public useTabsOnAdd
	Public arrEditTabs
	Public useTabsOnEdit
	Public arrViewTabs
	Public useTabsOnView
	Public arrRecsPerPage
	Public arrGroupsPerPage
	Public reportGroupFields
	Public pageSize
	Public isTableType
	Public eventsObject
	Public detailKeysByM
	Public isCaptchaOk
	Public captchaId
	Public lockingObj
	Public isUseVideo
	Public isUseAudio
	Public showAddInPopup
	Public showEditInPopup
	Public showViewInPopup
	Public isResizeColumns
	Public isUseCK
	Public isShowDetailTables
	Public filesToSave
	Public filesToMove
	Public filesToDelete
	Public Function init_RunnerPage_p1(ByRef params)
		DoAssignmentByRef init_RunnerPage_p1,method_RunnerPage_RunnerPage(me,params)
	End Function
	Public Function clearSessionKeys()
		DoAssignmentByRef clearSessionKeys,method_RunnerPage_clearSessionKeys(me)
	End Function
	Public Function setSessionVariables()
		DoAssignmentByRef setSessionVariables,method_RunnerPage_setSessionVariables(me)
	End Function
	Public Function addLookupSettings()
		DoAssignmentByRef addLookupSettings,method_RunnerPage_addLookupSettings(me)
	End Function
	Public Function fillGlobalSettings()
		DoAssignmentByRef fillGlobalSettings,method_RunnerPage_fillGlobalSettings(me)
	End Function
	Public Function fillTableSettings()
		DoAssignmentByRef fillTableSettings,method_RunnerPage_fillTableSettings(me)
	End Function
	Public Function fillFieldSettings()
		DoAssignmentByRef fillFieldSettings,method_RunnerPage_fillFieldSettings(me)
	End Function
	Public Function matchWithDetailKeys_p1(ByVal fName)
		DoAssignmentByRef matchWithDetailKeys_p1,method_RunnerPage_matchWithDetailKeys(me,fName)
	End Function
	Public Function fillPreload_p2(ByVal fName,ByVal value)
		DoAssignmentByRef fillPreload_p2,method_RunnerPage_fillPreload(me,fName,value)
	End Function
	Public Function hasDependField_p1(ByVal fName)
		DoAssignmentByRef hasDependField_p1,method_RunnerPage_hasDependField(me,fName)
	End Function
	Public Function getPreloadArr_p2(ByVal fName,ByVal value)
		DoAssignmentByRef getPreloadArr_p2,method_RunnerPage_getPreloadArr(me,fName,value)
	End Function
	Public Function getParentCtrlName_p1(ByVal fName)
		DoAssignmentByRef getParentCtrlName_p1,method_RunnerPage_getParentCtrlName(me,fName)
	End Function
	Public Function getParentVal_p1(ByVal fName)
		DoAssignmentByRef getParentVal_p1,method_RunnerPage_getParentVal(me,fName)
	End Function
	Public Function getSearchPreloadArr_p2(ByVal fName,ByVal fval)
		DoAssignmentByRef getSearchPreloadArr_p2,method_RunnerPage_getSearchPreloadArr(me,fName,fval)
	End Function
	Public Function addConfirmToFieldSettings()
		DoAssignmentByRef addConfirmToFieldSettings,method_RunnerPage_addConfirmToFieldSettings(me)
	End Function
	Public Function addCaptchaToFieldSettings()
		DoAssignmentByRef addCaptchaToFieldSettings,method_RunnerPage_addCaptchaToFieldSettings(me)
	End Function
	Public Function getStopLoading()
		DoAssignmentByRef getStopLoading,method_RunnerPage_getStopLoading(me)
	End Function
	Public Function fillValidation_p3(ByVal fData,ByVal val,ByRef arrSetVals)
		DoAssignmentByRef fillValidation_p3,method_RunnerPage_fillValidation(me,fData,val,arrSetVals)
	End Function
	Public Function getFieldsByPageType()
		DoAssignmentByRef getFieldsByPageType,method_RunnerPage_getFieldsByPageType(me)
	End Function
	Public Function fillSettings()
		DoAssignmentByRef fillSettings,method_RunnerPage_fillSettings(me)
	End Function
	Public Function fillFieldToolTips_p1(ByVal fName)
		DoAssignmentByRef fillFieldToolTips_p1,method_RunnerPage_fillFieldToolTips(me,fName)
	End Function
	Public Function fillControlsMap_p3(ByVal arr,ByVal addSet,ByVal fName)
		DoAssignmentByRef fillControlsMap_p3,method_RunnerPage_fillControlsMap(me,arr,addSet,fName)
	End Function
	Public Function fillControlsMap_p2(ByVal arr,ByVal addSet)
		DoAssignmentByRef fillControlsMap_p2,method_RunnerPage_fillControlsMap(me,arr,addSet,"")
	End Function
	Public Function fillControlsMap_p1(ByVal arr)
		DoAssignmentByRef fillControlsMap_p1,method_RunnerPage_fillControlsMap(me,arr,false,"")
	End Function
	Public Function fillControlsHtmlMap()
		DoAssignmentByRef fillControlsHtmlMap,method_RunnerPage_fillControlsHtmlMap(me)
	End Function
	Public Function fillSetCntrlMaps()
		DoAssignmentByRef fillSetCntrlMaps,method_RunnerPage_fillSetCntrlMaps(me)
	End Function
	Public Function fillCntrlTabGroups()
		DoAssignmentByRef fillCntrlTabGroups,method_RunnerPage_fillCntrlTabGroups(me)
	End Function
	Public Function isAppearOnTabs_p1(ByVal fName)
		DoAssignmentByRef isAppearOnTabs_p1,method_RunnerPage_isAppearOnTabs(me,fName)
	End Function
	Public Function getArrTabs()
		DoAssignmentByRef getArrTabs,method_RunnerPage_getArrTabs(me)
	End Function
	Public Function fillTimePickSettings_p2(ByVal field,ByVal value)
		DoAssignmentByRef fillTimePickSettings_p2,method_RunnerPage_fillTimePickSettings(me,field,value)
	End Function
	Public Function fillTimePickSettings_p1(ByVal field)
		DoAssignmentByRef fillTimePickSettings_p1,method_RunnerPage_fillTimePickSettings(me,field,"")
	End Function
	Public Function assignBodyEnd_p1(ByRef params)
		DoAssignmentByRef assignBodyEnd_p1,method_RunnerPage_assignBodyEnd(me,params)
	End Function
	Public Function genId()
		DoAssignmentByRef genId,method_RunnerPage_genId(me)
	End Function
	Public Function getPageType()
		DoAssignmentByRef getPageType,method_RunnerPage_getPageType(me)
	End Function
	Public Function AddJSCode_p1(ByVal jscode)
		DoAssignmentByRef AddJSCode_p1,method_RunnerPage_AddJSCode(me,jscode)
	End Function
	Public Function AddJSFileNoExt_p1(ByVal file)
		DoAssignmentByRef AddJSFileNoExt_p1,method_RunnerPage_AddJSFileNoExt(me,file)
	End Function
	Public Function AddJSFile_p4(ByVal file,ByVal req1,ByVal req2,ByVal req3)
		DoAssignmentByRef AddJSFile_p4,method_RunnerPage_AddJSFile(me,file,req1,req2,req3)
	End Function
	Public Function AddJSFile_p3(ByVal file,ByVal req1,ByVal req2)
		DoAssignmentByRef AddJSFile_p3,method_RunnerPage_AddJSFile(me,file,req1,req2,"")
	End Function
	Public Function AddJSFile_p2(ByVal file,ByVal req1)
		DoAssignmentByRef AddJSFile_p2,method_RunnerPage_AddJSFile(me,file,req1,"","")
	End Function
	Public Function AddJSFile_p1(ByVal file)
		DoAssignmentByRef AddJSFile_p1,method_RunnerPage_AddJSFile(me,file,"","","")
	End Function
	Public Function grabAllJsFiles()
		DoAssignmentByRef grabAllJsFiles,method_RunnerPage_grabAllJsFiles(me)
	End Function
	Public Function copyAllJsFiles_p1(ByVal jsFiles)
		DoAssignmentByRef copyAllJsFiles_p1,method_RunnerPage_copyAllJsFiles(me,jsFiles)
	End Function
	Public Function AddCSSFile_p1(ByVal file)
		DoAssignmentByRef AddCSSFile_p1,method_RunnerPage_AddCSSFile(me,file)
	End Function
	Public Function grabAllCssFiles()
		DoAssignmentByRef grabAllCssFiles,method_RunnerPage_grabAllCssFiles(me)
	End Function
	Public Function copyAllCssFiles_p1(ByVal cssFiles)
		DoAssignmentByRef copyAllCssFiles_p1,method_RunnerPage_copyAllCssFiles(me,cssFiles)
	End Function
	Public Function LoadJS_CSS()
		DoAssignmentByRef LoadJS_CSS,method_RunnerPage_LoadJS_CSS(me)
	End Function
	Public Function initForDetailsPreview()
		DoAssignmentByRef initForDetailsPreview,method_RunnerPage_initForDetailsPreview(me)
	End Function
	Public Function setLangParams()
		DoAssignmentByRef setLangParams,method_RunnerPage_setLangParams(me)
	End Function
	Public Function addCommonJs()
		DoAssignmentByRef addCommonJs,method_RunnerPage_addCommonJs(me)
	End Function
	Public Function PrepareJs()
		DoAssignmentByRef PrepareJs,method_RunnerPage_PrepareJs(me)
	End Function
	Public Function addButtonHandlers()
		DoAssignmentByRef addButtonHandlers,method_RunnerPage_addButtonHandlers(me)
	End Function
	Public Function addOnLoadJsEvent_p1(ByVal code)
		DoAssignmentByRef addOnLoadJsEvent_p1,method_RunnerPage_addOnLoadJsEvent(me,code)
	End Function
	Public Function setGoogleMapsParams_p1(ByVal fieldsArr)
		DoAssignmentByRef setGoogleMapsParams_p1,method_RunnerPage_setGoogleMapsParams(me,fieldsArr)
	End Function
	Public Function addBigGoogleMapMarkers_p2(ByRef data,ByVal viewLink)
		DoAssignmentByRef addBigGoogleMapMarkers_p2,method_RunnerPage_addBigGoogleMapMarkers(me,data,viewLink)
	End Function
	Public Function addBigGoogleMapMarkers_p1(ByRef data)
		DoAssignmentByRef addBigGoogleMapMarkers_p1,method_RunnerPage_addBigGoogleMapMarkers(me,data,"")
	End Function
	Public Function addGoogleMapData_p3(ByVal fName,ByRef data,ByVal viewLink)
		DoAssignmentByRef addGoogleMapData_p3,method_RunnerPage_addGoogleMapData(me,fName,data,viewLink)
	End Function
	Public Function addGoogleMapData_p2(ByVal fName,ByRef data)
		DoAssignmentByRef addGoogleMapData_p2,method_RunnerPage_addGoogleMapData(me,fName,data,"")
	End Function
	Public Function initGmaps()
		DoAssignmentByRef initGmaps,method_RunnerPage_initGmaps(me)
	End Function
	Public Function addCenterLink_p2(ByRef value,ByVal fName)
		DoAssignmentByRef addCenterLink_p2,method_RunnerPage_addCenterLink(me,value,fName)
	End Function
	Public Function createOldMenu()
		DoAssignmentByRef createOldMenu,method_RunnerPage_createOldMenu(me)
	End Function
	Public Function getPermissions_p1(ByVal tName)
		DoAssignmentByRef getPermissions_p1,method_RunnerPage_getPermissions(me,tName)
	End Function
	Public Function getPermissions()
		DoAssignmentByRef getPermissions,method_RunnerPage_getPermissions(me,"")
	End Function
	Public Function isCreateMenu()
		DoAssignmentByRef isCreateMenu,method_RunnerPage_isCreateMenu(me)
	End Function
	Public Function addRunLoading()
		DoAssignmentByRef addRunLoading,method_RunnerPage_addRunLoading(me)
	End Function
	Public Function eventExists_p1(ByVal name)
		DoAssignmentByRef eventExists_p1,method_RunnerPage_eventExists(me,name)
	End Function
	Public Function captchaExists()
		DoAssignmentByRef captchaExists,method_RunnerPage_captchaExists(me)
	End Function
	Public Function getNextPrevRecordKeys_p4(ByRef data,ByVal securityMode,ByRef var_next,ByRef prev)
		DoAssignmentByRef getNextPrevRecordKeys_p4,method_RunnerPage_getNextPrevRecordKeys(me,data,securityMode,var_next,prev)
	End Function
	Public Function doCaptchaCode()
		DoAssignmentByRef doCaptchaCode,method_RunnerPage_doCaptchaCode(me)
	End Function
	Public Function createCaptcha_p1(ByVal params)
		DoAssignmentByRef createCaptcha_p1,method_RunnerPage_createCaptcha(me,params)
	End Function
	Public Function createPerPage()
		DoAssignmentByRef createPerPage,method_RunnerPage_createPerPage(me)
	End Function
	Public Function ProcessFiles()
		DoAssignmentByRef ProcessFiles,method_RunnerPage_ProcessFiles(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"id", id
		setArrElement out,"totalCode", totalCode
		setArrElement out,"calendar", calendar
		setArrElement out,"timepicker", timepicker
		setArrElement out,"useIbox", useIbox
		setArrElement out,"isUseToolTips", isUseToolTips
		setArrElement out,"isUseAjaxSuggest", isUseAjaxSuggest
		setArrElement out,"pageType", pageType
		setArrElement out,"mode", mode
		setArrElement out,"isDisplayLoading", isDisplayLoading
		setArrElement out,"strOriginalTableName", strOriginalTableName
		setArrElement out,"strCaption", strCaption
		setArrElement out,"shortTableName", shortTableName
		setArrElement out,"sessionPrefix", sessionPrefix
		setArrElement out,"tName", tName
		setArrElement out,"conn", conn
		setArrElement out,"gOrderIndexes", gOrderIndexes
		setArrElement out,"gstrOrderBy", gstrOrderBy
		setArrElement out,"gPageSize", gPageSize
		setArrElement out,"xt", xt
		setArrElement out,"searchClauseObj", searchClauseObj
		setArrElement out,"needSearchClauseObj", needSearchClauseObj
		setArrElement out,"flyId", flyId
		setArrElement out,"includes_js", includes_js
		setArrElement out,"includes_jsreq", includes_jsreq
		setArrElement out,"includes_css", includes_css
		setArrElement out,"locale_info", locale_info
		setArrElement out,"recId", recId
		setArrElement out,"googleMapCfg", googleMapCfg
		setArrElement out,"menuTablesArr", menuTablesArr
		setArrElement out,"permis", permis
		setArrElement out,"isGroupSecurity", isGroupSecurity
		setArrElement out,"debugJSMode", debugJSMode
		setArrElement out,"listAjax", listAjax
		setArrElement out,"body", body
		setArrElement out,"onLoadJsEventCode", onLoadJsEventCode
		setArrElement out,"masterTable", masterTable
		setArrElement out,"useDetailsPreview", useDetailsPreview
		setArrElement out,"allDetailsTablesArr", allDetailsTablesArr
		setArrElement out,"jsSettings", jsSettings
		setArrElement out,"controlsHTMLMap", controlsHTMLMap
		setArrElement out,"controlsMap", controlsMap
		setArrElement out,"settingsMap", settingsMap
		setArrElement out,"pageAddLikeInline", pageAddLikeInline
		setArrElement out,"pageEditLikeInline", pageEditLikeInline
		setArrElement out,"arrAddTabs", arrAddTabs
		setArrElement out,"useTabsOnAdd", useTabsOnAdd
		setArrElement out,"arrEditTabs", arrEditTabs
		setArrElement out,"useTabsOnEdit", useTabsOnEdit
		setArrElement out,"arrViewTabs", arrViewTabs
		setArrElement out,"useTabsOnView", useTabsOnView
		setArrElement out,"arrRecsPerPage", arrRecsPerPage
		setArrElement out,"arrGroupsPerPage", arrGroupsPerPage
		setArrElement out,"reportGroupFields", reportGroupFields
		setArrElement out,"pageSize", pageSize
		setArrElement out,"isTableType", isTableType
		setArrElement out,"eventsObject", eventsObject
		setArrElement out,"detailKeysByM", detailKeysByM
		setArrElement out,"isCaptchaOk", isCaptchaOk
		setArrElement out,"captchaId", captchaId
		setArrElement out,"lockingObj", lockingObj
		setArrElement out,"isUseVideo", isUseVideo
		setArrElement out,"isUseAudio", isUseAudio
		setArrElement out,"showAddInPopup", showAddInPopup
		setArrElement out,"showEditInPopup", showEditInPopup
		setArrElement out,"showViewInPopup", showViewInPopup
		setArrElement out,"isResizeColumns", isResizeColumns
		setArrElement out,"isUseCK", isUseCK
		setArrElement out,"isShowDetailTables", isShowDetailTables
		setArrElement out,"filesToSave", filesToSave
		setArrElement out,"filesToMove", filesToMove
		setArrElement out,"filesToDelete", filesToDelete
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment id, dict("id")
		DoAssignment totalCode, dict("totalCode")
		DoAssignment calendar, dict("calendar")
		DoAssignment timepicker, dict("timepicker")
		DoAssignment useIbox, dict("useIbox")
		DoAssignment isUseToolTips, dict("isUseToolTips")
		DoAssignment isUseAjaxSuggest, dict("isUseAjaxSuggest")
		DoAssignment pageType, dict("pageType")
		DoAssignment mode, dict("mode")
		DoAssignment isDisplayLoading, dict("isDisplayLoading")
		DoAssignment strOriginalTableName, dict("strOriginalTableName")
		DoAssignment strCaption, dict("strCaption")
		DoAssignment shortTableName, dict("shortTableName")
		DoAssignment sessionPrefix, dict("sessionPrefix")
		DoAssignment tName, dict("tName")
		DoAssignment conn, dict("conn")
		DoAssignment gOrderIndexes, dict("gOrderIndexes")
		DoAssignment gstrOrderBy, dict("gstrOrderBy")
		DoAssignment gPageSize, dict("gPageSize")
		DoAssignment xt, dict("xt")
		DoAssignment searchClauseObj, dict("searchClauseObj")
		DoAssignment needSearchClauseObj, dict("needSearchClauseObj")
		DoAssignment flyId, dict("flyId")
		DoAssignment includes_js, dict("includes_js")
		DoAssignment includes_jsreq, dict("includes_jsreq")
		DoAssignment includes_css, dict("includes_css")
		DoAssignment locale_info, dict("locale_info")
		DoAssignment recId, dict("recId")
		DoAssignment googleMapCfg, dict("googleMapCfg")
		DoAssignment menuTablesArr, dict("menuTablesArr")
		DoAssignment permis, dict("permis")
		DoAssignment isGroupSecurity, dict("isGroupSecurity")
		DoAssignment debugJSMode, dict("debugJSMode")
		DoAssignment listAjax, dict("listAjax")
		DoAssignment body, dict("body")
		DoAssignment onLoadJsEventCode, dict("onLoadJsEventCode")
		DoAssignment masterTable, dict("masterTable")
		DoAssignment useDetailsPreview, dict("useDetailsPreview")
		DoAssignment allDetailsTablesArr, dict("allDetailsTablesArr")
		DoAssignment jsSettings, dict("jsSettings")
		DoAssignment controlsHTMLMap, dict("controlsHTMLMap")
		DoAssignment controlsMap, dict("controlsMap")
		DoAssignment settingsMap, dict("settingsMap")
		DoAssignment pageAddLikeInline, dict("pageAddLikeInline")
		DoAssignment pageEditLikeInline, dict("pageEditLikeInline")
		DoAssignment arrAddTabs, dict("arrAddTabs")
		DoAssignment useTabsOnAdd, dict("useTabsOnAdd")
		DoAssignment arrEditTabs, dict("arrEditTabs")
		DoAssignment useTabsOnEdit, dict("useTabsOnEdit")
		DoAssignment arrViewTabs, dict("arrViewTabs")
		DoAssignment useTabsOnView, dict("useTabsOnView")
		DoAssignment arrRecsPerPage, dict("arrRecsPerPage")
		DoAssignment arrGroupsPerPage, dict("arrGroupsPerPage")
		DoAssignment reportGroupFields, dict("reportGroupFields")
		DoAssignment pageSize, dict("pageSize")
		DoAssignment isTableType, dict("isTableType")
		DoAssignment eventsObject, dict("eventsObject")
		DoAssignment detailKeysByM, dict("detailKeysByM")
		DoAssignment isCaptchaOk, dict("isCaptchaOk")
		DoAssignment captchaId, dict("captchaId")
		DoAssignment lockingObj, dict("lockingObj")
		DoAssignment isUseVideo, dict("isUseVideo")
		DoAssignment isUseAudio, dict("isUseAudio")
		DoAssignment showAddInPopup, dict("showAddInPopup")
		DoAssignment showEditInPopup, dict("showEditInPopup")
		DoAssignment showViewInPopup, dict("showViewInPopup")
		DoAssignment isResizeColumns, dict("isResizeColumns")
		DoAssignment isUseCK, dict("isUseCK")
		DoAssignment isShowDetailTables, dict("isShowDetailTables")
		DoAssignment filesToSave, dict("filesToSave")
		DoAssignment filesToMove, dict("filesToMove")
		DoAssignment filesToDelete, dict("filesToDelete")
	End Sub
' end serialize
End Class
' RunnerPage implementation 
Function method_RunnerPage_RunnerPage(byref this_object,ByRef params)
	this_object.id = 1
	this_object.totalCode = ""
	this_object.calendar = false
	this_object.timepicker = false
	this_object.useIbox = false
	this_object.isUseToolTips = false
	this_object.isUseAjaxSuggest = true
	this_object.pageType = ""
	this_object.mode = 0
	this_object.isDisplayLoading = false
	this_object.strOriginalTableName = ""
	this_object.strCaption = ""
	this_object.shortTableName = ""
	this_object.sessionPrefix = ""
	this_object.tName = ""
	this_object.conn = ""
	doClassAssignment this_object,"gOrderIndexes",CreateDictionary()
	this_object.gstrOrderBy = ""
	this_object.gPageSize = 0
	this_object.xt = null
	this_object.searchClauseObj = null
	this_object.needSearchClauseObj = true
	this_object.flyId = 1
	doClassAssignment this_object,"includes_js",CreateDictionary()
	doClassAssignment this_object,"includes_jsreq",CreateDictionary()
	doClassAssignment this_object,"includes_css",CreateDictionary()
	doClassAssignment this_object,"locale_info",CreateDictionary()
	this_object.recId = 0
	doClassAssignment this_object,"googleMapCfg",CreateDictionary()
	doClassAssignment this_object,"menuTablesArr",CreateDictionary()
	doClassAssignment this_object,"permis",CreateDictionary()
	this_object.isGroupSecurity = true
	this_object.debugJSMode = false
	this_object.listAjax = false
	doClassAssignment this_object,"body",CreateDictionary()
	this_object.onLoadJsEventCode = ""
	this_object.masterTable = ""
	this_object.useDetailsPreview = false
	doClassAssignment this_object,"allDetailsTablesArr",CreateDictionary()
	doClassAssignment this_object,"jsSettings",CreateDictionary()
	doClassAssignment this_object,"controlsHTMLMap",CreateDictionary()
	doClassAssignment this_object,"controlsMap",CreateDictionary()
	doClassAssignment this_object,"settingsMap",CreateDictionary()
	this_object.pageAddLikeInline = false
	this_object.pageEditLikeInline = false
	doClassAssignment this_object,"arrAddTabs",CreateDictionary()
	this_object.useTabsOnAdd = false
	doClassAssignment this_object,"arrEditTabs",CreateDictionary()
	this_object.useTabsOnEdit = false
	doClassAssignment this_object,"arrViewTabs",CreateDictionary()
	this_object.useTabsOnView = false
	doClassAssignment this_object,"arrRecsPerPage",CreateDictionary()
	doClassAssignment this_object,"arrGroupsPerPage",CreateDictionary()
	this_object.reportGroupFields = false
	this_object.pageSize = 0
	this_object.isTableType = ""
	doClassAssignment this_object,"detailKeysByM",CreateDictionary()
	this_object.isCaptchaOk = true
	this_object.captchaId = ""
	this_object.lockingObj = null
	this_object.isUseVideo = false
	this_object.isUseAudio = false
	this_object.showAddInPopup = false
	this_object.showEditInPopup = false
	this_object.showViewInPopup = false
	this_object.isResizeColumns = false
	this_object.isUseCK = false
	this_object.isShowDetailTables = false
	doClassAssignment this_object,"filesToSave",CreateDictionary()
	doClassAssignment this_object,"filesToMove",CreateDictionary()
	doClassAssignment this_object,"filesToDelete",CreateDictionary()
	Dim this,i
	RunnerApply this_object,params
	this_object.debugJSMode = false
	if IsLess(this_object.flyId,CSmartDbl(this_object.id)+1) then
		this_object.flyId = CSmartDbl(this_object.id)+1
	end if
	if bValue(this_object.tName) then
		setArrElement this_object.permis,this_object.tName,this_object.getPermissions()
		doClassAssignmentByRef this_object,"eventsObject",getEventObject(this_object.tName)
	end if
	i = 0
	do while IsLess(i,asp_count(this_object.menuTablesArr))
		setArrElement this_object.permis,ArrayElement(ArrayElement(this_object.menuTablesArr,i),"dataSourceTName"),this_object.getPermissions_p1(ArrayElement(ArrayElement(this_object.menuTablesArr,i),"dataSourceTName"))
		i = CSmartDbl(i)+1
	loop
	if not bValue(this_object.sessionPrefix) then
		doClassAssignment this_object,"sessionPrefix",this_object.tName
	end if
	this_object.setSessionVariables 
	doClassAssignment this_object,"lockingObj",GetLockingObject(this_object.tName)
	doClassAssignment this_object,"detailKeysByM",GetDetailKeysByMasterTable(this_object.masterTable,this_object.tName,CreateDictionary())
	doClassAssignment this_object,"allDetailsTablesArr",GetDetailTablesArr(this_object.tName)
	doClassAssignment this_object,"strCaption",GetTableCaption(GoodFieldName(this_object.tName))
	doClassAssignment this_object,"isTableType",GetTableData(this_object.tName,".isTableType","")
	doClassAssignment this_object,"isUseVideo",GetTableData(this_object.tName,".isUseVideo",false)
	doClassAssignment this_object,"isUseAudio",GetTableData(this_object.tName,".isUseAudio",false)
	doClassAssignment this_object,"showAddInPopup",GetTableData(this_object.tName,".showAddInPopup",false)
	doClassAssignment this_object,"showEditInPopup",GetTableData(this_object.tName,".showEditInPopup",false)
	doClassAssignment this_object,"showViewInPopup",GetTableData(this_object.tName,".showViewInPopup",false)
	doClassAssignment this_object,"isResizeColumns",GetTableData(this_object.tName,".isResizeColumns",false)
	doClassAssignment this_object,"isUseAjaxSuggest",GetTableData(this_object.tName,".isUseAjaxSuggest",false)
	doClassAssignment this_object,"useDetailsPreview",GetTableData(this_object.tName,".useDetailsPreview",false)
	doClassAssignment this_object,"isShowDetailTables",displayDetailsOn(this_object.tName,this_object.pageType)
	setArrElement this_object.jsSettings,"tableDefSettings",CreateDictionary()
	setArrElement this_object.jsSettings,"fieldDefSettings",CreateDictionary()
	setArrElementN this_object.jsSettings,CreateArray2("tableSettings",this_object.tName),CreateDictionary()
	setArrElementN this_object.jsSettings,CreateArray3("tableSettings",this_object.tName,"fieldSettings"),CreateDictionary()
	setArrElement this_object.settingsMap,"globalSettings",CreateDictionary()
	setArrElementN this_object.settingsMap,CreateArray2("globalSettings","ext"),"asp"
	setArrElementN this_object.settingsMap,CreateArray2("globalSettings","charSet"),cCharset
	setArrElementN this_object.settingsMap,CreateArray2("globalSettings","debugMode"),this_object.debugJSMode
	setArrElementN this_object.settingsMap,CreateArray3("globalSettings","shortTNames",this_object.tName),GetTableURL(this_object.tName)
	setArrElementN this_object.settingsMap,CreateArray3("globalSettings","locale","dateFormat"),ArrayElement(locale_info,"LOCALE_IDATE")
	setArrElementN this_object.settingsMap,CreateArray3("globalSettings","locale","dateDelimiter"),ArrayElement(locale_info,"LOCALE_SDATE")
	setArrElement this_object.settingsMap,"tableSettings",CreateDictionary()
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","isUseToolTips"),CreateDictionary2("default",false,"jsName","isUseToolTips")
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","useIbox"),CreateDictionary2("default",false,"jsName","isUseIbox")
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","listIcons"),CreateDictionary2("default",false,"jsName","listIcons")
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","strCaption"),CreateDictionary2("default","","jsName","strCaption")
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","isUseAudio"),CreateDictionary2("default",false,"jsName","isUseAudio")
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","isUseVideo"),CreateDictionary2("default",false,"jsName","isUseVideo")
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","rowHighlite"),CreateDictionary2("default",false,"jsName","isUseHighlite")
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","showAddInPopup"),CreateDictionary2("default",false,"jsName","showAddInPopup")
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","showEditInPopup"),CreateDictionary2("default",false,"jsName","showEditInPopup")
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","showViewInPopup"),CreateDictionary2("default",false,"jsName","showViewInPopup")
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","isResizeColumns"),CreateDictionary2("default",false,"jsName","isUseResize")
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","isUseAjaxSuggest"),CreateDictionary2("default",true,"jsName","ajaxSuggest")
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","useDetailsPreview"),CreateDictionary2("default",false,"jsName","isUseDP")
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","isUsebuttonHandlers"),CreateDictionary2("default",false,"jsName","isUseButtons")
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","isVerLayout"),CreateDictionary2("default",false,"jsName","isVertLayout")
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","recsPerRowList"),CreateDictionary2("default",1,"jsName","recsPerRowList")
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","UseTimestamp"),CreateDictionary2("default",false,"jsName","isUseTimeStamp")
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","strName"),CreateDictionary2("default","","jsName","strName")
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","ShowTime"),CreateDictionary2("default",false,"jsName","showTime")
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","EditFormat"),CreateDictionary2("default","","jsName","editFormat")
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","DateEditType"),CreateDictionary2("default",EDIT_DATE_SIMPLE,"jsName","dateEditType")
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","RTEType"),CreateDictionary2("default","","jsName","RTEType")
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","ViewFormat"),CreateDictionary2("default","","jsName","viewFormat")
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","validateAs"),CreateDictionary2("default",null,"jsName","validation")
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","LastYearFactor"),CreateDictionary2("default",10,"jsName","lastYear")
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","InitialYearFactor"),CreateDictionary2("default",100,"jsName","initialYear")
	setArrElementN this_object.jsSettings,CreateArray3("tableSettings",this_object.tName,"strCaption"),this_object.strCaption
	setArrElementN this_object.jsSettings,CreateArray3("tableSettings",this_object.tName,"pageMode"),this_object.mode
	if bValue(this_object.listAjax) then
		setArrElementN this_object.jsSettings,CreateArray3("tableSettings",this_object.tName,"pageMode"),LIST_AJAX
	end if
	if bValue(this_object.lockingObj) then
		setArrElementN this_object.jsSettings,CreateArray3("tableSettings",this_object.tName,"locking"),true
	end if
	if bValue(asp_count(this_object.allDetailsTablesArr)) then
		if IsEqual(this_object.pageType,PAGE_LIST) then
			setArrElementN this_object.jsSettings,CreateArray3("tableSettings",this_object.tName,"detailTables"),CreateDictionary()
		end if
		setArrElementN this_object.jsSettings,CreateArray3("tableSettings",this_object.tName,"isShowDetails"),this_object.isShowDetailTables
		i = 0
		do while IsLess(i,asp_count(this_object.allDetailsTablesArr))
			setArrElementN this_object.settingsMap,CreateArray3("globalSettings","shortTNames",ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dDataSourceTable")),ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dShortTable")
			if IsEqual(this_object.pageType,PAGE_LIST) then
				setArrElementN this_object.jsSettings,CreateArray4("tableSettings",this_object.tName,"detailTables",ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dDataSourceTable")),CreateDictionary2("dispChildCount",ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dispChildCount"),"hideChild",ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"hideChild"))
			end if
			i = CSmartDbl(i)+1
		loop
		if (IsEqual(this_object.pageType,PAGE_ADD) or IsEqual(this_object.pageType,PAGE_EDIT)) and bValue(this_object.isShowDetailTables) then
			setArrElement this_object.controlsMap,"dControlsMap",CreateDictionary()
		end if
	end if
	setArrElement this_object.controlsMap,"video",CreateDictionary()
	setArrElement this_object.controlsMap,"toolTips",CreateDictionary()
	this_object.addLookupSettings 
	if not IsEqual(this_object.mode,LIST_DETAILS) then
		setArrElement this_object.controlsMap,"controls",CreateDictionary()
		if not bValue(this_object.pageAddLikeInline) and not bValue(this_object.pageEditLikeInline) then
			setArrElement this_object.controlsMap,"search",CreateDictionary()
			setArrElementN this_object.controlsMap,CreateArray2("search","searchBlocks"),CreateDictionary()
			setArrElementN this_object.controlsMap,CreateArray2("search","panelSearchFields"),GetTableData(this_object.tName,".panelSearchFields",CreateDictionary())
			setArrElementN this_object.controlsMap,CreateArray2("search","allSearchFields"),GetTableData(this_object.tName,".allSearchFields",CreateDictionary())
			if bValue(this_object.searchClauseObj) then
				setArrElementN this_object.controlsMap,CreateArray2("search","usedSrch"),this_object.searchClauseObj.isUsedSrch()
			end if
			if not IsEqual(this_object.pageType,PAGE_SEARCH) then
				setArrElementN this_object.controlsMap,CreateArray2("search","submitPageType"),this_object.pageType
			else
				if bValue(postvalue("rname")) then
					setArrElementN this_object.controlsMap,CreateArray2("search","submitPageType"),"dreport"
					setArrElementN this_object.controlsMap,CreateArray3("search","baseParams","rname"),postvalue("rname")
				else
					if bValue(postvalue("cname")) then
						setArrElementN this_object.controlsMap,CreateArray2("search","submitPageType"),"dchart"
						setArrElementN this_object.controlsMap,CreateArray3("search","baseParams","cname"),postvalue("cname")
					else
						setArrElementN this_object.controlsMap,CreateArray2("search","submitPageType"),this_object.isTableType
					end if
				end if
			end if
		end if
	end if
	this_object.calendar = bValue(this_object.calendar) or bValue(GetTableData(this_object.tName,".isUseCalendarForSearch",false))
	this_object.timepicker = bValue(this_object.timepicker) or bValue(GetTableData(this_object.tName,".isUseTimeForSearch",false))
	this_object.isUseToolTips = bValue(this_object.isUseToolTips) or bValue(GetTableData(this_object.tName,".isUseToolTips",false))
	setArrElement this_object.googleMapCfg,"APIcode",""
End Function
Function method_RunnerPage_clearSessionKeys(byref this_object)
	Dim var_POST,var_GET,sess_unset,var_SESSION,key
	if (IsEqual(this_object.pageType,PAGE_LIST) and not bValue(asp_count(RequestForm()))) and (not bValue(asp_count(Request.QueryString)) or IsEqual(asp_count(Request.QueryString),1) and not IsEmpty(GetRequestValue(Request.QueryString,"menuItemId"))) then
		Set sess_unset = (CreateDictionary())
		GetCollectionBounds Session,c_runnerpage_loopIdx5,c_runnerpage_loopMax5
		do while c_runnerpage_loopIdx5<=c_runnerpage_loopMax5
			key = GetCollectionKey(Session,c_runnerpage_loopIdx5)
			doAssignment value,Session(key)
			if IsEqual(asp_substr(key,0,CSmartDbl(asp_strlen(this_object.tName))+1),CSmartStr(this_object.tName) & "_") and IsFalse(asp_strpos(asp_substr(key,CSmartDbl(asp_strlen(this_object.tName))+1,empty),"_",empty)) then
				setArrElement sess_unset,asp_count(sess_unset),key
			end if
			c_runnerpage_loopIdx5=c_runnerpage_loopIdx5+1
		loop
		GetCollectionBounds sess_unset,c_runnerpage_loopIdx6,c_runnerpage_loopMax6
		do while c_runnerpage_loopIdx6<=c_runnerpage_loopMax6
			c_runnerpage_arrKey6 = GetCollectionKey(sess_unset,c_runnerpage_loopIdx6)
			doAssignment key,ArrayElement(sess_unset,c_runnerpage_arrKey6)
			asp_unsetElement Session,key
			c_runnerpage_loopIdx6=c_runnerpage_loopIdx6+1
		loop
	end if
End Function
Function method_RunnerPage_setSessionVariables(byref this_object)
	Dim this,var_SESSION,allSearchFields,params,var_REQUEST
	this_object.clearSessionKeys 
	if not IsEqual(this_object.masterTable,"") then
		setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_mastertable",this_object.masterTable
	else
		doClassAssignment this_object,"masterTable",Session(CSmartStr(this_object.sessionPrefix) & "_mastertable")
	end if
	doAssignmentByRef allSearchFields,GetTableData(this_object.tName,".allSearchFields",CreateDictionary())
	if bValue(this_object.needSearchClauseObj) and not bValue(this_object.searchClauseObj) then
		if not IsEmpty(Session(CSmartStr(this_object.sessionPrefix) & "_advsearch")) then
			doClassAssignment this_object,"searchClauseObj",unserialize(Session(CSmartStr(this_object.sessionPrefix) & "_advsearch"))
		else
			Set params = (CreateDictionary())
			setArrElement params,"tName",this_object.tName
			setArrElement params,"searchFieldsArr",allSearchFields
			setArrElement params,"sessionPrefix",this_object.sessionPrefix
			setArrElement params,"panelSearchFields",GetTableData(this_object.tName,".panelSearchFields",CreateDictionary())
			setArrElement params,"googleLikeFields",GetTableData(this_object.tName,".googleLikeFields",CreateDictionary())
			doClassAssignment this_object,"searchClauseObj",CreateClass("SearchClause",1,params,Empty,Empty,Empty,Empty,Empty,Empty)
		end if
	end if
	if bValue(GetRequestValue(Request,"pagesize")) then
		setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_pagesize",GetRequestValue(Request,"pagesize")
		setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_pagenumber",1
	end if
	this_object.pageSize = CSmartLng(Session(CSmartStr(this_object.sessionPrefix) & "_pagesize"))
End Function
Function method_RunnerPage_addLookupSettings(byref this_object)
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","CategoryControl"),CreateDictionary2("default","","jsName","categoryField")
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","DependentLookups"),CreateDictionary2("default",CreateDictionary(),"jsName","depLookups")
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","LCType"),CreateDictionary2("default",LCT_DROPDOWN,"jsName","lcType")
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","LookupTable"),CreateDictionary2("default","","jsName","lookupTable")
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","SelectSize"),CreateDictionary2("default",1,"jsName","selectSize")
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","LinkField"),CreateDictionary2("default","","jsName","linkField")
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","DisplayField"),CreateDictionary2("default","","jsName","dispField")
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","freeInput"),CreateDictionary2("default",false,"jsName","freeInput")
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","autoCompleteFieldsAdd"),CreateDictionary2("default",CreateDictionary(),"jsName","autoCompleteFieldsAdd")
	setArrElementN this_object.settingsMap,CreateArray2("fieldSettings","autoCompleteFieldsEdit"),CreateDictionary2("default",CreateDictionary(),"jsName","autoCompleteFieldsEdit")
End Function
Function method_RunnerPage_fillGlobalSettings(byref this_object)
	Dim key,val
	setArrElement this_object.jsSettings,"global",CreateDictionary()
	GetCollectionBounds ArrayElement(this_object.settingsMap,"globalSettings"),c_runnerpage_loopIdx7,c_runnerpage_loopMax7
	do while c_runnerpage_loopIdx7<=c_runnerpage_loopMax7
		key = GetCollectionKey(ArrayElement(this_object.settingsMap,"globalSettings"),c_runnerpage_loopIdx7)
		doAssignment val,ArrayElement(ArrayElement(this_object.settingsMap,"globalSettings"),key)
		setArrElementN this_object.jsSettings,CreateArray2("global",key),val
		c_runnerpage_loopIdx7=c_runnerpage_loopIdx7+1
	loop
	setArrElementN this_object.jsSettings,CreateArray2("global","idStartFrom"),this_object.flyId
End Function
Function method_RunnerPage_fillTableSettings(byref this_object)
	Dim key,tData,val,isDefault
	GetCollectionBounds ArrayElement(this_object.settingsMap,"tableSettings"),c_runnerpage_loopIdx8,c_runnerpage_loopMax8
	do while c_runnerpage_loopIdx8<=c_runnerpage_loopMax8
		key = GetCollectionKey(ArrayElement(this_object.settingsMap,"tableSettings"),c_runnerpage_loopIdx8)
		doAssignment val,ArrayElement(ArrayElement(this_object.settingsMap,"tableSettings"),key)
		if (not IsEqual(key,"useIbox") and not IsEqual(this_object.pageType,PAGE_SEARCH) or IsEqual(key,"useIbox") and not IsEqual(this_object.pageType,PAGE_SEARCH)) or IsEqual(this_object.pageType,PAGE_SEARCH) and not IsEqual(key,"useIbox") then
			doAssignmentByRef tData,GetTableData(this_object.tName,"." & CSmartStr(key),ArrayElement(val,"default"))
			isDefault = false
			if bValue(asp_is_array(tData)) then
				isDefault = not bValue(asp_count(tData))
			else
				if not bValue(asp_is_array(ArrayElement(val,"default"))) then
					isDefault = IsEqual(tData,ArrayElement(val,"default"))
				end if
			end if
			if not bValue(isDefault) then
				setArrElementN this_object.jsSettings,CreateArray3("tableSettings",this_object.tName,ArrayElement(val,"jsName")),tData
			end if
		end if
		if not bValue(asp_array_key_exists(ArrayElement(val,"jsName"),ArrayElement(this_object.jsSettings,"tableDefSettings"))) then
			setArrElementN this_object.jsSettings,CreateArray2("tableDefSettings",ArrayElement(val,"jsName")),ArrayElement(val,"default")
		end if
		c_runnerpage_loopIdx8=c_runnerpage_loopIdx8+1
	loop
End Function
Function method_RunnerPage_fillFieldSettings(byref this_object)
	Dim arrFields,this,fName,matchDK,fData,key,val,isDefault
	doAssignmentByRef arrFields,this_object.getFieldsByPageType()
	GetCollectionBounds arrFields,c_runnerpage_loopIdx9,c_runnerpage_loopMax9
	do while c_runnerpage_loopIdx9<=c_runnerpage_loopMax9
		c_runnerpage_arrKey9 = GetCollectionKey(arrFields,c_runnerpage_loopIdx9)
		doAssignment fName,ArrayElement(arrFields,c_runnerpage_arrKey9)
		if not bValue(asp_array_key_exists(fName,ArrayElement(ArrayElement(ArrayElement(this_object.jsSettings,"tableSettings"),this_object.tName),"fieldSettings"))) then
			setArrElementN this_object.jsSettings,CreateArray4("tableSettings",this_object.tName,"fieldSettings",fName),CreateDictionary()
		end if
		matchDK = false
		if bValue(this_object.matchWithDetailKeys_p1(fName)) then
			matchDK = true
		end if
		GetCollectionBounds ArrayElement(this_object.settingsMap,"fieldSettings"),c_runnerpage_loopIdx10,c_runnerpage_loopMax10
		do while c_runnerpage_loopIdx10<=c_runnerpage_loopMax10
			key = GetCollectionKey(ArrayElement(this_object.settingsMap,"fieldSettings"),c_runnerpage_loopIdx10)
			doAssignment val,ArrayElement(ArrayElement(this_object.settingsMap,"fieldSettings"),key)
			doAssignmentByRef fData,GetFieldData(this_object.tName,fName,key,ArrayElement(val,"default"))
			if not IsEqual(key,"validateAs") then
				if IsEqual(key,"EditFormat") then
					if bValue(matchDK) then
						fData = EDIT_FORMAT_READONLY
					end if
				else
					if IsEqual(key,"RTEType") then
						if bValue(UseRTEBasic(fName,"")) then
							fData = "RTE"
						else
							if bValue(UseRTEFCK(fName,"")) then
								fData = "RTECK"
								if not bValue(this_object.isUseCK) then
									this_object.isUseCK = true
								end if
								setArrElementN this_object.jsSettings,CreateArray5("tableSettings",this_object.tName,"fieldSettings",fName,"nWidth"),GetNCols(fName,"")
								setArrElementN this_object.jsSettings,CreateArray5("tableSettings",this_object.tName,"fieldSettings",fName,"nHeight"),GetNRows(fName,"")
							else
								if bValue(UseRTEInnova(fName,"")) then
									fData = "RTEINNOVA"
								end if
							end if
						end if
					else
						if IsEqual(key,"autoCompleteFieldsAdd") then
							doAssignmentByRef fData,GetFieldData(this_object.tName,fName,"autoCompleteFields",ArrayElement(val,"default"))
						else
							if IsEqual(key,"autoCompleteFieldsEdit") then
								if bValue(GetFieldData(this_object.tName,fName,"autoCompleteFieldsOnEdit",false)) then
									doAssignmentByRef fData,GetFieldData(this_object.tName,fName,"autoCompleteFields",ArrayElement(val,"default"))
								else
									Set fData = (CreateDictionary())
								end if
							end if
						end if
					end if
				end if
				isDefault = false
				if bValue(asp_is_array(fData)) then
					isDefault = not bValue(asp_count(fData))
				else
					if not bValue(asp_is_array(ArrayElement(val,"default"))) then
						isDefault = IsIdentical(fData,ArrayElement(val,"default"))
					end if
				end if
				if not bValue(isDefault) and not bValue(matchDK) then
					setArrElementN this_object.jsSettings,CreateArray5("tableSettings",this_object.tName,"fieldSettings",fName,ArrayElement(val,"jsName")),fData
				else
					if bValue(matchDK) and (IsEqual(key,"EditFormat") or IsEqual(key,"strName")) then
						setArrElementN this_object.jsSettings,CreateArray5("tableSettings",this_object.tName,"fieldSettings",fName,ArrayElement(val,"jsName")),fData
					end if
				end if
				if not bValue(asp_array_key_exists(ArrayElement(val,"jsName"),ArrayElement(this_object.jsSettings,"fieldDefSettings"))) then
					setArrElementN this_object.jsSettings,CreateArray2("fieldDefSettings",ArrayElement(val,"jsName")),ArrayElement(val,"default")
				end if
			else
				if not bValue(matchDK) then
					this_object.fillValidation_p3 fData,val,ArrayElement(ArrayElement(ArrayElement(ArrayElement(this_object.jsSettings,"tableSettings"),this_object.tName),"fieldSettings"),fName)
				end if
			end if
			c_runnerpage_loopIdx10=c_runnerpage_loopIdx10+1
		loop
		if bValue(GetLookupTable(fName,this_object.tName)) then
			setArrElementN this_object.jsSettings,CreateArray3("global","shortTNames",GetLookupTable(fName,this_object.tName)),GetTableURL(GetLookupTable(fName,this_object.tName))
		end if
		if IsEqual(GetEditFormat(fName,""),"Time") then
			this_object.fillTimePickSettings_p1 fName
		end if
		c_runnerpage_loopIdx9=c_runnerpage_loopIdx9+1
	loop
	if IsEqual(this_object.pageType,PAGE_REGISTER) then
		this_object.addConfirmToFieldSettings 
	end if
End Function
Function method_RunnerPage_matchWithDetailKeys(byref this_object,ByVal fName)
	Dim match,j
	match = false
	if bValue(this_object.detailKeysByM) then
		j = 0
		do while IsLess(j,asp_count(this_object.detailKeysByM))
			if IsEqual(ArrayElement(this_object.detailKeysByM,j),fName) then
				match = true
				exit do
			end if
			j = CSmartDbl(j)+1
		loop
	end if
	doAssignmentByRef method_RunnerPage_matchWithDetailKeys,match
	Exit Function
End Function
Function method_RunnerPage_fillPreload(byref this_object,ByVal fName,ByVal value)
	Dim preload,this
	preload = false
	if not bValue(this_object.matchWithDetailKeys_p1(fName)) and bValue(UseCategory(fName,this_object.tName)) then
		if (IsEqual(this_object.pageType,PAGE_ADD) or IsEqual(this_object.pageType,PAGE_EDIT)) or IsEqual(this_object.pageType,PAGE_REGISTER) then
			doAssignmentByRef preload,this_object.getPreloadArr_p2(fName,value)
		else
			doAssignmentByRef preload,this_object.getSearchPreloadArr_p2(fName,value)
		end if
	end if
	doAssignmentByRef method_RunnerPage_fillPreload,preload
	Exit Function
End Function
Function method_RunnerPage_hasDependField(byref this_object,ByVal fName)
	if (not IsEqual(GetEditFormat(fName,""),EDIT_FORMAT_LOOKUP_WIZARD) or not IsEqual(GetEditFormat(fName,""),EDIT_FORMAT_RADIO)) and not bValue(UseCategory(fName,"")) then
		method_RunnerPage_hasDependField = false
		Exit Function
	end if
	doAssignmentByRef method_RunnerPage_hasDependField,CategoryControl(fName,this_object.tName)
	Exit Function
End Function
Function method_RunnerPage_getPreloadArr(byref this_object,ByVal fName,ByVal value)
	Dim strCategoryControl,this,fieldAppear,categoryFieldAppear,output,valF,fVal
	doAssignmentByRef strCategoryControl,this_object.hasDependField_p1(fName)
	if IsFalse(strCategoryControl) then
		method_RunnerPage_getPreloadArr = false
		Exit Function
	end if
	fieldAppear = true
	if IsEqual(this_object.pageType,PAGE_ADD) then
		if not bValue(AppearOnInlineAdd(fName,"")) then
			fieldAppear = not IsEqual(this_object.mode,ADD_INLINE)
		else
			if not bValue(AppearOnAddPage(fName,"")) then
				fieldAppear = IsEqual(this_object.mode,ADD_INLINE)
			end if
		end if
		if IsEqual(this_object.mode,ADD_INLINE) then
			if bValue(AppearOnInlineAdd(strCategoryControl,"")) then
				categoryFieldAppear = true
			else
				categoryFieldAppear = false
			end if
		else
			if bValue(AppearOnAddPage(strCategoryControl,"")) then
				categoryFieldAppear = true
			else
				categoryFieldAppear = false
			end if
		end if
	else
		if IsEqual(this_object.pageType,PAGE_EDIT) then
			if not bValue(AppearOnInlineEdit(fName,"")) then
				fieldAppear = not IsEqual(this_object.mode,EDIT_INLINE)
			else
				if not bValue(AppearOnEditPage(fName,"")) then
					fieldAppear = IsEqual(this_object.mode,EDIT_INLINE)
				end if
			end if
			categoryFieldAppear = true
		else
			if bValue(strCategoryControl) then
				categoryFieldAppear = true
			else
				categoryFieldAppear = false
			end if
		end if
	end if
	if not bValue(fieldAppear) then
		method_RunnerPage_getPreloadArr = false
		Exit Function
	end if
	if bValue(GetFieldData(this_object.tName,fName,"freeInput",false)) then
		Set output = (CreateDictionary2(0,ArrayElement(value,fName),1,ArrayElement(value,fName)))
	else
		doAssignmentByRef output,loadSelectContent(fName,ArrayElement(value,strCategoryControl),categoryFieldAppear,ArrayElement(value,fName),true)
	end if
	valF = ""
	if bValue(asp_count(value)) then
		doAssignment valF,ArrayElement(value,fName)
	end if
	if IsEqual(this_object.pageType,PAGE_EDIT) then
		if IsEqual(SelectSize(fName,""),1) and not IsEqual(LookupControlType(fName,""),LCT_CBLIST) then
			doAssignment fVal,valF
		else
			doAssignmentByRef fVal,splitvalues(valF)
		end if
	else
		doAssignment fVal,valF
	end if
	doAssignmentByRef method_RunnerPage_getPreloadArr,CreateDictionary2("vals",output,"fVal",fVal)
	Exit Function
End Function
Function method_RunnerPage_getParentCtrlName(byref this_object,ByVal fName)
	doAssignmentByRef method_RunnerPage_getParentCtrlName,CategoryControl(fName,this_object.tName)
	Exit Function
End Function
Function method_RunnerPage_getParentVal(byref this_object,ByVal fName)
	Dim categoryFieldParams
	doAssignmentByRef categoryFieldParams,this_object.searchClauseObj.getSearchCtrlParams_p1(this_object.getParentCtrlName_p1(fName))
	if bValue(asp_count(categoryFieldParams)) then
		doAssignmentByRef method_RunnerPage_getParentVal,ArrayElement(ArrayElement(categoryFieldParams,0),"value1")
		Exit Function
	else
		method_RunnerPage_getParentVal = false
		Exit Function
	end if
End Function
Function method_RunnerPage_getSearchPreloadArr(byref this_object,ByVal fName,ByVal fval)
	Dim parentFieldName,parentVal,this,doFilter,output
	if not IsEqual(GetEditFormat(fName,""),EDIT_FORMAT_LOOKUP_WIZARD) and not bValue(UseCategory(fName,this_object.tName)) then
		method_RunnerPage_getSearchPreloadArr = false
		Exit Function
	end if
	doAssignmentByRef parentFieldName,CategoryControl(fName,this_object.tName)
	doAssignmentByRef parentVal,this_object.getParentVal_p1(fName)
	doFilter = true
	if IsFalse(parentVal) or IsIdentical(parentVal,"") then
		doFilter = false
	end if
	doAssignmentByRef output,loadSelectContent(fName,parentVal,doFilter,ArrayElement(fval,fName),true)
	doAssignmentByRef method_RunnerPage_getSearchPreloadArr,CreateDictionary2("vals",output,"fVal",ArrayElement(fval,fName))
	Exit Function
End Function
Function method_RunnerPage_addConfirmToFieldSettings(byref this_object)
	Dim arrSetVals
	Set arrSetVals = (CreateDictionary())
	setArrElement arrSetVals,"strName","confirm"
	setArrElement arrSetVals,"EditFormat","Password"
	setArrElementN arrSetVals,CreateArray3("validation","validationArr",empty),"IsRequired"
	setArrElementN this_object.jsSettings,CreateArray4("tableSettings",this_object.tName,"fieldSettings","confirm"),arrSetVals
End Function
Function method_RunnerPage_addCaptchaToFieldSettings(byref this_object)
	Dim arrSetVals
	Set arrSetVals = (CreateDictionary())
	setArrElement arrSetVals,"strName","captcha"
	setArrElement arrSetVals,"EditFormat","Text Field"
	setArrElementN arrSetVals,CreateArray3("validation","validationArr",empty),"IsRequired"
	setArrElementN this_object.jsSettings,CreateArray4("tableSettings",this_object.tName,"fieldSettings","captcha"),arrSetVals
End Function
Function method_RunnerPage_getStopLoading(byref this_object)
	Dim this
	this_object.AddJsCode_p1 (((vblf & _
		"stopLoading(" & CSmartStr(this_object.id)) & ",") & CSmartStr(this_object.mode)) & ");"
End Function
Function method_RunnerPage_fillValidation(byref this_object,ByVal fData,ByVal val,ByRef arrSetVals)
	if bValue(asp_count(fData)) then
		setArrElementN arrSetVals,CreateArray2(ArrayElement(val,"jsName"),"validationArr"),ArrayElement(fData,"basicValidate")
		if bValue(asp_array_key_exists("regExp",fData)) then
			setArrElementN arrSetVals,CreateArray2(ArrayElement(val,"jsName"),"regExp"),ArrayElement(fData,"regExp")
		else
			if not bValue(asp_array_key_exists(ArrayElement(val,"jsName"),ArrayElement(this_object.jsSettings,"fieldDefSettings"))) or not bValue(asp_array_key_exists("RegExp",ArrayElement(ArrayElement(this_object.jsSettings,"fieldDefSettings"),ArrayElement(val,"jsName")))) then
				setArrElementN this_object.jsSettings,CreateArray3("fieldDefSettings",ArrayElement(val,"jsName"),"regExp"),null
			end if
		end if
	else
		if not bValue(asp_array_key_exists(ArrayElement(val,"jsName"),ArrayElement(this_object.jsSettings,"fieldDefSettings"))) then
			setArrElementN this_object.jsSettings,CreateArray3("fieldDefSettings",ArrayElement(val,"jsName"),"regExp"),null
			setArrElementN this_object.jsSettings,CreateArray3("fieldDefSettings",ArrayElement(val,"jsName"),"validationArr"),ArrayElement(val,"default")
		else
			if not bValue(asp_array_key_exists("validationArr",ArrayElement(ArrayElement(this_object.jsSettings,"fieldDefSettings"),ArrayElement(val,"jsName")))) then
				setArrElementN this_object.jsSettings,CreateArray3("fieldDefSettings",ArrayElement(val,"jsName"),"validationArr"),ArrayElement(val,"default")
			end if
		end if
	end if
End Function
Function method_RunnerPage_getFieldsByPageType(byref this_object)
	Dim arrFields,allfields,app,fName
	Set arrFields = (CreateDictionary())
	doAssignmentByRef allfields,GetFieldsList(this_object.tName)
	GetCollectionBounds allfields,c_runnerpage_loopIdx12,c_runnerpage_loopMax12
	c_runnerpage_exitLoop12=false
	do while c_runnerpage_loopIdx12<=c_runnerpage_loopMax12
		c_runnerpage_exitLoop12=false
		do
			c_runnerpage_arrKey12 = GetCollectionKey(allfields,c_runnerpage_loopIdx12)
			doAssignment fName,ArrayElement(allfields,c_runnerpage_arrKey12)
			if IsEqual(this_object.pageType,PAGE_EDIT) and bValue(this_object.pageEditLikeInline) then
				doAssignmentByRef app,AppearOnCurrentPage(fName,this_object.pageType,true)
			else
				if IsEqual(this_object.pageType,PAGE_ADD) and bValue(this_object.pageAddLikeInline) then
					doAssignmentByRef app,AppearOnCurrentPage(fName,this_object.pageType,true)
				else
					doAssignmentByRef app,AppearOnCurrentPage(fName,this_object.pageType,false)
				end if
			end if
			if IsIdentical(app,"break") then
				c_runnerpage_exitLoop12=true
				exit do
			else
				if not bValue(app) then
					exit do
				end if
			end if
			setArrElement arrFields,asp_count(arrFields),fName
		loop while false
		if c_runnerpage_exitLoop12 then _
			exit do
		c_runnerpage_loopIdx12=c_runnerpage_loopIdx12+1
	loop
	doAssignmentByRef method_RunnerPage_getFieldsByPageType,arrFields
	Exit Function
End Function
Function method_RunnerPage_fillSettings(byref this_object)
	Dim this
	this_object.fillGlobalSettings 
	this_object.fillTableSettings 
	this_object.fillFieldSettings 
End Function
Function method_RunnerPage_fillFieldToolTips(byref this_object,ByVal fName)
	Dim toolTipText
	doAssignmentByRef toolTipText,GetFieldToolTip(this_object.tName,fName)
	if bValue(asp_strlen(toolTipText)) then
		setArrElementN this_object.controlsMap,CreateArray2("toolTips",fName),toolTipText
	end if
End Function
Function method_RunnerPage_fillControlsMap(byref this_object,ByVal arr,ByVal addSet,ByVal fName)
	Dim key,val,vkey,vval,i
	if not bValue(addSet) then
		GetCollectionBounds arr,c_runnerpage_loopIdx13,c_runnerpage_loopMax13
		do while c_runnerpage_loopIdx13<=c_runnerpage_loopMax13
			key = GetCollectionKey(arr,c_runnerpage_loopIdx13)
			doAssignment val,ArrayElement(arr,key)
			setArrElementN this_object.controlsMap,CreateArray2(key,empty),val
			c_runnerpage_loopIdx13=c_runnerpage_loopIdx13+1
		loop
	else
		GetCollectionBounds arr,c_runnerpage_loopIdx14,c_runnerpage_loopMax14
		do while c_runnerpage_loopIdx14<=c_runnerpage_loopMax14
			key = GetCollectionKey(arr,c_runnerpage_loopIdx14)
			doAssignment val,ArrayElement(arr,key)
			GetCollectionBounds val,c_runnerpage_loopIdx15,c_runnerpage_loopMax15
			do while c_runnerpage_loopIdx15<=c_runnerpage_loopMax15
				vkey = GetCollectionKey(val,c_runnerpage_loopIdx15)
				doAssignment vval,ArrayElement(val,vkey)
				if not bValue(fName) then
					setArrElementN this_object.controlsMap,CreateArray3(key,CSmartDbl(asp_count(ArrayElement(this_object.controlsMap,key)))-1,vkey),vval
				else
					i = 0
					do while IsLess(i,asp_count(ArrayElement(this_object.controlsMap,key)))
						if IsEqual(ArrayElement(ArrayElement(ArrayElement(this_object.controlsMap,key),i),"fieldName"),fName) then
							setArrElementN this_object.controlsMap,CreateArray3(key,i,vkey),vval
							exit do
						end if
						i = CSmartDbl(i)+1
					loop
				end if
				c_runnerpage_loopIdx15=c_runnerpage_loopIdx15+1
			loop
			c_runnerpage_loopIdx14=c_runnerpage_loopIdx14+1
		loop
	end if
End Function
Function method_RunnerPage_fillControlsHtmlMap(byref this_object)
	Dim key,val
	setArrElement this_object.controlsHTMLMap,this_object.tName,CreateDictionary()
	setArrElementN this_object.controlsHTMLMap,CreateArray2(this_object.tName,this_object.pageType),CreateDictionary()
	setArrElementN this_object.controlsHTMLMap,CreateArray3(this_object.tName,this_object.pageType,this_object.id),CreateDictionary()
	setArrElement this_object.controlsMap,"gMaps",this_object.googleMapCfg
	GetCollectionBounds this_object.controlsMap,c_runnerpage_loopIdx17,c_runnerpage_loopMax17
	do while c_runnerpage_loopIdx17<=c_runnerpage_loopMax17
		key = GetCollectionKey(this_object.controlsMap,c_runnerpage_loopIdx17)
		doAssignment val,ArrayElement(this_object.controlsMap,key)
		setArrElementN this_object.controlsHTMLMap,CreateArray4(this_object.tName,this_object.pageType,this_object.id,key),val
		c_runnerpage_loopIdx17=c_runnerpage_loopIdx17+1
	loop
End Function
Function method_RunnerPage_fillSetCntrlMaps(byref this_object)
	Dim this
	this_object.fillSettings 
	this_object.fillControlsHTMLMap 
End Function
Function method_RunnerPage_fillCntrlTabGroups(byref this_object)
	Dim arrTabs,this,beginGroup,tabGroupName,i,tabC,tabN,tabsAndFields,j,f
	doAssignmentByRef arrTabs,this_object.getArrTabs()
	setArrElement this_object.controlsMap,"tabs",CreateDictionary()
	setArrElement this_object.controlsMap,"sections",CreateDictionary()
	if not bValue(arrTabs) then
		method_RunnerPage_fillCntrlTabGroups = false
		Exit Function
	end if
	beginGroup = false
	tabGroupName = ""
	i = 0
	do while IsLess(i,asp_count(arrTabs))
		doAssignment tabC,ArrayElement(arrTabs,i)
		doAssignment tabN,IIF(IsLess(CSmartDbl(i)+1,asp_count(arrTabs)),ArrayElement(arrTabs,CSmartDbl(i)+1),false)
		if not bValue(ArrayElement(tabC,"nType")) then
			if not bValue(beginGroup) then
				beginGroup = true
				doAssignment tabGroupName,ArrayElement(tabC,"tabId")
			end if
			if bValue(beginGroup) then
				if (not bValue(tabN) or bValue(ArrayElement(tabN,"nType"))) or not IsEqual(ArrayElement(tabN,"tabGroup"),ArrayElement(tabC,"tabGroup")) then
					Set tabsAndFields = (CreateDictionary())
					j = 0
					do while IsLess(j,asp_count(arrTabs))
						if IsEqual(ArrayElement(ArrayElement(arrTabs,j),"tabGroup"),ArrayElement(tabC,"tabGroup")) then
							setArrElement tabsAndFields,ArrayElement(ArrayElement(arrTabs,j),"tabId"),CreateDictionary()
							f = 0
							do while IsLess(f,asp_count(ArrayElement(ArrayElement(arrTabs,j),"arrFields")))
								setArrElementN tabsAndFields,CreateArray2(ArrayElement(ArrayElement(arrTabs,j),"tabId"),empty),ArrayElement(ArrayElement(ArrayElement(arrTabs,j),"arrFields"),f)
								f = CSmartDbl(f)+1
							loop
						end if
						j = CSmartDbl(j)+1
					loop
					setArrElementN this_object.controlsMap,CreateArray2("tabs","tabGroup_" & CSmartStr(tabGroupName)),tabsAndFields
					beginGroup = false
				end if
			end if
		else
			setArrElementN this_object.controlsMap,CreateArray2("sections","section_" & CSmartStr(ArrayElement(tabC,"tabId"))),CreateDictionary()
			f = 0
			do while IsLess(f,asp_count(ArrayElement(ArrayElement(arrTabs,i),"arrFields")))
				setArrElementN this_object.controlsMap,CreateArray3("sections","section_" & CSmartStr(ArrayElement(tabC,"tabId")),empty),ArrayElement(ArrayElement(ArrayElement(arrTabs,i),"arrFields"),f)
				f = CSmartDbl(f)+1
			loop
		end if
		i = CSmartDbl(i)+1
	loop
End Function
Function method_RunnerPage_isAppearOnTabs(byref this_object,ByVal fName)
	Dim match,arrTabs,this,val
	match = false
	doAssignmentByRef arrTabs,this_object.getArrTabs()
	if not bValue(arrTabs) then
		doAssignmentByRef method_RunnerPage_isAppearOnTabs,match
		Exit Function
	end if
	GetCollectionBounds arrTabs,c_runnerpage_loopIdx22,c_runnerpage_loopMax22
	do while c_runnerpage_loopIdx22<=c_runnerpage_loopMax22
		tab = GetCollectionKey(arrTabs,c_runnerpage_loopIdx22)
		doAssignment val,ArrayElement(arrTabs,tab)
		if bValue(asp_in_array(fName,ArrayElement(val,"arrFields"),false)) then
			match = true
			exit do
		end if
		c_runnerpage_loopIdx22=c_runnerpage_loopIdx22+1
	loop
	doAssignmentByRef method_RunnerPage_isAppearOnTabs,match
	Exit Function
End Function
Function method_RunnerPage_getArrTabs(byref this_object)
	if IsEqual(this_object.pageType,PAGE_EDIT) then
		doAssignmentByRef method_RunnerPage_getArrTabs,this_object.arrEditTabs
		Exit Function
	else
		if IsEqual(this_object.pageType,PAGE_ADD) then
			doAssignmentByRef method_RunnerPage_getArrTabs,this_object.arrAddTabs
			Exit Function
		else
			if IsEqual(this_object.pageType,PAGE_VIEW) then
				doAssignmentByRef method_RunnerPage_getArrTabs,this_object.arrViewTabs
				Exit Function
			else
				method_RunnerPage_getArrTabs = false
				Exit Function
			end if
		end if
	end if
End Function
Function method_RunnerPage_fillTimePickSettings(byref this_object,ByVal field,ByVal value)
	Dim timeAttrs,convention,locAmPm,tpVal,range,h,minutes,m,timePickSet,this
	doAssignmentByRef timeAttrs,GetFieldData(this_object.tName,field,"FormatTimeAttrs",CreateDictionary())
	if bValue(asp_count(timeAttrs)) and bValue(ArrayElement(timeAttrs,"useTimePicker")) then
		doAssignment convention,ArrayElement(timeAttrs,"hours")
		doAssignmentByRef locAmPm,getLacaleAmPmForTimePicker(convention,true)
		doAssignmentByRef tpVal,getValForTimePicker(GetFieldType(field,""),value,ArrayElement(locAmPm,"locale"))
		Set range = (CreateDictionary())
		if IsEqual(convention,24) then
			h = 0
			do while IsLess(h,convention)
				setArrElement range,asp_count(range),h
				h = CSmartDbl(h)+1
			loop
		else
			h = 1
			do while IsLessOrEqual(h,convention)
				setArrElement range,asp_count(range),h
				h = CSmartDbl(h)+1
			loop
		end if
		Set minutes = (CreateDictionary())
		m = 0
		do while IsLess(m,60)
			setArrElement minutes,asp_count(minutes),m
			m = CSmartDbl(m)+CSmartDbl(ArrayElement(timeAttrs,"minutes"))
		loop
		Set timePickSet = (CreateDictionary6("convention",convention,"range",range,"apm",CreateDictionary2(Empty,ArrayElement(locAmPm,"am"),Empty,ArrayElement(locAmPm,"pm")),"rangeMin",minutes,"locale",ArrayElement(locAmPm,"locale"),"showSec",ArrayElement(timeAttrs,"showSeconds")))
		if IsLess(0,asp_count(ArrayElement(tpVal,"dbtime"))) then
			setArrElement timePickSet,"hover",CreateDictionary3("0",ArrayElement(ArrayElement(tpVal,"dbtime"),3),"1",ArrayElement(ArrayElement(tpVal,"dbtime"),4),"2",ArrayElement(ArrayElement(tpVal,"dbtime"),5))
		end if
		if not bValue(asp_array_key_exists(field,ArrayElement(ArrayElement(ArrayElement(this_object.jsSettings,"tableSettings"),this_object.tName),"fieldSettings"))) then
			setArrElementN this_object.jsSettings,CreateArray4("tableSettings",this_object.tName,"fieldSettings",field),CreateDictionary()
			setArrElementN this_object.jsSettings,CreateArray5("tableSettings",this_object.tName,"fieldSettings",field,"timePick"),timePickSet
		else
			if not bValue(asp_array_key_exists("timePick",ArrayElement(ArrayElement(ArrayElement(ArrayElement(this_object.jsSettings,"tableSettings"),this_object.tName),"fieldSettings"),field))) then
				setArrElementN this_object.jsSettings,CreateArray5("tableSettings",this_object.tName,"fieldSettings",field,"timePick"),timePickSet
			end if
		end if
		if not bValue(asp_array_key_exists("timePick",ArrayElement(this_object.jsSettings,"fieldDefSettings"))) then
			setArrElementN this_object.jsSettings,CreateArray2("fieldDefSettings","timePick"),CreateDictionary()
		end if
		this_object.fillControlsMap_p3 CreateDictionary1("controls",CreateDictionary1("open",IIF(ArrayElement(tpVal,"val"),true,false))),true,field
	end if
End Function
Function method_RunnerPage_assignBodyEnd(byref this_object,ByRef params)
	Dim this
	this_object.fillSetCntrlMaps 
	ResponseWrite ((((("<script>" & vbcrlf & _
		"			window.controlsMap = '" & CSmartStr(jsreplace(my_json_encode(this_object.controlsHTMLMap)))) & "'; ") & CSmartStr(this_object.PrepareJS())) & vbcrlf & _
		"			window.settings = '") & CSmartStr(jsreplace(my_json_encode(this_object.jsSettings)))) & "'; 					" & vbcrlf & _
		"		</script>"
End Function
Function method_RunnerPage_genId(byref this_object)
	this_object.flyId = CSmartDbl(this_object.flyId)+1
	doClassAssignment this_object,"recId",this_object.flyId
	doAssignmentByRef method_RunnerPage_genId,this_object.flyId
	Exit Function
End Function
Function method_RunnerPage_getPageType(byref this_object)
	doAssignmentByRef method_RunnerPage_getPageType,this_object.pageType
	Exit Function
End Function
Function method_RunnerPage_AddJSCode(byref this_object,ByVal jscode)
	this_object.totalCode = CSmartStr(this_object.totalCode) & CSmartStr(jscode)
End Function
Function method_RunnerPage_AddJSFileNoExt(byref this_object,ByVal file)
	setArrElement this_object.includes_js,asp_count(this_object.includes_js),file
End Function
Function method_RunnerPage_AddJSFile(byref this_object,ByVal file,ByVal req1,ByVal req2,ByVal req3)
	file = CSmartStr(file) & ".js"
	setArrElement this_object.includes_js,asp_count(this_object.includes_js),file
	if not IsEqual(req1,"") then
		setArrElement this_object.includes_jsreq,file,CreateDictionary1(Empty,CSmartStr(req1) & ".js")
	end if
	if not IsEqual(req2,"") then
		setArrElementN this_object.includes_jsreq,CreateArray2(file,empty),CSmartStr(req2) & ".js"
	end if
	if not IsEqual(req3,"") then
		setArrElementN this_object.includes_jsreq,CreateArray2(file,empty),CSmartStr(req3) & ".js"
	end if
End Function
Function method_RunnerPage_grabAllJsFiles(byref this_object)
	Dim jsFiles,file
	Set jsFiles = (CreateDictionary())
	GetCollectionBounds this_object.includes_js,c_runnerpage_loopIdx26,c_runnerpage_loopMax26
	do while c_runnerpage_loopIdx26<=c_runnerpage_loopMax26
		c_runnerpage_arrKey26 = GetCollectionKey(this_object.includes_js,c_runnerpage_loopIdx26)
		doAssignment file,ArrayElement(this_object.includes_js,c_runnerpage_arrKey26)
		setArrElement jsFiles,file,CreateDictionary()
		if bValue(asp_array_key_exists(file,this_object.includes_jsreq)) then
			setArrElement jsFiles,file,ArrayElement(this_object.includes_jsreq,file)
		end if
		c_runnerpage_loopIdx26=c_runnerpage_loopIdx26+1
	loop
	doClassAssignment this_object,"includes_js",CreateDictionary()
	doClassAssignment this_object,"includes_jsreq",CreateDictionary()
	doAssignmentByRef method_RunnerPage_grabAllJsFiles,jsFiles
	Exit Function
End Function
Function method_RunnerPage_copyAllJsFiles(byref this_object,ByVal jsFiles)
	Dim file,reqFiles,rFile
	GetCollectionBounds jsFiles,c_runnerpage_loopIdx27,c_runnerpage_loopMax27
	do while c_runnerpage_loopIdx27<=c_runnerpage_loopMax27
		file = GetCollectionKey(jsFiles,c_runnerpage_loopIdx27)
		doAssignment reqFiles,ArrayElement(jsFiles,file)
		setArrElement this_object.includes_js,asp_count(this_object.includes_js),file
		if bValue(asp_array_key_exists(file,this_object.includes_jsreq)) then
			GetCollectionBounds reqFiles,c_runnerpage_loopIdx28,c_runnerpage_loopMax28
			c_runnerpage_exitLoop28=false
			do while c_runnerpage_loopIdx28<=c_runnerpage_loopMax28
				c_runnerpage_exitLoop28=false
				do
					c_runnerpage_arrKey28 = GetCollectionKey(reqFiles,c_runnerpage_loopIdx28)
					doAssignment rFile,ArrayElement(reqFiles,c_runnerpage_arrKey28)
					if bValue(asp_array_key_exists(rFile,ArrayElement(this_object.includes_jsreq,file))) then
						exit do
					end if
					setArrElementN this_object.includes_jsreq,CreateArray2(file,empty),rFile
				loop while false
				if c_runnerpage_exitLoop28 then _
					exit do
				c_runnerpage_loopIdx28=c_runnerpage_loopIdx28+1
			loop
		else
			setArrElement this_object.includes_jsreq,file,reqFiles
		end if
		c_runnerpage_loopIdx27=c_runnerpage_loopIdx27+1
	loop
End Function
Function method_RunnerPage_AddCSSFile(byref this_object,ByVal file)
	setArrElement this_object.includes_css,asp_count(this_object.includes_css),file
End Function
Function method_RunnerPage_grabAllCssFiles(byref this_object)
	Dim cssFiles
	doAssignment cssFiles,this_object.includes_css
	doClassAssignment this_object,"includes_css",CreateDictionary()
	doAssignmentByRef method_RunnerPage_grabAllCssFiles,cssFiles
	Exit Function
End Function
Function method_RunnerPage_copyAllCssFiles(byref this_object,ByVal cssFiles)
	Dim this
	GetCollectionBounds cssFiles,c_runnerpage_loopIdx29,c_runnerpage_loopMax29
	do while c_runnerpage_loopIdx29<=c_runnerpage_loopMax29
		c_runnerpage_arrKey29 = GetCollectionKey(cssFiles,c_runnerpage_loopIdx29)
		doAssignment file,ArrayElement(cssFiles,c_runnerpage_arrKey29)
		this_object.AddCSSFile_p1 file
		c_runnerpage_loopIdx29=c_runnerpage_loopIdx29+1
	loop
End Function
Function method_RunnerPage_LoadJS_CSS(byref this_object)
	Dim out,file,req,this
	doClassAssignment this_object,"includes_js",asp_array_unique(this_object.includes_js)
	doClassAssignment this_object,"includes_css",asp_array_unique(this_object.includes_css)
	out = ""
	if bValue(asp_count(this_object.includes_css)) then
		out = CSmartStr(out) & vbcrlf & _
			"Runner.util.ScriptLoader.loadCSS(["
		GetCollectionBounds this_object.includes_css,c_runnerpage_loopIdx30,c_runnerpage_loopMax30
		do while c_runnerpage_loopIdx30<=c_runnerpage_loopMax30
			c_runnerpage_arrKey30 = GetCollectionKey(this_object.includes_css,c_runnerpage_loopIdx30)
			doAssignment file,ArrayElement(this_object.includes_css,c_runnerpage_arrKey30)
			out = CSmartStr(out) & (("'" & CSmartStr(file)) & "', ")
			c_runnerpage_loopIdx30=c_runnerpage_loopIdx30+1
		loop
		doAssignmentByRef out,asp_substr(out,0,CSmartDbl(asp_strlen(out))-2)
		out = CSmartStr(out) & "]);" & vbcrlf
	end if
	GetCollectionBounds this_object.includes_js,c_runnerpage_loopIdx31,c_runnerpage_loopMax31
	do while c_runnerpage_loopIdx31<=c_runnerpage_loopMax31
		c_runnerpage_arrKey31 = GetCollectionKey(this_object.includes_js,c_runnerpage_loopIdx31)
		doAssignment file,ArrayElement(this_object.includes_js,c_runnerpage_arrKey31)
		out = CSmartStr(out) & (("Runner.util.ScriptLoader.addJS(['" & CSmartStr(file)) & "']")
		if bValue(asp_array_key_exists(file,this_object.includes_jsreq)) then
			GetCollectionBounds ArrayElement(this_object.includes_jsreq,file),c_runnerpage_loopIdx32,c_runnerpage_loopMax32
			do while c_runnerpage_loopIdx32<=c_runnerpage_loopMax32
				c_runnerpage_arrKey32 = GetCollectionKey(ArrayElement(this_object.includes_jsreq,file),c_runnerpage_loopIdx32)
				doAssignment req,ArrayElement(ArrayElement(this_object.includes_jsreq,file),c_runnerpage_arrKey32)
				out = CSmartStr(out) & ((",'" & CSmartStr(req)) & "'")
				c_runnerpage_loopIdx32=c_runnerpage_loopIdx32+1
			loop
		end if
		out = CSmartStr(out) & ");" & vbcrlf
		c_runnerpage_loopIdx31=c_runnerpage_loopIdx31+1
	loop
	out = CSmartStr(out) & (CSmartStr(this_object.initForDetailsPreview()) & " Runner.util.ScriptLoader.load();")
	doAssignmentByRef method_RunnerPage_LoadJS_CSS,out
	Exit Function
End Function
Function method_RunnerPage_initForDetailsPreview(byref this_object)
	method_RunnerPage_initForDetailsPreview = ""
	Exit Function
End Function
Function method_RunnerPage_setLangParams(byref this_object)
End Function
Function method_RunnerPage_addCommonJs(byref this_object)
	Dim this
	this_object.AddJSFile_p1 "include/json"
	this_object.AddJSFile_p1 "include/cookies"
	this_object.AddJSFile_p1 "include/lang/" & CSmartStr(getLangFileName(mlang_getcurrentlang()))
	if bValue(ArrayElement(this_object.googleMapCfg,"isUseGoogleMap")) and not bValue(postvalue("onFly")) then
		setArrElement this_object.body,"begin",CSmartStr(ArrayElement(this_object.body,"begin")) & (("<script src=""http://maps.google.com/maps?file=api&amp;v=2&amp;sensor=true&amp;key=" & CSmartStr(ArrayElement(this_object.googleMapCfg,"APIcode"))) & """ type=""text/javascript""></script>")
	end if
	if IsIdentical(this_object.debugJSMode,true) then
		this_object.AddJSFile_p2 "include/runnerJS/ControlConstants","include/lang/" & CSmartStr(getLangFileName(mlang_getcurrentlang()))
		this_object.AddJSFile_p2 "include/runnerJS/RunnerEvent","include/runnerJS/IEHelper"
		this_object.AddJSFile_p2 "include/runnerJS/Validate","include/runnerJS/RunnerEvent"
		this_object.AddJSFile_p2 "include/runnerJS/ControlManager","include/runnerJS/Validate"
		this_object.AddJSFile_p2 "include/runnerJS/button","include/runnerJS/ControlManager"
		this_object.AddJSFile_p2 "include/runnerJS/Control","include/runnerJS/ControlManager"
		this_object.AddJSFile_p2 "include/runnerJS/ReadOnly","include/runnerJS/Control"
		this_object.AddJSFile_p2 "include/runnerJS/TextAreaControl","include/runnerJS/Control"
		this_object.AddJSFile_p2 "include/runnerJS/TextFieldControl","include/runnerJS/Control"
		this_object.AddJSFile_p2 "include/runnerJS/TimeFieldControl","include/runnerJS/Control"
		this_object.AddJSFile_p2 "include/runnerJS/RteControl","include/runnerJS/Control"
		this_object.AddJSFile_p2 "include/runnerJS/FileControl","include/runnerJS/Control"
		this_object.AddJSFile_p2 "include/runnerJS/DateFieldControl","include/runnerJS/Control"
		this_object.AddJSFile_p2 "include/runnerJS/LookupWizard","include/runnerJS/Control"
		this_object.AddJSFile_p2 "include/runnerJS/RadioControl","include/runnerJS/LookupWizard"
		this_object.AddJSFile_p2 "include/runnerJS/DropDown","include/runnerJS/LookupWizard"
		this_object.AddJSFile_p2 "include/runnerJS/CheckBox","include/runnerJS/LookupWizard"
		this_object.AddJSFile_p2 "include/runnerJS/TextFieldLookup","include/runnerJS/LookupWizard"
		this_object.AddJSFile_p2 "include/runnerJS/EditBoxLookup","include/runnerJS/TextFieldLookup"
		this_object.AddJSFile_p2 "include/runnerJS/ListPageLookup","include/runnerJS/TextFieldLookup"
		this_object.AddJSFile_p1 "include/runnerJS/InlineEdit"
		this_object.AddJSFile_p2 "include/runnerJS/pages/PageConstants","include/runnerJS/ListPageLookup"
		this_object.AddJSFile_p2 "include/runnerJS/pages/PageManager","include/runnerJS/pages/PageConstants"
		this_object.AddJSFile_p2 "include/runnerJS/pages/PageSettings","include/runnerJS/pages/PageManager"
		this_object.AddJSFile_p2 "include/runnerJS/pages/RunnerPage","include/runnerJS/pages/PageSettings"
		this_object.AddJSFile_p2 "include/runnerJS/pages/SearchPage","include/runnerJS/pages/RunnerPage"
		this_object.AddJSFile_p2 "include/runnerJS/pages/ViewPage","include/runnerJS/pages/RunnerPage"
		this_object.AddJSFile_p2 "include/runnerJS/pages/EditorPage","include/runnerJS/pages/RunnerPage"
		this_object.AddJSFile_p2 "include/runnerJS/pages/AddPageFly","include/runnerJS/pages/EditorPage"
		this_object.AddJSFile_p2 "include/runnerJS/pages/AddPage","include/runnerJS/pages/AddPageFly"
		this_object.AddJSFile_p2 "include/runnerJS/pages/EditPage","include/runnerJS/pages/EditorPage"
		this_object.AddJSFile_p2 "include/runnerJS/pages/DataPageWithSearch","include/runnerJS/pages/RunnerPage"
		this_object.AddJSFile_p2 "include/runnerJS/pages/ListPageFly","include/runnerJS/pages/DataPageWithSearch"
		this_object.AddJSFile_p2 "include/runnerJS/pages/ListPage","include/runnerJS/pages/DataPageWithSearch"
		this_object.AddJSFile_p2 "include/runnerJS/pages/ListPageAjax","include/runnerJS/pages/ListPage"
		this_object.AddJSFile_p2 "include/runnerJS/pages/ReportPage","include/runnerJS/pages/DataPageWithSearch"
		this_object.AddJSFile_p2 "include/runnerJS/pages/ChartPage","include/runnerJS/pages/DataPageWithSearch"
		this_object.AddJSFile_p2 "include/runnerJS/pages/MembersPage","include/runnerJS/pages/ListPage"
		this_object.AddJSFile_p2 "include/runnerJS/pages/RightsPage","include/runnerJS/pages/ListPage"
		this_object.AddJSFile_p2 "include/runnerJS/pages/ExportPage","include/runnerJS/pages/RunnerPage"
		this_object.AddJSFile_p2 "include/runnerJS/pages/ImportPage","include/runnerJS/pages/RunnerPage"
		this_object.AddJSFile_p2 "include/runnerJS/pages/RegisterPage","include/runnerJS/pages/RunnerPage"
		this_object.AddJSFile_p1 "include/runnerJS/SearchForm"
		this_object.AddJSFile_p2 "include/runnerJS/SearchFormWithUI","include/runnerJS/SearchForm"
		this_object.AddJSFile_p2 "include/runnerJS/SearchController","include/runnerJS/SearchFormWithUI"
		this_object.AddJSFile_p1 "include/runnerJS/RunnerForm"
		this_object.AddJSFile_p1 "include/yui/yahoo"
		this_object.AddJSFile_p2 "include/yui/cookie","include/yui/yahoo"
		this_object.AddJSFile_p2 "include/yui/dom","include/yui/cookie"
		this_object.AddJSFile_p2 "include/yui/event","include/yui/dom"
		this_object.AddJSFile_p2 "include/yui/element","include/yui/event"
		this_object.AddJSFile_p2 "include/yui/tabview","include/yui/element"
		this_object.AddJSFile_p2 "include/yui/datasource","include/yui/element"
		this_object.AddJSFile_p2 "include/yui/dragdrop","include/yui/datasource"
		this_object.AddJSFile_p2 "include/yui/animation","include/yui/dragdrop"
		this_object.AddJSFile_p2 "include/yui/connection","include/yui/animation"
		this_object.AddJSFile_p2 "include/yui/container","include/yui/connection"
		this_object.AddJSFile_p2 "include/yui/datatable","include/yui/dragdrop"
		this_object.AddJSFile_p2 "include/yui/json","include/yui/datatable"
		this_object.AddJSFile_p2 "include/yui/resize","include/yui/json"
		this_object.AddCSSFile_p1 "include/yui/container"
		this_object.AddCSSFile_p1 "include/yui/container-skin"
		this_object.AddCSSFile_p1 "include/yui/resize"
		this_object.AddCSSFile_p1 "include/yui/resize-skin"
		this_object.AddCSSFile_p1 "include/yui/tabview"
		this_object.AddCSSFile_p1 "include/yui/tabview-skin"
		if bValue(this_object.calendar) then
			this_object.AddJSFile_p3 "include/yui/calendar","include/yui/event","include/yui/container"
			this_object.AddCSSFile_p1 "include/yui/calendar"
			this_object.AddCSSFile_p1 "include/yui/calendar-skin"
		end if
		this_object.AddCSSFile_p1 "include/redefineStyle"
		if bValue(GetTableData(this_object.tName,".addPageEvents",false)) and not IsEqual(this_object.mode,LIST_DETAILS) then
			this_object.AddJSFile_p3 "include/runnerJS/events/pageevents_" & CSmartStr(GetTableData(this_object.tName,".shortTableName","")),"include/runnerJS/pages/PageSettings","include/runnerJS/button"
		end if
	else
		this_object.AddJSFile_p2 "include/runnerJS/RunnerControls","include/lang/" & CSmartStr(getLangFileName(mlang_getcurrentlang()))
		this_object.AddJSFile_p2 "include/runnerJS/RunnerPages","include/runnerJS/RunnerControls"
		this_object.AddJSFile_p1 "include/yui/yuiAll"
		this_object.AddCSSFile_p1 "include/yui/container"
		this_object.AddCSSFile_p1 "include/yui/container-skin"
		this_object.AddCSSFile_p1 "include/yui/resize"
		this_object.AddCSSFile_p1 "include/yui/resize-skin"
		if bValue(GetTableData(this_object.tName,".addPageEvents",false)) and not IsEqual(this_object.mode,LIST_DETAILS) then
			this_object.AddJSFile_p3 "include/runnerJS/events/pageevents_" & CSmartStr(GetTableData(this_object.tName,".shortTableName","")),"include/runnerJS/RunnerControls","include/runnerJS/RunnerPages"
		end if
		if bValue(this_object.calendar) then
			this_object.AddCSSFile_p1 "include/yui/calendar"
			this_object.AddCSSFile_p1 "include/yui/calendar-skin"
		end if
		this_object.AddCSSFile_p1 "include/yui/tabview"
		this_object.AddCSSFile_p1 "include/yui/tabview-skin"
		this_object.AddCSSFile_p1 "include/redefineStyle"
	end if
	this_object.AddJSFile_p1 "include/customlabels"
	if bValue(this_object.isUseAjaxSuggest) then
		this_object.AddJSFile_p1 "include/ajaxsuggest"
	end if
	if bValue(this_object.timepicker) then
		this_object.AddJSFile_p1 "include/ui"
		this_object.AddJSFile_p2 "include/jquery.utils","include/ui"
		this_object.AddJSFile_p2 "include/ui.dropslide","include/jquery.utils"
		this_object.AddJSFile_p2 "include/ui.timepickr","include/ui.dropslide"
		this_object.AddCSSFile_p1 "include/ui.dropslide"
	end if
	if bValue(this_object.isUseToolTips) then
		this_object.AddJSFile_p1 "include/jquery.dimensions"
		this_object.AddJSFile_p2 "include/jquery.inputhintbox","include/jquery.dimensions"
		this_object.AddCSSFile_p1 "include/tooltip"
	end if
	if bValue(this_object.useIbox) then
		this_object.AddJSFile_p1 "include/ibox"
		this_object.AddCSSFile_p1 "include/ibox"
	end if
	if (not bValue(this_object.useDetailsPreview) and ((bValue(this_object.showAddInPopup) or bValue(this_object.showEditInPopup)) or bValue(this_object.showViewInPopup))) and ((bValue(displayDetailsOn(this_object.tName,PAGE_EDIT)) or bValue(displayDetailsOn(this_object.tName,PAGE_ADD))) or bValue(displayDetailsOn(this_object.tName,PAGE_VIEW))) then
		this_object.AddJSFile_p1 "include/detailspreview"
	end if
	if bValue(this_object.isUseCK) then
		this_object.addJSFile_p1 "plugins/ckeditor/ckeditor"
	end if
	if bValue(this_object.isUseVideo) then
		this_object.addJSFile_p1 "include/video/flowplayer-3.2.3.min"
	end if
	if bValue(this_object.isUseAudio) then
		this_object.AddJSFileNoExt_p1 "http://mediaplayer.yahoo.com/js"
	end if
	if (IsEqual(this_object.pageType,PAGE_EDIT) and bValue(this_object.useTabsOnEdit) or IsEqual(this_object.pageType,PAGE_ADD) and bValue(this_object.useTabsOnAdd)) or IsEqual(this_object.pageType,PAGE_VIEW) and bValue(this_object.useTabsOnView) then
		this_object.AddCSSFile_p1 "include/redefineStyle"
	end if
End Function
Function method_RunnerPage_PrepareJs(byref this_object)
	Dim js,this
	doAssignmentByRef method_RunnerPage_PrepareJs,doAssignmentByRef(js,this_object.LoadJS_CSS())
	Exit Function
End Function
Function method_RunnerPage_addButtonHandlers(byref this_object)
	Dim this
	if not bValue(GetTableData(this_object.tName,".isUsebuttonHandlers",false)) then
		method_RunnerPage_addButtonHandlers = false
		Exit Function
	end if
	if IsIdentical(this_object.debugJSMode,true) then
		this_object.AddJSFile_p2 "include/runnerJS/events/pageevents_" & CSmartStr(GetTableData(this_object.tName,".shortTableName","")),"include/runnerJS/pages/PageSettings"
	else
		this_object.AddJSFile_p3 "include/runnerJS/events/pageevents_" & CSmartStr(GetTableData(this_object.tName,".shortTableName","")),"include/runnerJS/RunnerControls","include/runnerJS/RunnerPages"
	end if
	this_object.AddJSFile_p1 "include/json"
	method_RunnerPage_addButtonHandlers = true
	Exit Function
End Function
Function method_RunnerPage_addOnLoadJsEvent(byref this_object,ByVal code)
	this_object.onLoadJsEventCode = CSmartStr(this_object.onLoadJsEventCode) & (";" & vbcrlf & CSmartStr(code))
End Function
Function method_RunnerPage_setGoogleMapsParams(byref this_object,ByVal fieldsArr)
	Dim f,fieldMap
	setArrElement this_object.googleMapCfg,"isUseMainMaps",IsLess(0,asp_count(ArrayElement(this_object.googleMapCfg,"mainMapIds")))
	GetCollectionBounds fieldsArr,c_runnerpage_loopIdx33,c_runnerpage_loopMax33
	do while c_runnerpage_loopIdx33<=c_runnerpage_loopMax33
		c_runnerpage_arrKey33 = GetCollectionKey(fieldsArr,c_runnerpage_loopIdx33)
		doAssignment f,ArrayElement(fieldsArr,c_runnerpage_arrKey33)
		if IsEqual(ArrayElement(f,"viewFormat"),FORMAT_MAP) then
			setArrElement this_object.googleMapCfg,"isUseFieldsMaps",true
			setArrElementN this_object.googleMapCfg,CreateArray2("fieldsAsMap",ArrayElement(f,"fName")),CreateDictionary()
			doAssignmentByRef fieldMap,GetFieldData(this_object.tName,ArrayElement(f,"fName"),"mapData",CreateDictionary())
			setArrElementN this_object.googleMapCfg,CreateArray3("fieldsAsMap",ArrayElement(f,"fName"),"width"),IIF(ArrayElement(fieldMap,"width"),ArrayElement(fieldMap,"width"),0)
			setArrElementN this_object.googleMapCfg,CreateArray3("fieldsAsMap",ArrayElement(f,"fName"),"height"),IIF(ArrayElement(fieldMap,"height"),ArrayElement(fieldMap,"height"),0)
			setArrElementN this_object.googleMapCfg,CreateArray3("fieldsAsMap",ArrayElement(f,"fName"),"addressField"),ArrayElement(fieldMap,"address")
			setArrElementN this_object.googleMapCfg,CreateArray3("fieldsAsMap",ArrayElement(f,"fName"),"latField"),ArrayElement(fieldMap,"lat")
			setArrElementN this_object.googleMapCfg,CreateArray3("fieldsAsMap",ArrayElement(f,"fName"),"lngField"),ArrayElement(fieldMap,"lng")
			if not IsEmpty(ArrayElement(fieldMap,"zoom")) then
				setArrElementN this_object.googleMapCfg,CreateArray3("fieldsAsMap",ArrayElement(f,"fName"),"zoom"),ArrayElement(fieldMap,"zoom")
			end if
		end if
		c_runnerpage_loopIdx33=c_runnerpage_loopIdx33+1
	loop
	setArrElement this_object.googleMapCfg,"isUseGoogleMap",bValue(ArrayElement(this_object.googleMapCfg,"isUseMainMaps")) or bValue(ArrayElement(this_object.googleMapCfg,"isUseFieldsMaps"))
End Function
Function method_RunnerPage_addBigGoogleMapMarkers(byref this_object,ByRef data,ByVal viewLink)
	Dim latF,mapId,lngF,addressF,descF,markerArr
	GetCollectionBounds ArrayElement(this_object.googleMapCfg,"mainMapIds"),c_runnerpage_loopIdx34,c_runnerpage_loopMax34
	do while c_runnerpage_loopIdx34<=c_runnerpage_loopMax34
		c_runnerpage_arrKey34 = GetCollectionKey(ArrayElement(this_object.googleMapCfg,"mainMapIds"),c_runnerpage_loopIdx34)
		doAssignment mapId,ArrayElement(ArrayElement(this_object.googleMapCfg,"mainMapIds"),c_runnerpage_arrKey34)
		doAssignment latF,ArrayElement(ArrayElement(ArrayElement(this_object.googleMapCfg,"mapsData"),mapId),"latField")
		doAssignment lngF,ArrayElement(ArrayElement(ArrayElement(this_object.googleMapCfg,"mapsData"),mapId),"lngField")
		doAssignment addressF,ArrayElement(ArrayElement(ArrayElement(this_object.googleMapCfg,"mapsData"),mapId),"addressField")
		doAssignment descF,ArrayElement(ArrayElement(ArrayElement(this_object.googleMapCfg,"mapsData"),mapId),"descField")
		Set markerArr = (CreateDictionary())
		setArrElement markerArr,"lat",IIF(ArrayElement(data,latF),ArrayElement(data,latF),"")
		setArrElement markerArr,"lng",IIF(ArrayElement(data,lngF),ArrayElement(data,lngF),"")
		setArrElement markerArr,"address",IIF(ArrayElement(data,addressF),ArrayElement(data,addressF),"")
		setArrElement markerArr,"desc",IIF(ArrayElement(data,descF),ArrayElement(data,descF),ArrayElement(markerArr,"address"))
		setArrElement markerArr,"link",viewLink
		setArrElement markerArr,"recId",this_object.recId
		setArrElementN this_object.googleMapCfg,CreateArray4("mapsData",mapId,"markers",empty),markerArr
		c_runnerpage_loopIdx34=c_runnerpage_loopIdx34+1
	loop
End Function
Function method_RunnerPage_addGoogleMapData(byref this_object,ByVal fName,ByRef data,ByVal viewLink)
	Dim fieldMap,mapData,address,lat,lng,desc,zoom
	doAssignmentByRef fieldMap,GetFieldData(this_object.tName,fName,"mapData",CreateDictionary())
	setArrElement mapData,"mapFieldValue",ArrayElement(data,fName)
	doAssignment address,IIF(ArrayElement(data,ArrayElement(fieldMap,"address")),ArrayElement(data,ArrayElement(fieldMap,"address")),"")
	doAssignment lat,IIF(ArrayElement(data,ArrayElement(fieldMap,"lat")),ArrayElement(data,ArrayElement(fieldMap,"lat")),"")
	doAssignment lng,IIF(ArrayElement(data,ArrayElement(fieldMap,"lng")),ArrayElement(data,ArrayElement(fieldMap,"lng")),"")
	doAssignment desc,IIF(ArrayElement(data,ArrayElement(fieldMap,"desc")),ArrayElement(data,ArrayElement(fieldMap,"desc")),address)
	if not IsEmpty(ArrayElement(fieldMap,"zoom")) then
		doAssignment zoom,ArrayElement(fieldMap,"zoom")
	else
		zoom = ""
	end if
	setArrElement mapData,"fName",fName
	setArrElement mapData,"zoom",zoom
	setArrElement mapData,"type","FIELD_MAP"
	setArrElementN mapData,CreateArray2("markers",empty),CreateDictionary6("address",address,"lat",lat,"lng",lng,"link",viewLink,"desc",desc,"recId",this_object.recId)
	setArrElementN this_object.googleMapCfg,CreateArray2("mapsData",(("littleMap_" & CSmartStr(GoodFieldName(fName))) & "_") & CSmartStr(this_object.recId)),mapData
	setArrElementN this_object.googleMapCfg,CreateArray2("fieldMapsIds",empty),(("littleMap_" & CSmartStr(GoodFieldName(fName))) & "_") & CSmartStr(this_object.recId)
	doAssignmentByRef method_RunnerPage_addGoogleMapData,ArrayElement(ArrayElement(this_object.googleMapCfg,"mapsData"),(("littleMap_" & CSmartStr(GoodFieldName(fName))) & "_") & CSmartStr(this_object.recId))
	Exit Function
End Function
Function method_RunnerPage_initGmaps(byref this_object)
	Dim mapId,this
	if bValue(ArrayElement(this_object.googleMapCfg,"isUseGoogleMap")) then
		GetCollectionBounds ArrayElement(this_object.googleMapCfg,"mainMapIds"),c_runnerpage_loopIdx35,c_runnerpage_loopMax35
		do while c_runnerpage_loopIdx35<=c_runnerpage_loopMax35
			c_runnerpage_arrKey35 = GetCollectionKey(ArrayElement(this_object.googleMapCfg,"mainMapIds"),c_runnerpage_loopIdx35)
			doAssignment mapId,ArrayElement(ArrayElement(this_object.googleMapCfg,"mainMapIds"),c_runnerpage_arrKey35)
			if IsIdentical(ArrayElement(ArrayElement(ArrayElement(this_object.googleMapCfg,"mapsData"),mapId),"showCenterLink"),1) then
				setArrElement this_object.googleMapCfg,"centerLinkText",ArrayElement(ArrayElement(ArrayElement(this_object.googleMapCfg,"mapsData"),mapId),"centerLinkText")
				exit do
			end if
			c_runnerpage_loopIdx35=c_runnerpage_loopIdx35+1
		loop
		setArrElement this_object.googleMapCfg,"id",this_object.id
		this_object.AddJSFile_p1 "include/json"
		this_object.AddJSFile_p1 "include/runnerJS/gmap"
		if not bValue(ArrayElement(this_object.googleMapCfg,"APIcode")) then
			setArrElement this_object.googleMapCfg,"APIcode",""
		end if
		setArrElementByRef this_object.controlsMap,"gMaps",this_object.googleMapCfg
	end if
End Function
Function method_RunnerPage_addCenterLink(byref this_object,ByRef value,ByVal fName)
	Dim mapId
	if bValue(ArrayElement(this_object.googleMapCfg,"isUseMainMaps")) then
		GetCollectionBounds ArrayElement(this_object.googleMapCfg,"mainMapIds"),c_runnerpage_loopIdx36,c_runnerpage_loopMax36
		c_runnerpage_exitLoop36=false
		do while c_runnerpage_loopIdx36<=c_runnerpage_loopMax36
			c_runnerpage_exitLoop36=false
			do
				c_runnerpage_arrKey36 = GetCollectionKey(ArrayElement(this_object.googleMapCfg,"mainMapIds"),c_runnerpage_loopIdx36)
				doAssignment mapId,ArrayElement(ArrayElement(this_object.googleMapCfg,"mainMapIds"),c_runnerpage_arrKey36)
				if not IsEqual(ArrayElement(ArrayElement(ArrayElement(this_object.googleMapCfg,"mapsData"),mapId),"addressField"),fName) or not bValue(ArrayElement(ArrayElement(ArrayElement(this_object.googleMapCfg,"mapsData"),mapId),"showCenterLink")) then
					exit do
				end if
				if IsIdentical(ArrayElement(ArrayElement(ArrayElement(this_object.googleMapCfg,"mapsData"),mapId),"showCenterLink"),1) then
					doAssignment value,ArrayElement(ArrayElement(ArrayElement(this_object.googleMapCfg,"mapsData"),mapId),"centerLinkText")
				end if
				method_RunnerPage_addCenterLink = ((((("<a href=""#"" type=""centerOnMarker" & CSmartStr(this_object.id)) & """ recId=""") & CSmartStr(this_object.recId)) & """>") & CSmartStr(value)) & "</a>"
				Exit Function
			loop while false
			if c_runnerpage_exitLoop36 then _
				exit do
			c_runnerpage_loopIdx36=c_runnerpage_loopIdx36+1
		loop
	end if
	doAssignmentByRef method_RunnerPage_addCenterLink,value
	Exit Function
End Function
Function method_RunnerPage_createOldMenu(byref this_object)
	Dim allowedTablesCount,redirect_menu,i,page,strPerm
	allowedTablesCount = 0
	redirect_menu = ""
	i = 0
	do while IsLess(i,asp_count(this_object.menuTablesArr))
		if bValue(ArrayElement(ArrayElement(this_object.permis,ArrayElement(ArrayElement(this_object.menuTablesArr,i),"dataSourceTName")),"add")) or bValue(ArrayElement(ArrayElement(this_object.permis,ArrayElement(ArrayElement(this_object.menuTablesArr,i),"dataSourceTName")),"search")) then
			this_object.xt.assign_p2 CSmartStr(ArrayElement(ArrayElement(this_object.menuTablesArr,i),"dataSourceTName")) & "_tablelink",true
			page = ""
			if bValue(ArrayElement(ArrayElement(this_object.permis,ArrayElement(ArrayElement(this_object.menuTablesArr,i),"dataSourceTName")),"search")) and (IsEqual(ArrayElement(ArrayElement(this_object.menuTablesArr,i),"nType"),titTABLE) or IsEqual(ArrayElement(ArrayElement(this_object.menuTablesArr,i),"nType"),titVIEW)) then
				page = "list"
				if bValue(ArrayElement(ArrayElement(this_object.permis,ArrayElement(ArrayElement(this_object.menuTablesArr,i),"dataSourceTName")),"add")) then
					doAssignmentByRef strPerm,GetUserPermissions(ArrayElement(ArrayElement(this_object.menuTablesArr,i),"strDataSourceTable"))
				end if
				if (not IsEmpty(strPerm) and not IsFalse(asp_strpos(strPerm,"A",empty))) and IsFalse(asp_strpos(strPerm,"S",empty)) then
					page = "add"
				end if
			else
				if bValue(ArrayElement(ArrayElement(this_object.permis,ArrayElement(ArrayElement(this_object.menuTablesArr,i),"dataSourceTName")),"add")) and (IsEqual(ArrayElement(ArrayElement(this_object.menuTablesArr,i),"nType"),titTABLE) or IsEqual(ArrayElement(ArrayElement(this_object.menuTablesArr,i),"nType"),titVIEW)) then
					page = "add"
				else
					if IsEqual(ArrayElement(ArrayElement(this_object.menuTablesArr,i),"nType"),titREPORT) then
						page = "report"
					else
						if IsEqual(ArrayElement(ArrayElement(this_object.menuTablesArr,i),"nType"),titCHART) then
							page = "chart"
						end if
					end if
				end if
			end if
			this_object.xt.assign_p2 CSmartStr(ArrayElement(this_object.menuTablesArr,i)) & "_tablelink_attrs",((("href=""" & CSmartStr(ArrayElement(this_object.menuTablesArr,i))) & "_") & CSmartStr(page)) & ".asp"""
			this_object.xt.assign_p2 ("" & CSmartStr(ArrayElement(this_object.menuTablesArr,i))) & "_optionattrs",((("value=""" & CSmartStr(ArrayElement(this_object.menuTablesArr,i))) & "_") & CSmartStr(page)) & ".asp"""
			redirect_menu = ((CSmartStr(ArrayElement(ArrayElement(this_object.menuTablesArr,i),"shortTName")) & "_") & CSmartStr(page)) & ".asp"
			allowedTablesCount = CSmartDbl(allowedTablesCount)+1
		end if
		i = CSmartDbl(i)+1
	loop
	doAssignmentByRef method_RunnerPage_createOldMenu,CreateDictionary2("menuTablesCount",allowedTablesCount,"urlForRedirect",redirect_menu)
	Exit Function
End Function
Function method_RunnerPage_getPermissions(byref this_object,ByVal tName)
	Dim resArr,strPerm
	Set resArr = (CreateDictionary())
	if not bValue(this_object.isGroupSecurity) then
		setArrElement resArr,"add",true
		setArrElement resArr,"delete",true
		setArrElement resArr,"edit",true
		setArrElement resArr,"search",true
		setArrElement resArr,"export",true
		setArrElement resArr,"import",true
	else
		if not bValue(tName) then
			doAssignment tName,this_object.tName
		end if
		doAssignmentByRef strPerm,GetUserPermissions(tName)
		setArrElement resArr,"add",not IsFalse(asp_strpos(strPerm,"A",empty))
		setArrElement resArr,"delete",not IsFalse(asp_strpos(strPerm,"D",empty))
		setArrElement resArr,"edit",not IsFalse(asp_strpos(strPerm,"E",empty))
		setArrElement resArr,"search",not IsFalse(asp_strpos(strPerm,"S",empty))
		setArrElement resArr,"export",not IsFalse(asp_strpos(strPerm,"P",empty))
		setArrElement resArr,"import",not IsFalse(asp_strpos(strPerm,"I",empty))
	end if
	doAssignmentByRef method_RunnerPage_getPermissions,resArr
	Exit Function
End Function
Function method_RunnerPage_isCreateMenu(byref this_object)
	Dim menuTable
	GetCollectionBounds this_object.menuTablesArr,c_runnerpage_loopIdx38,c_runnerpage_loopMax38
	do while c_runnerpage_loopIdx38<=c_runnerpage_loopMax38
		c_runnerpage_arrKey38 = GetCollectionKey(this_object.menuTablesArr,c_runnerpage_loopIdx38)
		doAssignment menuTable,ArrayElement(this_object.menuTablesArr,c_runnerpage_arrKey38)
		if bValue(ArrayElement(ArrayElement(this_object.permis,ArrayElement(menuTable,"dataSourceTName")),"add")) or bValue(ArrayElement(ArrayElement(this_object.permis,ArrayElement(menuTable,"dataSourceTName")),"search")) then
			method_RunnerPage_isCreateMenu = true
			Exit Function
		end if
		c_runnerpage_loopIdx38=c_runnerpage_loopIdx38+1
	loop
	method_RunnerPage_isCreateMenu = false
	Exit Function
End Function
Function method_RunnerPage_addRunLoading(byref this_object)
End Function
Function method_RunnerPage_eventExists(byref this_object,ByVal name)
	if not bValue(this_object.eventsObject) then
		method_RunnerPage_eventExists = false
		Exit Function
	end if
	doAssignmentByRef method_RunnerPage_eventExists,this_object.eventsObject.exists_p1(name)
	Exit Function
End Function
Function method_RunnerPage_captchaExists(byref this_object)
	if not bValue(this_object.eventsObject) then
		method_RunnerPage_captchaExists = false
		Exit Function
	end if
	doAssignmentByRef method_RunnerPage_captchaExists,this_object.eventsObject.existsCAPTCHA_p1(this_object.pageType)
	Exit Function
End Function
Function method_RunnerPage_getNextPrevRecordKeys(byref this_object,ByRef data,ByVal securityMode,ByRef var_next,ByRef prev)
	Dim var_SESSION,prevExpr,nextExpr,where_next,where_prev,order_next,order_prev,arrFieldForSort,arrHowFieldSort,query,where,having,orderindexes,strOrderBy,keyIndexes,tKeys,ki,k,i,tail,fieldName,fullName,asc,value,nextop,prevop,oWhere,sql_next,sql_prev,oHaving,sql,res_next,row_next,res_prev,row_prev
	Set var_next = (CreateDictionary())
	Set prev = (CreateDictionary())
	if bValue(Session(CSmartStr(this_object.sessionPrefix) & "_noNextPrev")) then
		Exit Function
	end if
	prevExpr = ""
	nextExpr = ""
	where_next = ""
	where_prev = ""
	order_next = ""
	order_prev = ""
	Set arrFieldForSort = (CreateDictionary())
	Set arrHowFieldSort = (CreateDictionary())
	doAssignmentByRef query,GetQueryObject(this_object.tName)
	doAssignment where,Session(CSmartStr(this_object.sessionPrefix) & "_where")
	if not bValue(asp_strlen(where)) then
		doAssignmentByRef where,SecuritySQL(securityMode,"")
	end if
	doAssignment having,Session(CSmartStr(this_object.sessionPrefix) & "_having")
	doAssignmentByRef orderindexes,GetTableData(this_object.tName,".orderindexes",CreateDictionary())
	doAssignmentByRef strOrderBy,GetTableData(this_object.tName,".strOrderBy","")
	Set keyIndexes = (CreateDictionary())
	doAssignmentByRef tKeys,GetTableKeys(this_object.tName)
	GetCollectionBounds tKeys,c_runnerpage_loopIdx39,c_runnerpage_loopMax39
	do while c_runnerpage_loopIdx39<=c_runnerpage_loopMax39
		c_runnerpage_arrKey39 = GetCollectionKey(tKeys,c_runnerpage_loopIdx39)
		doAssignment k,ArrayElement(tKeys,c_runnerpage_arrKey39)
		doAssignmentByRef ki,GetFieldIndex(k,this_object.tName)
		if bValue(ki) then
			setArrElement keyIndexes,asp_count(keyIndexes),ki
		end if
		c_runnerpage_loopIdx39=c_runnerpage_loopIdx39+1
	loop
	if not bValue(asp_count(keyIndexes)) then
		setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_noNextPrev",1
		Exit Function
	end if
	if bValue(Session(CSmartStr(this_object.sessionPrefix) & "_arrFieldForSort")) and bValue(Session(CSmartStr(this_object.sessionPrefix) & "_arrHowFieldSort")) then
		doAssignment arrFieldForSort,Session(CSmartStr(this_object.sessionPrefix) & "_arrFieldForSort")
		doAssignment arrHowFieldSort,Session(CSmartStr(this_object.sessionPrefix) & "_arrHowFieldSort")
	else
		if bValue(asp_count(orderindexes)) then
			i = 0
			do while IsLess(i,asp_count(orderindexes))
				setArrElement arrFieldForSort,asp_count(arrFieldForSort),ArrayElement(ArrayElement(orderindexes,i),0)
				setArrElement arrHowFieldSort,asp_count(arrHowFieldSort),ArrayElement(ArrayElement(orderindexes,i),1)
				i = CSmartDbl(i)+1
			loop
		else
			if not IsEqual(strOrderBy,"") then
				setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_noNextPrev",1
				Exit Function
			end if
		end if
		i = 0
		do while IsLess(i,asp_count(keyIndexes))
			if IsFalse(asp_array_search(ArrayElement(keyIndexes,i),arrFieldForSort,false)) then
				setArrElement arrFieldForSort,asp_count(arrFieldForSort),ArrayElement(keyIndexes,i)
				setArrElement arrHowFieldSort,asp_count(arrHowFieldSort),"ASC"
			end if
			i = CSmartDbl(i)+1
		loop
		setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_arrFieldForSort",arrFieldForSort
		setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_arrHowFieldSort",arrHowFieldSort
	end if
	if not bValue(asp_count(arrFieldForSort)) then
		setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_noNextPrev",1
		Exit Function
	end if
	i = 0
	c_runnerpage_exitLoop42=false
	do while IsLess(i,asp_count(arrFieldForSort))
		c_runnerpage_exitLoop42=false
		do
			if not bValue(GetFieldByIndex(ArrayElement(arrFieldForSort,i),this_object.tName)) then
				exit do
			end if
			if IsEqual(order_next,"") then
				order_next = " ORDER BY "
				order_prev = " ORDER BY "
			else
				order_next = CSmartStr(order_next) & ","
				order_prev = CSmartStr(order_prev) & ","
			end if
			order_next = CSmartStr(order_next) & ((CSmartStr(ArrayElement(arrFieldForSort,i)) & " ") & CSmartStr(ArrayElement(arrHowFieldSort,i)))
			order_prev = CSmartStr(order_prev) & ((CSmartStr(ArrayElement(arrFieldForSort,i)) & " ") & CSmartStr(IIF(IsEqual(ArrayElement(arrHowFieldSort,i),"DESC"),"ASC","DESC")))
		loop while false
		if c_runnerpage_exitLoop42 then _
			exit do
		i = CSmartDbl(i)+1
	loop
	tail = ""
	i = 0
	c_runnerpage_exitLoop43=false
	do while IsLess(i,asp_count(arrFieldForSort))
		c_runnerpage_exitLoop43=false
		do
			doAssignmentByRef fieldName,GetFieldByIndex(ArrayElement(arrFieldForSort,i),this_object.tName)
			if not bValue(fieldName) then
				exit do
			end if
			if not bValue(query.HasGroupBy()) then
				doAssignmentByRef fullName,GetFullFieldName(fieldName,this_object.tName)
			else
				doAssignment fullName,fieldName
			end if
			asc = IsEqual(ArrayElement(arrHowFieldSort,i),"ASC")
			if not bValue(isnull(ArrayElement(data,fieldName))) then
				doAssignmentByRef value,make_db_value(fieldName,ArrayElement(data,fieldName),"","",this_object.tName)
				doAssignment nextop,IIF(asc,">","<")
				doAssignment prevop,IIF(asc,"<",">")
				nextExpr = (CSmartStr(fullName) & CSmartStr(nextop)) & CSmartStr(value)
				prevExpr = (CSmartStr(fullName) & CSmartStr(prevop)) & CSmartStr(value)
				if IsEqual(nextop,"<") then
					nextExpr = CSmartStr(nextExpr) & ((" or " & CSmartStr(fullName)) & " IS NULL")
				else
					prevExpr = CSmartStr(prevExpr) & ((" or " & CSmartStr(fullName)) & " IS NULL")
				end if
				if IsLess(i,CSmartDbl(asp_count(arrFieldForSort))-1) then
					nextExpr = CSmartStr(nextExpr) & (((" or " & CSmartStr(fullName)) & "=") & CSmartStr(value))
					prevExpr = CSmartStr(prevExpr) & (((" or " & CSmartStr(fullName)) & "=") & CSmartStr(value))
				end if
			else
				if bValue(asc) then
					nextExpr = CSmartStr(fullName) & " IS NOT NULL"
				else
					prevExpr = CSmartStr(fullName) & " IS NOT NULL"
				end if
				if IsLess(i,CSmartDbl(asp_count(arrFieldForSort))-1) then
					if not IsEqual(nextExpr,"") then
						nextExpr = CSmartStr(nextExpr) & " or "
					end if
					nextExpr = CSmartStr(nextExpr) & (CSmartStr(fullName) & " IS NULL")
					if not IsEqual(prevExpr,"") then
						prevExpr = CSmartStr(prevExpr) & " or "
					end if
					prevExpr = CSmartStr(prevExpr) & (CSmartStr(fullName) & " IS NULL")
				end if
			end if
			if IsEqual(nextExpr,"") then
				nextExpr = " 1=0 "
			end if
			if IsEqual(prevExpr,"") then
				prevExpr = " 1=0 "
			end if
			if IsLess(0,i) then
				where_next = CSmartStr(where_next) & " AND "
				where_prev = CSmartStr(where_prev) & " AND "
			end if
			where_next = CSmartStr(where_next) & ("(" & CSmartStr(nextExpr))
			where_prev = CSmartStr(where_prev) & ("(" & CSmartStr(prevExpr))
			tail = CSmartStr(tail) & ")"
		loop while false
		if c_runnerpage_exitLoop43 then _
			exit do
		i = CSmartDbl(i)+1
	loop
	where_next = CSmartStr(where_next) & CSmartStr(tail)
	where_prev = CSmartStr(where_prev) & CSmartStr(tail)
	if ((IsEqual(where_next,"") or IsEqual(order_next,"")) or IsEqual(where_prev,"")) or IsEqual(order_prev,"") then
		setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_noNextPrev",1
		Exit Function
	end if
	if IsNull(query) then
		Exit Function
	end if
	if not bValue(query.HasGroupBy()) then
		doAssignmentByRef oWhere,query.Where()
		doAssignmentByRef where,whereAdd(where,oWhere.toSql_p1(query))
		doAssignmentByRef where_next,whereAdd(where_next,where)
		doAssignmentByRef where_prev,whereAdd(where_prev,where)
		query.ReplaceFieldsWithDummies_p1 GetBinaryFieldsIndices(this_object.tName)
		doAssignmentByRef sql_next,query.toSql_p2(where_next,order_next)
		doAssignmentByRef sql_prev,query.toSql_p2(where_prev,order_prev)
	else
		doAssignmentByRef oWhere,query.Where()
		doAssignmentByRef oHaving,query.Having()
		doAssignmentByRef where,whereAdd(where,oWhere.toSql_p1(query))
		doAssignmentByRef having,whereAdd(having,oHaving.toSql_p1(query))
		query.ReplaceFieldsWithDummies_p1 GetBinaryFieldsIndices(this_object.tName)
		sql = ("select * from (" & CSmartStr(query.toSql_p3(where,"",having))) & ") prevnextquery"
		sql_next = ((CSmartStr(sql) & " WHERE ") & CSmartStr(where_next)) & CSmartStr(order_next)
		sql_prev = ((CSmartStr(sql) & " WHERE ") & CSmartStr(where_prev)) & CSmartStr(order_prev)
	end if
	doAssignmentByRef sql_next,AddTop(sql_next,1)
	doAssignmentByRef sql_prev,AddTop(sql_prev,1)
	doAssignmentByRef res_next,db_query(sql_next,conn)
	if bValue(res_next) then
		if bValue(doAssignmentByRef(row_next,db_fetch_array(res_next))) then
			GetCollectionBounds tKeys,c_runnerpage_loopIdx44,c_runnerpage_loopMax44
			do while c_runnerpage_loopIdx44<=c_runnerpage_loopMax44
				i = GetCollectionKey(tKeys,c_runnerpage_loopIdx44)
				doAssignment k,ArrayElement(tKeys,i)
				setArrElement var_next,CSmartDbl(i)+1,ArrayElement(row_next,k)
				c_runnerpage_loopIdx44=c_runnerpage_loopIdx44+1
			loop
		end if
		db_closequery res_next
	end if
	doAssignmentByRef res_prev,db_query(sql_prev,conn)
	if bValue(doAssignmentByRef(row_prev,db_fetch_array(res_prev))) then
		GetCollectionBounds tKeys,c_runnerpage_loopIdx45,c_runnerpage_loopMax45
		do while c_runnerpage_loopIdx45<=c_runnerpage_loopMax45
			i = GetCollectionKey(tKeys,c_runnerpage_loopIdx45)
			doAssignment k,ArrayElement(tKeys,i)
			setArrElement prev,CSmartDbl(i)+1,ArrayElement(row_prev,k)
			c_runnerpage_loopIdx45=c_runnerpage_loopIdx45+1
		loop
	end if
	db_closequery res_prev
End Function
Function method_RunnerPage_doCaptchaCode(byref this_object)
	Dim var_SESSION,var_POST,sbmPage,controls,this
	if not (not IsEmpty(Session(CSmartStr(this_object.tName) & "_count_captcha"))) or IsLess(4,Session(CSmartStr(this_object.tName) & "_count_captcha")) then
		setArrElement Session,CSmartStr(this_object.tName) & "_count_captcha",0
	end if
	if (IsEqual(this_object.pageType,PAGE_EDIT) and IsEqual(GetRequestValue(RequestForm(),"a"),"edited") or IsEqual(this_object.pageType,PAGE_ADD) and IsEqual(GetRequestValue(RequestForm(),"a"),"added")) or IsEqual(this_object.pageType,PAGE_REGISTER) and IsEqual(GetRequestValue(RequestForm(),"btnSubmit"),"Register") then
		sbmPage = true
	else
		sbmPage = false
	end if
	if IsEqual(Session(CSmartStr(this_object.tName) & "_count_captcha"),0) or IsLess(4,Session(CSmartStr(this_object.tName) & "_count_captcha")) then
		if bValue(sbmPage) then
			if not IsEqual(asp_strtolower(postvalue("value_captcha_" & CSmartStr(this_object.id))),asp_strtolower(Session("captcha"))) then
				if IsEqual(this_object.pageType,PAGE_REGISTER) then
					doClassAssignment this_object,"captchaId",ArrayElement(ArrayElement(globalEvents.captchas,this_object.pageType),"strName")
					globalEvents.callCAPTCHA_p1 this_object
				else
					doClassAssignment this_object,"captchaId",ArrayElement(ArrayElement(this_object.eventsObject.captchas,this_object.pageType),"strName")
					this_object.eventsObject.callCAPTCHA_p1 this_object
				end if
				this_object.isCaptchaOk = false
			else
				this_object.isCaptchaOk = true
			end if
		else
			if IsEqual(this_object.pageType,PAGE_REGISTER) then
				doClassAssignment this_object,"captchaId",ArrayElement(ArrayElement(globalEvents.captchas,this_object.pageType),"strName")
				globalEvents.callCAPTCHA_p1 this_object
			else
				doClassAssignment this_object,"captchaId",ArrayElement(ArrayElement(this_object.eventsObject.captchas,this_object.pageType),"strName")
				this_object.eventsObject.callCAPTCHA_p1 this_object
			end if
		end if
		Set controls = (CreateDictionary1("controls",CreateDictionary()))
		setArrElementN controls,CreateArray2("controls","ctrlInd"),0
		setArrElementN controls,CreateArray2("controls","id"),this_object.id
		setArrElementN controls,CreateArray2("controls","fieldName"),"captcha"
		setArrElementN controls,CreateArray2("controls","mode"),this_object.pageType
		this_object.fillControlsMap_p1 controls
		this_object.addCaptchaToFieldSettings 
	else
		if bValue(sbmPage) then
			this_object.isCaptchaOk = true
		end if
	end if
End Function
Function method_RunnerPage_createCaptcha(byref this_object,ByVal params)
	Dim captchaHTML
	captchaHTML = ((((((("<div class=""captcha_block"">" & vbcrlf & _
		"			<object width=""210"" height=""65"" data=""securitycode.swf?ext=asp"" type=""application/x-shockwave-flash"">" & vbcrlf & _
		"				<param value=""securitycode.swf?ext=asp"" name=""movie""/>" & vbcrlf & _
		"				<param value=""opaque"" name=""wmode""/>" & vbcrlf & _
		"				<a href=""http://www.macromedia.com/go/getflashplayer""><img alt=""Download Flash"" src=""""/></a>" & vbcrlf & _
		"			</object>" & vbcrlf & _
		"				<div style=""white-space: nowrap;"">" & CSmartStr("Type the code you see above")) & ":</div>" & vbcrlf & _
		"			<span id=""edit") & CSmartStr(this_object.id)) & "_captcha_0"">" & vbcrlf & _
		"				<input type=""text"" value="""" name=""value_captcha_") & CSmartStr(this_object.id)) & """ style="""" id=""value_captcha_") & CSmartStr(this_object.id)) & """/>" & vbcrlf & _
		"				<font color=""red"">*</font>" & vbcrlf & _
		"			</span>"
	if PropertyExists(this_object,"isCaptchaOk") and not bValue(this_object.isCaptchaOk) then
		captchaHTML = CSmartStr(captchaHTML) & (("<div class=""error"">" & CSmartStr("Invalid security code.")) & "</div>")
	end if
	ResponseWrite CSmartStr(captchaHTML) & "</div>"
End Function
Function method_RunnerPage_createPerPage(byref this_object)
	Dim rpp,i
	rpp = ((("<span id=""recordspp_block" & CSmartStr(this_object.id)) & """ name=""recordspp_block") & CSmartStr(this_object.id)) & """>"
	if IsEqual(this_object.isTableType,"report") and bValue(this_object.reportGroupFields) then
		rpp = CSmartStr(rpp) & (((CSmartStr("Groups per page") & ":&nbsp;<select onChange=""javascript: document.location='") & CSmartStr(htmlspecialchars(asp_rawurlencode(this_object.shortTableName)))) & "_report.asp?pagesize='+this.options[this.selectedIndex].value;"">")
		i = 0
		do while IsLess(i,asp_count(this_object.arrGroupsPerPage))
			if not IsEqual(ArrayElement(this_object.arrGroupsPerPage,i),-1) then
				rpp = CSmartStr(rpp) & (((((("<option value=""" & CSmartStr(ArrayElement(this_object.arrGroupsPerPage,i))) & """ ") & CSmartStr(IIF(IsEqual(this_object.pageSize,ArrayElement(this_object.arrGroupsPerPage,i)),"selected",""))) & ">") & CSmartStr(ArrayElement(this_object.arrGroupsPerPage,i))) & "</option>")
			else
				rpp = CSmartStr(rpp) & (((("<option value=""-1"" " & CSmartStr(IIF(IsEqual(this_object.pageSize,ArrayElement(this_object.arrGroupsPerPage,i)),"selected",""))) & ">") & CSmartStr("Show all")) & "</option>")
			end if
			i = CSmartDbl(i)+1
		loop
	else
		rpp = CSmartStr(rpp) & (CSmartStr("Records Per Page") & ":&nbsp;")
		if IsEqual(this_object.isTableType,"report") then
			rpp = CSmartStr(rpp) & (("<select onChange=""javascript: document.location='" & CSmartStr(htmlspecialchars(asp_rawurlencode(this_object.shortTableName)))) & "_report.asp?pagesize='+this.options[this.selectedIndex].value;"">")
		else
			rpp = CSmartStr(rpp) & (("<select id=""recordspp" & CSmartStr(this_object.id)) & """>")
		end if
		i = 0
		do while IsLess(i,asp_count(this_object.arrRecsPerPage))
			if not IsEqual(ArrayElement(this_object.arrRecsPerPage,i),-1) then
				rpp = CSmartStr(rpp) & (((((("<option value=""" & CSmartStr(ArrayElement(this_object.arrRecsPerPage,i))) & """ ") & CSmartStr(IIF(IsEqual(this_object.pageSize,ArrayElement(this_object.arrRecsPerPage,i)),"selected",""))) & ">") & CSmartStr(ArrayElement(this_object.arrRecsPerPage,i))) & "</option>")
			else
				rpp = CSmartStr(rpp) & (((("<option value=""-1"" " & CSmartStr(IIF(IsEqual(this_object.pageSize,ArrayElement(this_object.arrRecsPerPage,i)),"selected",""))) & ">") & "Show all") & "</option>")
			end if
			i = CSmartDbl(i)+1
		loop
	end if
	rpp = CSmartStr(rpp) & "</select></span>"
	if IsEqual(this_object.isTableType,"report") then
		this_object.xt.assign_p2 "grpsPerPage",rpp
	else
		this_object.xt.assign_p2 "recsPerPage",rpp
	end if
End Function
Function method_RunnerPage_ProcessFiles(byref this_object)
	Dim f
	GetCollectionBounds this_object.filesToDelete,c_runnerpage_loopIdx48,c_runnerpage_loopMax48
	do while c_runnerpage_loopIdx48<=c_runnerpage_loopMax48
		c_runnerpage_arrKey48 = GetCollectionKey(this_object.filesToDelete,c_runnerpage_loopIdx48)
		doAssignment f,ArrayElement(this_object.filesToDelete,c_runnerpage_arrKey48)
		f.Delete 
		c_runnerpage_loopIdx48=c_runnerpage_loopIdx48+1
	loop
	GetCollectionBounds this_object.filesToMove,c_runnerpage_loopIdx49,c_runnerpage_loopMax49
	do while c_runnerpage_loopIdx49<=c_runnerpage_loopMax49
		c_runnerpage_arrKey49 = GetCollectionKey(this_object.filesToMove,c_runnerpage_loopIdx49)
		doAssignment f,ArrayElement(this_object.filesToMove,c_runnerpage_arrKey49)
		f.Move 
		c_runnerpage_loopIdx49=c_runnerpage_loopIdx49+1
	loop
	GetCollectionBounds this_object.filesToSave,c_runnerpage_loopIdx50,c_runnerpage_loopMax50
	do while c_runnerpage_loopIdx50<=c_runnerpage_loopMax50
		c_runnerpage_arrKey50 = GetCollectionKey(this_object.filesToSave,c_runnerpage_loopIdx50)
		doAssignment f,ArrayElement(this_object.filesToSave,c_runnerpage_arrKey50)
		f.Save 
		c_runnerpage_loopIdx50=c_runnerpage_loopIdx50+1
	loop
End Function
%>
