<%

'------ Class ListPage extends RunnerPage ------
Class ListPage
	Public recordsOnPage
	Public colsOnPage
	Public gsqlHead
	Public gsqlFrom
	Public gsqlWhereExpr
	Public gsqlGroupBy
	Public gsqlHaving
	Public querySQL
	Public subQueriesSupp
	Public arrFieldForSort
	Public arrHowFieldSort
	Public exportTo
	Public deleteRecs
	Public listFields
	Public totalsFields
	Public recSet
	Public nSecOptions
	Public isDynamicPerm
	Public isVerLayout
	Public masterKeysReq
	Public masterKeysByD
	Public detailKeysByD
	Public strOrderBy
	Public recNo
	Public maxPages
	Public maxRecs
	Public rowId
	Public myPage
	Public selectedRecs
	Public recordsDeleted
	Public rowsFound
	Public strWhereClause
	Public strHavingClause
	Public subQueriesSupAccess
	Public origTName
	Public isAddWebRep
	Public isUseInlineJs
	Public isUseInlineAdd
	Public isUseInlineEdit
	Public useCustomLabels
	Public globSearchFields
	Public panelSearchFields
	Public arrKeyFields
	Public numRowsFromSQL
	Public is508
	Public audit
	Public recIds
	Public noRecordsFirstPage
	Public tableGroupBy
	Public recsPerRowList
	Public delFile
	Public dbType
	Public rowHighlite
	Public mainTableOwnerID
	Public moveNext
	Public listIcons
	Public widthListIcons
	Public edit
	Public inlineEdit
	Public iCopy
	Public iView
	Public printFriendly
	Public createLoginPage
	Public includesArr
	Public searchControlBuilder
	Public templatefile
	Public searchPanel
	Public arrFieldSpanVal
	Public lockDelRec
	Public firstTime
	Public deleteMessage
	Public gMapFields
	Public nLoginMethod
	Public theSameFieldsType
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
	Public Function init_ListPage_p1(ByRef params)
		DoAssignmentByRef init_ListPage_p1,method_ListPage_ListPage(me,params)
	End Function
	Public Function addCommonHtml()
		DoAssignmentByRef addCommonHtml,method_ListPage_addCommonHtml(me)
	End Function
	Public Function displayMasterTableInfo()
		DoAssignmentByRef displayMasterTableInfo,method_ListPage_displayMasterTableInfo(me)
	End Function
	Public Function addCommonJs()
		DoAssignmentByRef addCommonJs,method_ListPage_addCommonJs(me)
	End Function
	Public Function addJsForGrid()
		DoAssignmentByRef addJsForGrid,method_ListPage_addJsForGrid(me)
	End Function
	Public Function prepareForResizeColumns()
		DoAssignmentByRef prepareForResizeColumns,method_ListPage_prepareForResizeColumns(me)
	End Function
	Public Function processMasterKeyValue()
		DoAssignmentByRef processMasterKeyValue,method_ListPage_processMasterKeyValue(me)
	End Function
	Public Function beforeProcessEvent()
		DoAssignmentByRef beforeProcessEvent,method_ListPage_beforeProcessEvent(me)
	End Function
	Public Function setSessionVariables()
		DoAssignmentByRef setSessionVariables,method_ListPage_setSessionVariables(me)
	End Function
	Public Function orderLinksAttr()
		DoAssignmentByRef orderLinksAttr,method_ListPage_orderLinksAttr(me)
	End Function
	Public Function setLinksAttr_p2(ByVal field,ByVal sort)
		DoAssignmentByRef setLinksAttr_p2,method_ListPage_setLinksAttr(me,field,sort)
	End Function
	Public Function setLinksAttr_p1(ByVal field)
		DoAssignmentByRef setLinksAttr_p1,method_ListPage_setLinksAttr(me,field,"")
	End Function
	Public Function buildOrderParams()
		DoAssignmentByRef buildOrderParams,method_ListPage_buildOrderParams(me)
	End Function
	Public Function deleteRecords()
		DoAssignmentByRef deleteRecords,method_ListPage_deleteRecords(me)
	End Function
	Public Function rulePRG()
		DoAssignmentByRef rulePRG,method_ListPage_rulePRG(me)
	End Function
	Public Function BeforeShowList()
		DoAssignmentByRef BeforeShowList,method_ListPage_BeforeShowList(me)
	End Function
	Public Function buildSearchPanel()
		DoAssignmentByRef buildSearchPanel,method_ListPage_buildSearchPanel(me)
	End Function
	Public Function assignAdmin()
		DoAssignmentByRef assignAdmin,method_ListPage_assignAdmin(me)
	End Function
	Public Function commonAssign()
		DoAssignmentByRef commonAssign,method_ListPage_commonAssign(me)
	End Function
	Public Function addAssignForGrid()
		DoAssignmentByRef addAssignForGrid,method_ListPage_addAssignForGrid(me)
	End Function
	Public Function createMapDiv_p1(ByRef params)
		DoAssignmentByRef createMapDiv_p1,method_ListPage_createMapDiv(me,params)
	End Function
	Public Function importLinksAttrs()
		DoAssignmentByRef importLinksAttrs,method_ListPage_importLinksAttrs(me)
	End Function
	Public Function inlineAddLinksAttrs()
		DoAssignmentByRef inlineAddLinksAttrs,method_ListPage_inlineAddLinksAttrs(me)
	End Function
	Public Function selectAllLinkAttrs()
		DoAssignmentByRef selectAllLinkAttrs,method_ListPage_selectAllLinkAttrs(me)
	End Function
	Public Function checkboxColumnAttrs()
		DoAssignmentByRef checkboxColumnAttrs,method_ListPage_checkboxColumnAttrs(me)
	End Function
	Public Function getPrintExportLinkAttrs_p1(ByVal page)
		DoAssignmentByRef getPrintExportLinkAttrs_p1,method_ListPage_getPrintExportLinkAttrs(me,page)
	End Function
	Public Function buttonShowHideStyle_p1(ByVal link)
		DoAssignmentByRef buttonShowHideStyle_p1,method_ListPage_buttonShowHideStyle(me,link)
	End Function
	Public Function buttonShowHideStyle()
		DoAssignmentByRef buttonShowHideStyle,method_ListPage_buttonShowHideStyle(me,"")
	End Function
	Public Function editSelectedLinkAttrs()
		DoAssignmentByRef editSelectedLinkAttrs,method_ListPage_editSelectedLinkAttrs(me)
	End Function
	Public Function saveAllLinkAttrs()
		DoAssignmentByRef saveAllLinkAttrs,method_ListPage_saveAllLinkAttrs(me)
	End Function
	Public Function cancelAllLinkAttrs()
		DoAssignmentByRef cancelAllLinkAttrs,method_ListPage_cancelAllLinkAttrs(me)
	End Function
	Public Function deleteSelectedLink()
		DoAssignmentByRef deleteSelectedLink,method_ListPage_deleteSelectedLink(me)
	End Function
	Public Function deleteSelectedAttrs()
		DoAssignmentByRef deleteSelectedAttrs,method_ListPage_deleteSelectedAttrs(me)
	End Function
	Public Function getFormInputsHTML()
		DoAssignmentByRef getFormInputsHTML,method_ListPage_getFormInputsHTML(me)
	End Function
	Public Function getFormTargetHTML()
		DoAssignmentByRef getFormTargetHTML,method_ListPage_getFormTargetHTML(me)
	End Function
	Public Function buildPagination()
		DoAssignmentByRef buildPagination,method_ListPage_buildPagination(me)
	End Function
	Public Function getPaginationLink_p3(ByVal pageNum,ByVal linkText,ByVal cls)
		DoAssignmentByRef getPaginationLink_p3,method_ListPage_getPaginationLink(me,pageNum,linkText,cls)
	End Function
	Public Function getPaginationLink_p2(ByVal pageNum,ByVal linkText)
		DoAssignmentByRef getPaginationLink_p2,method_ListPage_getPaginationLink(me,pageNum,linkText,false)
	End Function
	Public Function addWhereWithMasterTable()
		DoAssignmentByRef addWhereWithMasterTable,method_ListPage_addWhereWithMasterTable(me)
	End Function
	Public Function seekPageInRecSet_p1(ByVal strSQL)
		DoAssignmentByRef seekPageInRecSet_p1,method_ListPage_seekPageInRecSet(me,strSQL)
	End Function
	Public Function buildSQL()
		DoAssignmentByRef buildSQL,method_ListPage_buildSQL(me)
	End Function
	Public Function addMasterDetailSubQuery()
		DoAssignmentByRef addMasterDetailSubQuery,method_ListPage_addMasterDetailSubQuery(me)
	End Function
	Public Function fillGridShowInfo_p1(ByRef rowInfoArr)
		DoAssignmentByRef fillGridShowInfo_p1,method_ListPage_fillGridShowInfo(me,rowInfoArr)
	End Function
	Public Function beforeProccessRow()
		DoAssignmentByRef beforeProccessRow,method_ListPage_beforeProccessRow(me)
	End Function
	Public Function countWidthListIcons_p1(ByVal row)
		DoAssignmentByRef countWidthListIcons_p1,method_ListPage_countWidthListIcons(me,row)
	End Function
	Public Function assignListIconsColumn_p3(ByVal editPermis,ByVal addPermis,ByVal searchPermis)
		DoAssignmentByRef assignListIconsColumn_p3,method_ListPage_assignListIconsColumn(me,editPermis,addPermis,searchPermis)
	End Function
	Public Function assignListIconsColumn_p2(ByVal editPermis,ByVal addPermis)
		DoAssignmentByRef assignListIconsColumn_p2,method_ListPage_assignListIconsColumn(me,editPermis,addPermis,1)
	End Function
	Public Function assignListIconsColumn_p1(ByVal editPermis)
		DoAssignmentByRef assignListIconsColumn_p1,method_ListPage_assignListIconsColumn(me,editPermis,1,1)
	End Function
	Public Function assignListIconsColumn()
		DoAssignmentByRef assignListIconsColumn,method_ListPage_assignListIconsColumn(me,1,1,1)
	End Function
	Public Function fillGridData()
		DoAssignmentByRef fillGridData,method_ListPage_fillGridData(me)
	End Function
	Public Function setAttrAlign_p1(ByVal i)
		DoAssignmentByRef setAttrAlign_p1,method_ListPage_setAttrAlign(me,i)
	End Function
	Public Function countTotals_p2(ByRef totals,ByRef data)
		DoAssignmentByRef countTotals_p2,method_ListPage_countTotals(me,totals,data)
	End Function
	Public Function buildTotals_p1(ByRef totals)
		DoAssignmentByRef buildTotals_p1,method_ListPage_buildTotals(me,totals)
	End Function
	Public Function outputFieldValue_p2(ByVal field,ByVal state)
		DoAssignmentByRef outputFieldValue_p2,method_ListPage_outputFieldValue(me,field,state)
	End Function
	Public Function addSpanVal_p2(ByVal fName,ByRef data)
		DoAssignmentByRef addSpanVal_p2,method_ListPage_addSpanVal(me,fName,data)
	End Function
	Public Function addSpansForGridCells_p2(ByRef record,ByRef data)
		DoAssignmentByRef addSpansForGridCells_p2,method_ListPage_addSpansForGridCells(me,record,data)
	End Function
	Public Function proccessRecordValue_p3(ByRef data,ByRef keylink,ByVal listFieldInfo)
		DoAssignmentByRef proccessRecordValue_p3,method_ListPage_proccessRecordValue(me,data,keylink,listFieldInfo)
	End Function
	Public Function proccessDetailGridInfo_p2(ByRef record,ByRef data)
		DoAssignmentByRef proccessDetailGridInfo_p2,method_ListPage_proccessDetailGridInfo(me,record,data)
	End Function
	Public Function countDetailsRecsNoSubQ_p3(ByVal i,ByRef data,ByRef detailid)
		DoAssignmentByRef countDetailsRecsNoSubQ_p3,method_ListPage_countDetailsRecsNoSubQ(me,i,data,detailid)
	End Function
	Public Function checkDetailAndMasterFieldTypes()
		DoAssignmentByRef checkDetailAndMasterFieldTypes,method_ListPage_checkDetailAndMasterFieldTypes(me)
	End Function
	Public Function isDispGrid()
		DoAssignmentByRef isDispGrid,method_ListPage_isDispGrid(me)
	End Function
	Public Function fillCheckAttr_p3(ByRef record,ByVal data,ByVal keyblock)
		DoAssignmentByRef fillCheckAttr_p3,method_ListPage_fillCheckAttr(me,record,data,keyblock)
	End Function
	Public Function callJSCodeAfterRecordEdited()
		DoAssignmentByRef callJSCodeAfterRecordEdited,method_ListPage_callJSCodeAfterRecordEdited(me)
	End Function
	Public Function addDivSearchWin()
		DoAssignmentByRef addDivSearchWin,method_ListPage_addDivSearchWin(me)
	End Function
	Public Function prepareForBuildPage()
		DoAssignmentByRef prepareForBuildPage,method_ListPage_prepareForBuildPage(me)
	End Function
	Public Function showPage()
		DoAssignmentByRef showPage,method_ListPage_showPage(me)
	End Function
	Public Function createListPage_p2(ByVal table,ByVal options)
		DoAssignmentByRef createListPage_p2,method_ListPage_createListPage(me,table,options)
	End Function
	Public Function isAdminTable()
		DoAssignmentByRef isAdminTable,method_ListPage_isAdminTable(me)
	End Function
	Public Function clearSessionKeys()
		DoAssignmentByRef clearSessionKeys,method_RunnerPage_clearSessionKeys(me)
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
		setArrElement out,"recordsOnPage", recordsOnPage
		setArrElement out,"colsOnPage", colsOnPage
		setArrElement out,"gsqlHead", gsqlHead
		setArrElement out,"gsqlFrom", gsqlFrom
		setArrElement out,"gsqlWhereExpr", gsqlWhereExpr
		setArrElement out,"gsqlGroupBy", gsqlGroupBy
		setArrElement out,"gsqlHaving", gsqlHaving
		setArrElement out,"querySQL", querySQL
		setArrElement out,"subQueriesSupp", subQueriesSupp
		setArrElement out,"arrFieldForSort", arrFieldForSort
		setArrElement out,"arrHowFieldSort", arrHowFieldSort
		setArrElement out,"exportTo", exportTo
		setArrElement out,"deleteRecs", deleteRecs
		setArrElement out,"listFields", listFields
		setArrElement out,"totalsFields", totalsFields
		setArrElement out,"recSet", recSet
		setArrElement out,"nSecOptions", nSecOptions
		setArrElement out,"isDynamicPerm", isDynamicPerm
		setArrElement out,"isVerLayout", isVerLayout
		setArrElement out,"masterKeysReq", masterKeysReq
		setArrElement out,"masterKeysByD", masterKeysByD
		setArrElement out,"detailKeysByD", detailKeysByD
		setArrElement out,"strOrderBy", strOrderBy
		setArrElement out,"recNo", recNo
		setArrElement out,"maxPages", maxPages
		setArrElement out,"maxRecs", maxRecs
		setArrElement out,"rowId", rowId
		setArrElement out,"myPage", myPage
		setArrElement out,"selectedRecs", selectedRecs
		setArrElement out,"recordsDeleted", recordsDeleted
		setArrElement out,"rowsFound", rowsFound
		setArrElement out,"strWhereClause", strWhereClause
		setArrElement out,"strHavingClause", strHavingClause
		setArrElement out,"subQueriesSupAccess", subQueriesSupAccess
		setArrElement out,"origTName", origTName
		setArrElement out,"isAddWebRep", isAddWebRep
		setArrElement out,"isUseInlineJs", isUseInlineJs
		setArrElement out,"isUseInlineAdd", isUseInlineAdd
		setArrElement out,"isUseInlineEdit", isUseInlineEdit
		setArrElement out,"useCustomLabels", useCustomLabels
		setArrElement out,"globSearchFields", globSearchFields
		setArrElement out,"panelSearchFields", panelSearchFields
		setArrElement out,"arrKeyFields", arrKeyFields
		setArrElement out,"numRowsFromSQL", numRowsFromSQL
		setArrElement out,"is508", is508
		setArrElement out,"audit", audit
		setArrElement out,"recIds", recIds
		setArrElement out,"noRecordsFirstPage", noRecordsFirstPage
		setArrElement out,"tableGroupBy", tableGroupBy
		setArrElement out,"recsPerRowList", recsPerRowList
		setArrElement out,"delFile", delFile
		setArrElement out,"dbType", dbType
		setArrElement out,"rowHighlite", rowHighlite
		setArrElement out,"mainTableOwnerID", mainTableOwnerID
		setArrElement out,"moveNext", moveNext
		setArrElement out,"listIcons", listIcons
		setArrElement out,"widthListIcons", widthListIcons
		setArrElement out,"edit", edit
		setArrElement out,"inlineEdit", inlineEdit
		setArrElement out,"iCopy", iCopy
		setArrElement out,"iView", iView
		setArrElement out,"printFriendly", printFriendly
		setArrElement out,"createLoginPage", createLoginPage
		setArrElement out,"includesArr", includesArr
		setArrElement out,"searchControlBuilder", searchControlBuilder
		setArrElement out,"templatefile", templatefile
		setArrElement out,"searchPanel", searchPanel
		setArrElement out,"arrFieldSpanVal", arrFieldSpanVal
		setArrElement out,"lockDelRec", lockDelRec
		setArrElement out,"firstTime", firstTime
		setArrElement out,"deleteMessage", deleteMessage
		setArrElement out,"gMapFields", gMapFields
		setArrElement out,"nLoginMethod", nLoginMethod
		setArrElement out,"theSameFieldsType", theSameFieldsType
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
		DoAssignment recordsOnPage, dict("recordsOnPage")
		DoAssignment colsOnPage, dict("colsOnPage")
		DoAssignment gsqlHead, dict("gsqlHead")
		DoAssignment gsqlFrom, dict("gsqlFrom")
		DoAssignment gsqlWhereExpr, dict("gsqlWhereExpr")
		DoAssignment gsqlGroupBy, dict("gsqlGroupBy")
		DoAssignment gsqlHaving, dict("gsqlHaving")
		DoAssignment querySQL, dict("querySQL")
		DoAssignment subQueriesSupp, dict("subQueriesSupp")
		DoAssignment arrFieldForSort, dict("arrFieldForSort")
		DoAssignment arrHowFieldSort, dict("arrHowFieldSort")
		DoAssignment exportTo, dict("exportTo")
		DoAssignment deleteRecs, dict("deleteRecs")
		DoAssignment listFields, dict("listFields")
		DoAssignment totalsFields, dict("totalsFields")
		DoAssignment recSet, dict("recSet")
		DoAssignment nSecOptions, dict("nSecOptions")
		DoAssignment isDynamicPerm, dict("isDynamicPerm")
		DoAssignment isVerLayout, dict("isVerLayout")
		DoAssignment masterKeysReq, dict("masterKeysReq")
		DoAssignment masterKeysByD, dict("masterKeysByD")
		DoAssignment detailKeysByD, dict("detailKeysByD")
		DoAssignment strOrderBy, dict("strOrderBy")
		DoAssignment recNo, dict("recNo")
		DoAssignment maxPages, dict("maxPages")
		DoAssignment maxRecs, dict("maxRecs")
		DoAssignment rowId, dict("rowId")
		DoAssignment myPage, dict("myPage")
		DoAssignment selectedRecs, dict("selectedRecs")
		DoAssignment recordsDeleted, dict("recordsDeleted")
		DoAssignment rowsFound, dict("rowsFound")
		DoAssignment strWhereClause, dict("strWhereClause")
		DoAssignment strHavingClause, dict("strHavingClause")
		DoAssignment subQueriesSupAccess, dict("subQueriesSupAccess")
		DoAssignment origTName, dict("origTName")
		DoAssignment isAddWebRep, dict("isAddWebRep")
		DoAssignment isUseInlineJs, dict("isUseInlineJs")
		DoAssignment isUseInlineAdd, dict("isUseInlineAdd")
		DoAssignment isUseInlineEdit, dict("isUseInlineEdit")
		DoAssignment useCustomLabels, dict("useCustomLabels")
		DoAssignment globSearchFields, dict("globSearchFields")
		DoAssignment panelSearchFields, dict("panelSearchFields")
		DoAssignment arrKeyFields, dict("arrKeyFields")
		DoAssignment numRowsFromSQL, dict("numRowsFromSQL")
		DoAssignment is508, dict("is508")
		DoAssignment audit, dict("audit")
		DoAssignment recIds, dict("recIds")
		DoAssignment noRecordsFirstPage, dict("noRecordsFirstPage")
		DoAssignment tableGroupBy, dict("tableGroupBy")
		DoAssignment recsPerRowList, dict("recsPerRowList")
		DoAssignment delFile, dict("delFile")
		DoAssignment dbType, dict("dbType")
		DoAssignment rowHighlite, dict("rowHighlite")
		DoAssignment mainTableOwnerID, dict("mainTableOwnerID")
		DoAssignment moveNext, dict("moveNext")
		DoAssignment listIcons, dict("listIcons")
		DoAssignment widthListIcons, dict("widthListIcons")
		DoAssignment edit, dict("edit")
		DoAssignment inlineEdit, dict("inlineEdit")
		DoAssignment iCopy, dict("iCopy")
		DoAssignment iView, dict("iView")
		DoAssignment printFriendly, dict("printFriendly")
		DoAssignment createLoginPage, dict("createLoginPage")
		DoAssignment includesArr, dict("includesArr")
		DoAssignment searchControlBuilder, dict("searchControlBuilder")
		DoAssignment templatefile, dict("templatefile")
		DoAssignment searchPanel, dict("searchPanel")
		DoAssignment arrFieldSpanVal, dict("arrFieldSpanVal")
		DoAssignment lockDelRec, dict("lockDelRec")
		DoAssignment firstTime, dict("firstTime")
		DoAssignment deleteMessage, dict("deleteMessage")
		DoAssignment gMapFields, dict("gMapFields")
		DoAssignment nLoginMethod, dict("nLoginMethod")
		DoAssignment theSameFieldsType, dict("theSameFieldsType")
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
' ListPage implementation 
Function method_ListPage_ListPage(byref this_object,ByRef params)
	this_object.recordsOnPage = 0
	this_object.colsOnPage = 1
	this_object.gsqlHead = ""
	this_object.gsqlFrom = ""
	this_object.gsqlWhereExpr = ""
	this_object.gsqlGroupBy = ""
	this_object.gsqlHaving = ""
	this_object.querySQL = ""
	this_object.subQueriesSupp = true
	doClassAssignment this_object,"arrFieldForSort",CreateDictionary()
	doClassAssignment this_object,"arrHowFieldSort",CreateDictionary()
	this_object.exportTo = false
	this_object.deleteRecs = false
	doClassAssignment this_object,"listFields",CreateDictionary()
	doClassAssignment this_object,"totalsFields",CreateDictionary()
	this_object.recSet = null
	this_object.nSecOptions = 0
	this_object.isDynamicPerm = false
	this_object.isVerLayout = false
	doClassAssignment this_object,"masterKeysReq",CreateDictionary()
	doClassAssignment this_object,"masterKeysByD",CreateDictionary()
	doClassAssignment this_object,"detailKeysByD",CreateDictionary()
	this_object.strOrderBy = ""
	this_object.recNo = 1
	this_object.maxPages = 1
	this_object.maxRecs = 0
	this_object.rowId = 0
	this_object.myPage = 0
	doClassAssignment this_object,"selectedRecs",CreateDictionary()
	this_object.recordsDeleted = 0
	this_object.rowsFound = false
	this_object.strWhereClause = ""
	this_object.strHavingClause = ""
	this_object.subQueriesSupAccess = false
	this_object.origTName = ""
	this_object.isAddWebRep = true
	this_object.isUseInlineJs = false
	this_object.isUseInlineAdd = false
	this_object.isUseInlineEdit = false
	this_object.useCustomLabels = false
	doClassAssignment this_object,"globSearchFields",CreateDictionary()
	doClassAssignment this_object,"panelSearchFields",CreateDictionary()
	doClassAssignment this_object,"arrKeyFields",CreateDictionary()
	this_object.numRowsFromSQL = 0
	this_object.is508 = false
	this_object.audit = null
	doClassAssignment this_object,"recIds",CreateDictionary()
	this_object.noRecordsFirstPage = false
	this_object.tableGroupBy = false
	this_object.recsPerRowList = 0
	this_object.delFile = false
	this_object.dbType = 0
	this_object.rowHighlite = false
	this_object.mainTableOwnerID = ""
	this_object.moveNext = 0
	this_object.listIcons = false
	this_object.widthListIcons = 0
	this_object.edit = false
	this_object.inlineEdit = false
	this_object.iCopy = false
	this_object.iView = false
	this_object.printFriendly = false
	this_object.createLoginPage = false
	doClassAssignment this_object,"includesArr",CreateDictionary()
	this_object.searchControlBuilder = null
	this_object.templatefile = ""
	this_object.searchPanel = null
	doClassAssignment this_object,"arrFieldSpanVal",CreateDictionary()
	this_object.firstTime = 0
	this_object.deleteMessage = ""
	doClassAssignment this_object,"gMapFields",CreateDictionary()
	this_object.theSameFieldsType = false
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
	Dim this,i,tField
	method_RunnerPage_RunnerPage this_object,params
	this_object.beforeProcessEvent 
	this_object.processMasterKeyValue 
	this_object.setLangParams 
	doClassAssignment this_object,"allDetailsTablesArr",GetDetailTablesArr(this_object.tName)
	i = 0
	do while IsLess(i,asp_count(this_object.allDetailsTablesArr))
		asp_include getabspath(("include/" & CSmartStr(ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dShortTable"))) & "_settings.asp"),true
		setArrElement this_object.permis,ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dDataSourceTable"),this_object.getPermissions_p1(ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dDataSourceTable"))
		setArrElement this_object.masterKeysByD,i,GetMasterKeysByDetailTable(ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dDataSourceTable"),this_object.tName,CreateDictionary())
		setArrElement this_object.detailKeysByD,i,GetDetailKeysByDetailTable(ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dDataSourceTable"),this_object.tName,CreateDictionary())
		i = CSmartDbl(i)+1
	loop
	doClassAssignment this_object,"theSameFieldsType",this_object.checkDetailAndMasterFieldTypes()
	this_object.genId 
	doClassAssignment this_object,"is508",isEnableSection508()
	this_object.templatefile = CSmartStr(this_object.shortTableName) & "_list.htm"
	GetCollectionBounds this_object.totalsFields,c_listpage_loopIdx3,c_listpage_loopMax3
	do while c_listpage_loopIdx3<=c_listpage_loopMax3
		c_listpage_arrKey3 = GetCollectionKey(this_object.totalsFields,c_listpage_loopIdx3)
		doAssignment tField,ArrayElement(this_object.totalsFields,c_listpage_arrKey3)
		setArrElementN this_object.jsSettings,CreateArray4("tableSettings",this_object.tName,"totalFields",empty),CreateDictionary3("type",ArrayElement(tField,"totalsType"),"fName",ArrayElement(tField,"fName"),"format",ArrayElement(tField,"viewFormat"))
		if IsEqual(ArrayElement(tField,"totalsType"),"COUNT") then
			this_object.outputFieldValue_p2 ArrayElement(tField,"fName"),1
		else
			this_object.outputFieldValue_p2 ArrayElement(tField,"fName"),2
		end if
		c_listpage_loopIdx3=c_listpage_loopIdx3+1
	loop
	setArrElementN this_object.jsSettings,CreateArray3("tableSettings",this_object.tName,"permissions"),ArrayElement(this_object.permis,this_object.tName)
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","isUseInlineEdit"),CreateDictionary2("default",false,"jsName","isInlineEdit")
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","isUseInlineAdd"),CreateDictionary2("default",false,"jsName","isInlineAdd")
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","copy"),CreateDictionary2("default",false,"jsName","copy")
	setArrElementN this_object.settingsMap,CreateArray2("tableSettings","view"),CreateDictionary2("default",false,"jsName","view")
	i = 0
	do while IsLess(i,asp_count(this_object.listFields))
		setArrElementN this_object.jsSettings,CreateArray4("tableSettings",this_object.tName,"listFields",empty),ArrayElement(ArrayElement(this_object.listFields,i),"fName")
		if IsEqual(ArrayElement(ArrayElement(this_object.listFields,i),"viewFormat"),FORMAT_MAP) then
			setArrElement this_object.gMapFields,asp_count(this_object.gMapFields),i
		end if
		i = CSmartDbl(i)+1
	loop
End Function
Function method_ListPage_addCommonHtml(byref this_object)
	Dim this
	setArrElement this_object.body,"begin",CSmartStr(ArrayElement(this_object.body,"begin")) & "<div id=""search_suggest"" class=""search_suggest""></div>"
	setArrElement this_object.body,"begin",CSmartStr(ArrayElement(this_object.body,"begin")) & "<div id=""master_details"" onmouseover=""RollDetailsLink.showPopup();"" onmouseout=""RollDetailsLink.hidePopup();""> </div>"
	if bValue(this_object.is508) then
		setArrElement this_object.body,"begin",CSmartStr(ArrayElement(this_object.body,"begin")) & (((("<a href=""#skipdata"" title=""" & CSmartStr("Skip to table data")) & """ style=""width:1px;height:1px;overflow:hidden;display:block;"">") & CSmartStr("Skip to table data")) & "</a>")
		setArrElement this_object.body,"begin",CSmartStr(ArrayElement(this_object.body,"begin")) & (((("<a href=""#skipmenu"" title=""" & CSmartStr("Skip to menu")) & """ style=""width:1px;height:1px;overflow:hidden;display:block;"">") & CSmartStr("Skip to menu")) & "</a>")
		setArrElement this_object.body,"begin",CSmartStr(ArrayElement(this_object.body,"begin")) & (((("<a href=""#skipsearch"" title=""" & CSmartStr("Skip to search")) & """ style=""width:1px;height:1px;overflow:hidden;display:block;"">") & CSmartStr("Skip to search")) & "</a>")
		setArrElement this_object.body,"begin",CSmartStr(ArrayElement(this_object.body,"begin")) & (((("<a href=""templates/helpshortcut.htm"" title=""" & CSmartStr("Hotkeys reference")) & """ style=""width:1px;height:1px;overflow:hidden;display:block;"">") & CSmartStr("Hotkeys reference")) & "</a>")
	end if
	this_object.displayMasterTableInfo 
End Function
Function method_ListPage_displayMasterTableInfo(byref this_object)
	Dim masterTablesInfoArr,i,detailKeys,j,masterKeys,var_SESSION,params,master
	doAssignmentByRef masterTablesInfoArr,GetMasterTablesArr(this_object.tName)
	i = 0
	do while IsLess(i,asp_count(masterTablesInfoArr))
		if IsEqual(this_object.masterTable,ArrayElement(ArrayElement(masterTablesInfoArr,i),"mDataSourceTable")) then
			if bValue(ArrayElement(ArrayElement(masterTablesInfoArr,i),"dispInfo")) then
				doAssignment detailKeys,ArrayElement(ArrayElement(masterTablesInfoArr,i),"detailKeys")
				j = 0
				do while IsLess(j,asp_count(detailKeys))
					setArrElement masterKeys,asp_count(masterKeys),Session((CSmartStr(this_object.sessionPrefix) & "_masterkey") & CSmartStr(CSmartDbl(j)+1))
					j = CSmartDbl(j)+1
				loop
				Set params = (CreateDictionary2("detailtable",this_object.tName,"keys",masterKeys))
				Set master = (CreateDictionary())
				setArrElement master,"func","DisplayMasterTableInfo_" & CSmartStr(ArrayElement(ArrayElement(masterTablesInfoArr,i),"mShortTable"))
				setArrElement master,"params",params
				this_object.xt.assignbyref_p2 "showmasterfile",master
			end if
			this_object.xt.assign_p2 "mastertable_block",true
			this_object.xt.assign_p2 "backtomasterlink_attrs",("href=""" & CSmartStr(ArrayElement(ArrayElement(masterTablesInfoArr,i),"mShortTable"))) & "_list.asp?a=return"""
			this_object.xt.assign_p2 "backtomasterlink_caption",GetTableCaption(GoodFieldName(ArrayElement(ArrayElement(masterTablesInfoArr,i),"mDataSourceTable")))
		end if
		i = CSmartDbl(i)+1
	loop
End Function
Function method_ListPage_addCommonJs(byref this_object)
	Dim this,onLoadJsCode,i
	method_RunnerPage_addCommonJs this_object
	if bValue(this_object.useDetailsPreview) then
		this_object.AddJSFile_p1 "include/detailspreview"
	end if
	this_object.addJsForGrid 
	this_object.addButtonHandlers 
	doAssignmentByRef onLoadJsCode,GetTableData(this_object.tName,".jsOnloadList","")
	this_object.addOnLoadJsEvent_p1 onLoadJsCode
End Function
Function method_ListPage_addJsForGrid(byref this_object)
	Dim this,hLite,lIcons,rCol
	if bValue(this_object.isResizeColumns) then
		this_object.prepareForResizeColumns 
	end if
	this_object.callJSCodeAfterRecordEdited 
	if bValue(this_object.is508) then
		if not bValue(this_object.isVerLayout) then
			doAssignment hLite,IIF(this_object.rowHighlite,"true","false")
			doAssignment lIcons,IIF(this_object.listIcons,"true","false")
			doAssignment rCol,IIF(this_object.isResizeColumns,"true","false")
			this_object.AddJsCode_p1 (((((((vblf & _
				"setHoverForTR(false," & CSmartStr(this_object.id)) & ",") & CSmartStr(hLite)) & ",") & CSmartStr(lIcons)) & ",") & CSmartStr(rCol)) & ");"
		end if
	end if
	if bValue(this_object.useIbox) then
		this_object.AddJSFile_p1 "include/ibox"
		this_object.AddCSSFile_p1 "include/ibox"
		this_object.AddJsCode_p1 vblf & _
			"init_ibox();"
	end if
	this_object.initGmaps 
End Function
Function method_ListPage_prepareForResizeColumns(byref this_object)
	Dim this
	if not IsEqual(this_object.mode,LIST_AJAX) then
		this_object.AddJSFile_p1 "include/resizebleGrid"
		this_object.AddCSSFile_p1 "include/stylesheets"
	end if
	setArrElementN this_object.jsSettings,CreateArray3("tableSettings",this_object.tName,"numRows"),IIF(this_object.numRowsFromSQL,true,false)
End Function
Function method_ListPage_processMasterKeyValue(byref this_object)
	Dim i,var_SESSION
	if bValue(asp_count(this_object.masterKeysReq)) then
		i = 1
		do while IsLessOrEqual(i,asp_count(this_object.masterKeysReq))
			setArrElement Session,(CSmartStr(this_object.sessionPrefix) & "_masterkey") & CSmartStr(i),ArrayElement(this_object.masterKeysReq,i)
			i = CSmartDbl(i)+1
		loop
		if not IsEmpty(Session((CSmartStr(this_object.sessionPrefix) & "_masterkey") & CSmartStr(i))) then
			asp_unsetElement Session,(CSmartStr(this_object.sessionPrefix) & "_masterkey") & CSmartStr(i)
		end if
	else
		if bValue(asp_count(this_object.detailKeysByM)) then
			i = 0
			do while IsLess(i,asp_count(this_object.detailKeysByM))
				if not IsEmpty(Session((CSmartStr(this_object.sessionPrefix) & "_masterkey") & CSmartStr(CSmartDbl(i)+1))) then
					setArrElement this_object.masterKeysReq,CSmartDbl(i)+1,Session((CSmartStr(this_object.sessionPrefix) & "_masterkey") & CSmartStr(CSmartDbl(i)+1))
				end if
				i = CSmartDbl(i)+1
			loop
		end if
	end if
	setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_search",0
	if bValue(this_object.firstTime) then
		setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_pagenumber",1
	end if
End Function
Function method_ListPage_beforeProcessEvent(byref this_object)
	Dim this
	if bValue(this_object.eventExists_p1("BeforeProcessList")) then
		this_object.eventsObject.BeforeProcessList_p1 this_object.conn
	end if
End Function
Function method_ListPage_setSessionVariables(byref this_object)
	Dim var_SESSION,var_REQUEST
	method_RunnerPage_setSessionVariables this_object
	this_object.searchClauseObj.parseRequest 
	setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_advsearch",serialize(this_object.searchClauseObj)
	if bValue(GetRequestValue(Request,"orderby")) then
		setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_orderby",GetRequestValue(Request,"orderby")
	end if
	if bValue(GetRequestValue(Request,"goto")) then
		setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_pagenumber",GetRequestValue(Request,"goto")
	end if
	this_object.myPage = CSmartLng(Session(CSmartStr(this_object.sessionPrefix) & "_pagenumber"))
	if not bValue(this_object.myPage) then
		this_object.myPage = 1
	end if
	if not bValue(this_object.pageSize) then
		doClassAssignment this_object,"pageSize",this_object.gPageSize
	end if
End Function
Function method_ListPage_orderLinksAttr(byref this_object)
	Dim i
	i = 0
	do while IsLess(i,asp_count(this_object.listFields))
		this_object.xt.assign_p2 CSmartStr(GoodFieldName(ArrayElement(ArrayElement(this_object.listFields,i),"fName"))) & "_orderlinkattrs",this_object.setLinksAttr_p1(GoodFieldName(ArrayElement(ArrayElement(this_object.listFields,i),"fName")))
		this_object.xt.assign_p2 CSmartStr(GoodFieldName(ArrayElement(ArrayElement(this_object.listFields,i),"fName"))) & "_fieldheader",true
		i = CSmartDbl(i)+1
	loop
End Function
Function method_ListPage_setLinksAttr(byref this_object,ByVal field,ByVal sort)
	Dim href,orderlinkattrs
	href = ((CSmartStr(this_object.shortTableName) & "_list.asp?orderby=") & CSmartStr(IIF(not IsEqual(sort,""),IIF(IsEqual(sort,"a"),"d","a"),"a"))) & CSmartStr(field)
	orderlinkattrs = ((((((((("id=""order_" & CSmartStr(field)) & "_") & CSmartStr(this_object.id)) & """ name=""order_") & CSmartStr(field)) & "_") & CSmartStr(this_object.id)) & """ href=""") & CSmartStr(href)) & """"
	doAssignmentByRef method_ListPage_setLinksAttr,orderlinkattrs
	Exit Function
End Function
Function method_ListPage_buildOrderParams(byref this_object)
	Dim this,order_ind,arrImplicitSortFields,key,var_SESSION,tKeys,i,lenkey,idsearch,lenArr,order_field,order_dir,var_REQUEST,condition,orderlinkattrs
	this_object.orderLinksAttr 
	order_ind = -1
	doClassAssignment this_object,"arrFieldForSort",CreateDictionary()
	doClassAssignment this_object,"arrHowFieldSort",CreateDictionary()
	Set arrImplicitSortFields = (CreateDictionary())
	Set key = (CreateDictionary())
	if bValue(Session(CSmartStr(this_object.sessionPrefix) & "_arrFieldForSort")) then
		doClassAssignment this_object,"arrFieldForSort",Session(CSmartStr(this_object.sessionPrefix) & "_arrFieldForSort")
	end if
	if bValue(Session(CSmartStr(this_object.sessionPrefix) & "_arrHowFieldSort")) then
		doClassAssignment this_object,"arrHowFieldSort",Session(CSmartStr(this_object.sessionPrefix) & "_arrHowFieldSort")
	end if
	if bValue(Session(CSmartStr(this_object.sessionPrefix) & "_key")) then
		doAssignment key,Session(CSmartStr(this_object.sessionPrefix) & "_key")
	else
		doAssignmentByRef tKeys,GetTableKeys(this_object.tName)
		i = 0
		do while IsLess(i,asp_count(tKeys))
			if bValue(GetFieldIndex(ArrayElement(tKeys,i),"")) then
				setArrElement key,asp_count(key),GetFieldIndex(ArrayElement(tKeys,i),"")
			end if
			i = CSmartDbl(i)+1
		loop
		setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_key",key
	end if
	doAssignmentByRef lenkey,asp_count(key)
	if not (not IsEmpty(Session(CSmartStr(this_object.sessionPrefix) & "_order"))) then
		doClassAssignment this_object,"arrFieldForSort",CreateDictionary()
		doClassAssignment this_object,"arrHowFieldSort",CreateDictionary()
		if bValue(asp_count(this_object.gOrderIndexes)) then
			i = 0
			do while IsLess(i,asp_count(this_object.gOrderIndexes))
				setArrElement this_object.arrFieldForSort,asp_count(this_object.arrFieldForSort),ArrayElement(ArrayElement(this_object.gOrderIndexes,i),0)
				setArrElement this_object.arrHowFieldSort,asp_count(this_object.arrHowFieldSort),ArrayElement(ArrayElement(this_object.gOrderIndexes,i),1)
				i = CSmartDbl(i)+1
			loop
		else
			if not IsEqual(this_object.gstrOrderBy,"") then
				setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_noNextPrev",1
			end if
		end if
		if bValue(asp_count(key)) and bValue(this_object.moveNext) then
			i = 0
			do while IsLess(i,lenkey)
				doAssignmentByRef idsearch,asp_array_search(ArrayElement(key,i),this_object.arrFieldForSort,false)
				if IsFalse(idsearch) then
					setArrElement this_object.arrFieldForSort,asp_count(this_object.arrFieldForSort),ArrayElement(key,i)
					setArrElement this_object.arrHowFieldSort,asp_count(this_object.arrHowFieldSort),"ASC"
					setArrElement arrImplicitSortFields,asp_count(arrImplicitSortFields),ArrayElement(key,i)
				end if
				i = CSmartDbl(i)+1
			loop
		end if
	end if
	doAssignmentByRef lenArr,asp_count(this_object.arrFieldForSort)
	if bValue(Session(CSmartStr(this_object.sessionPrefix) & "_orderby")) then
		doAssignmentByRef order_field,GetFieldByGoodFieldName(asp_substr(Session(CSmartStr(this_object.sessionPrefix) & "_orderby"),1,empty),"")
		doAssignmentByRef order_dir,asp_substr(Session(CSmartStr(this_object.sessionPrefix) & "_orderby"),0,1)
		doAssignmentByRef order_ind,GetFieldIndex(order_field,"")
		if bValue(order_ind) then
			if (not bValue(GetRequestValue(Request,"a")) and not bValue(GetRequestValue(Request,"goto"))) and not bValue(GetRequestValue(Request,"pagesize")) then
				if bValue(GetRequestValue(Request,"ctrl")) then
					doAssignmentByRef idsearch,asp_array_search(order_ind,this_object.arrFieldForSort,false)
					if IsFalse(idsearch) then
						if IsEqual(order_dir,"a") then
							setArrElement this_object.arrFieldForSort,asp_count(this_object.arrFieldForSort),order_ind
							setArrElement this_object.arrHowFieldSort,asp_count(this_object.arrHowFieldSort),"ASC"
						else
							setArrElement this_object.arrFieldForSort,asp_count(this_object.arrFieldForSort),order_ind
							setArrElement this_object.arrHowFieldSort,asp_count(this_object.arrHowFieldSort),"DESC"
						end if
					else
						setArrElement this_object.arrHowFieldSort,idsearch,IIF(IsEqual(order_dir,"a"),"ASC","DESC")
					end if
				else
					doClassAssignment this_object,"arrFieldForSort",CreateDictionary()
					doClassAssignment this_object,"arrHowFieldSort",CreateDictionary()
					if not (not bValue(Session(CSmartStr(this_object.sessionPrefix) & "_orderNo"))) then
						asp_unsetElement Session,CSmartStr(this_object.sessionPrefix) & "_orderNo"
					end if
					setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_noNextPrev",0
					if IsEqual(order_dir,"a") then
						setArrElement this_object.arrFieldForSort,asp_count(this_object.arrFieldForSort),order_ind
						setArrElement this_object.arrHowFieldSort,asp_count(this_object.arrHowFieldSort),"ASC"
					else
						setArrElement this_object.arrFieldForSort,asp_count(this_object.arrFieldForSort),order_ind
						setArrElement this_object.arrHowFieldSort,asp_count(this_object.arrHowFieldSort),"DESC"
					end if
				end if
			end if
		end if
	end if
	doAssignmentByRef lenArr,asp_count(this_object.arrFieldForSort)
	condition = true
	if not bValue(asp_count(this_object.arrFieldForSort)) then
		condition = false
	else
		if not bValue(ArrayElement(this_object.arrFieldForSort,0)) then
			condition = false
		end if
	end if
	if bValue(condition) then
		i = 0
		do while IsLess(i,lenArr)
			doAssignmentByRef order_field,GetFieldByIndex(ArrayElement(this_object.arrFieldForSort,i),"")
			doAssignment order_dir,IIF(IsEqual(ArrayElement(this_object.arrHowFieldSort,i),"ASC"),"a","d")
			doAssignmentByRef idsearch,asp_array_search(ArrayElement(this_object.arrFieldForSort,i),arrImplicitSortFields,false)
			if IsFalse(idsearch) then
				this_object.xt.assign_section_p3 CSmartStr(GoodFieldName(order_field)) & "_fieldheader","",((("<img " & CSmartStr(IIF(IsEqual(this_object.is508,true),"alt="" "" ",""))) & "src=""images/") & CSmartStr(IIF(IsEqual(order_dir,"a"),"up","down"))) & ".gif"" border=0>"
			end if
			if not IsFalse(idsearch) and bValue(asp_in_array(GetFieldIndex(order_field,""),arrImplicitSortFields,false)) then
				doAssignmentByRef orderlinkattrs,this_object.setLinksAttr_p1(GoodFieldName(order_field))
			else
				doAssignmentByRef orderlinkattrs,this_object.setLinksAttr_p2(GoodFieldName(order_field),order_dir)
			end if
			this_object.xt.assign_p2 CSmartStr(GoodFieldName(order_field)) & "_orderlinkattrs",orderlinkattrs
			i = CSmartDbl(i)+1
		loop
	end if
	if IsLess(0,lenArr) then
		i = 0
		do while IsLess(i,lenArr)
			this_object.strOrderBy = CSmartStr(this_object.strOrderBy) & CSmartStr(IIF(GetFieldByIndex(ArrayElement(this_object.arrFieldForSort,i),""),((CSmartStr(IIF(not IsEqual(this_object.strOrderBy,""),", "," ORDER BY ")) & CSmartStr(ArrayElement(this_object.arrFieldForSort,i))) & " ") & CSmartStr(ArrayElement(this_object.arrHowFieldSort,i)),""))
			i = CSmartDbl(i)+1
		loop
	end if
	if IsEqual(Session(CSmartStr(this_object.sessionPrefix) & "_noNextPrev"),1) then
		doClassAssignment this_object,"strOrderBy",this_object.gstrOrderBy
	end if
End Function
Function method_ListPage_deleteRecords(byref this_object)
	Dim var_REQUEST,i,keys,ind,arr,keyblock,where,strSQl,retval,this,deletedrs,deleted_values,tdeleteMessage,lockRecord,lockWhere,keysvalue,lockSQL,lockSet,data,var_SESSION
	this_object.deleteMessage = ""
	if bValue(GetRequestValue(Request,"mdelete")) then
		GetCollectionBounds GetRequestValue(Request,"mdelete"),c_listpage_loopIdx16,c_listpage_loopMax16
		do while c_listpage_loopIdx16<=c_listpage_loopMax16
			c_listpage_arrKey16 = GetCollectionKey(GetRequestValue(Request,"mdelete"),c_listpage_loopIdx16)
			doAssignment ind,GetRequestValue(GetRequestValue(Request,"mdelete"),c_listpage_arrKey16)
			i = 0
			do while IsLess(i,asp_count(this_object.arrKeyFields))
				setArrElement keys,ArrayElement(this_object.arrKeyFields,i),GetRequestValue(GetRequestValue(Request,"mdelete" & CSmartStr(CSmartDbl(i)+1)),mdeleteIndex(ind))
				i = CSmartDbl(i)+1
			loop
			setArrElement this_object.selectedRecs,asp_count(this_object.selectedRecs),keys
			c_listpage_loopIdx16=c_listpage_loopIdx16+1
		loop
	else
		if bValue(GetRequestValue(Request,"selection")) then
			GetCollectionBounds GetRequestValue(Request,"selection"),c_listpage_loopIdx18,c_listpage_loopMax18
			c_listpage_exitLoop18=false
			do while c_listpage_loopIdx18<=c_listpage_loopMax18
				c_listpage_exitLoop18=false
				do
					c_listpage_arrKey18 = GetCollectionKey(GetRequestValue(Request,"selection"),c_listpage_loopIdx18)
					doAssignment keyblock,GetRequestValue(GetRequestValue(Request,"selection"),c_listpage_arrKey18)
					doAssignmentByRef arr,explode("&",keyblock)
					if IsLess(asp_count(arr),asp_count(this_object.arrKeyFields)) then
						exit do
					end if
					i = 0
					do while IsLess(i,asp_count(this_object.arrKeyFields))
						setArrElement keys,ArrayElement(this_object.arrKeyFields,i),asp_urldecode(ArrayElement(arr,i))
						i = CSmartDbl(i)+1
					loop
					setArrElement this_object.selectedRecs,asp_count(this_object.selectedRecs),keys
				loop while false
				if c_listpage_exitLoop18 then _
					exit do
				c_listpage_loopIdx18=c_listpage_loopIdx18+1
			loop
		end if
	end if
	this_object.recordsDeleted = 0
	doClassAssignment this_object,"lockDelRec",CreateDictionary()
	GetCollectionBounds this_object.selectedRecs,c_listpage_loopIdx20,c_listpage_loopMax20
	do while c_listpage_loopIdx20<=c_listpage_loopMax20
		c_listpage_arrKey20 = GetCollectionKey(this_object.selectedRecs,c_listpage_loopIdx20)
		doAssignment keys,ArrayElement(this_object.selectedRecs,c_listpage_arrKey20)
		doAssignmentByRef where,KeyWhere(keys,"")
		if (not IsEqual(this_object.nSecOptions,ADVSECURITY_ALL) and IsEqual(this_object.nLoginMethod,SECURITY_TABLE)) and bValue(this_object.createLoginPage) then
			doAssignmentByRef where,whereAdd(where,SecuritySQL("Delete",""))
		end if
		strSQl = (("delete from " & CSmartStr(AddTableWrappers(this_object.origTName))) & " where ") & CSmartStr(where)
		retval = true
		if (bValue(this_object.eventExists_p1("AfterDelete")) or bValue(this_object.eventExists_p1("BeforeDelete"))) or bValue(this_object.audit) then
			doAssignmentByRef deletedrs,db_query(gSQLWhere_having(this_object.gsqlHead,this_object.gsqlFrom,this_object.gsqlWhereExpr,this_object.gsqlGroupBy,this_object.gsqlHaving,where,""),this_object.conn)
			doAssignmentByRef deleted_values,db_fetch_array(deletedrs)
		end if
		if bValue(this_object.eventExists_p1("BeforeDelete")) then
			doAssignment tdeleteMessage,this_object.deleteMessage
			doAssignmentByRef retval,this_object.eventsObject.BeforeDelete_p3(where,deleted_values,tdeleteMessage)
			doClassAssignment this_object,"deleteMessage",tdeleteMessage
		end if
		lockRecord = false
		if bValue(this_object.lockingObj) then
			lockWhere = ""
			GetCollectionBounds keys,c_listpage_loopIdx21,c_listpage_loopMax21
			do while c_listpage_loopIdx21<=c_listpage_loopMax21
				c_listpage_arrKey21 = GetCollectionKey(keys,c_listpage_loopIdx21)
				doAssignment keysvalue,ArrayElement(keys,c_listpage_arrKey21)
				lockWhere = CSmartStr(lockWhere) & (CSmartStr(asp_rawurlencode(keysvalue)) & "&")
				c_listpage_loopIdx21=c_listpage_loopIdx21+1
			loop
			doAssignmentByRef lockWhere,asp_substr(lockWhere,0,-1)
			lockSQL = ((((((((((("select * from " & CSmartStr(AddTableWrappers(""))) & " where ") & CSmartStr(AddFieldWrappers("keys"))) & "='") & CSmartStr(db_addslashes(lockWhere))) & "' and ") & CSmartStr(AddFieldWrappers("table"))) & "='") & CSmartStr(db_addslashes(this_object.origTName))) & "' and ") & CSmartStr(AddFieldWrappers("action"))) & "=1"
			doAssignmentByRef lockSet,db_query(lockSQL,this_object.conn)
			if bValue(doAssignmentByRef(data,db_fetch_array(lockSet))) then
				lockRecord = true
				setArrElement this_object.lockDelRec,asp_count(this_object.lockDelRec),keys
			end if
			setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_lockDelRec",this_object.lockDelRec
		end if
		if not bValue(lockRecord) then
			if IsEqual(GetRequestValue(Request,"a"),"delete") then
				if bValue(retval) then
					this_object.recordsDeleted = CSmartDbl(this_object.recordsDeleted)+1
					if bValue(this_object.delFile) then
						DeleteUploadedFiles where,""
					end if
					LogInfo strSQl
					db_exec strSQl,this_object.conn
					if bValue(this_object.audit) and bValue(deleted_values) then
						this_object.audit.LogDelete_p3 this_object.tName,deleted_values,keys
					end if
					if bValue(this_object.eventExists_p1("AfterDelete")) then
						doAssignment tdeleteMessage,this_object.deleteMessage
						this_object.eventsObject.AfterDelete_p3 where,deleted_values,tdeleteMessage
						doClassAssignment this_object,"deleteMessage",tdeleteMessage
					end if
				end if
			end if
		end if
		if bValue(asp_strlen(this_object.deleteMessage)) then
			this_object.xt.assignbyref_p2 "user_message",this_object.deleteMessage
		end if
		c_listpage_loopIdx20=c_listpage_loopIdx20+1
	loop
	if bValue(asp_count(this_object.selectedRecs)) then
		if bValue(this_object.eventExists_p1("AfterMassDelete")) then
			this_object.eventsObject.AfterMassDelete_p1 this_object.recordsDeleted
		end if
	end if
End Function
Function method_ListPage_rulePRG(byref this_object)
	Dim this
	if (bValue(no_output_done()) and bValue(asp_count(this_object.selectedRecs))) and not bValue(asp_strlen(this_object.deleteMessage)) then
		asp_header ((("Location: " & CSmartStr(this_object.shortTableName)) & "_") & CSmartStr(this_object.getPageType())) & ".asp?a=return"
		Response.End
	end if
End Function
Function method_ListPage_BeforeShowList(byref this_object)
	Dim this,templatefile
	if bValue(this_object.eventExists_p1("BeforeShowList")) then
		doAssignment templatefile,this_object.templatefile
		this_object.eventsObject.BeforeShowList_p2 this_object.xt,templatefile
		doClassAssignment this_object,"templatefile",templatefile
	end if
End Function
Function method_ListPage_buildSearchPanel(byref this_object)
End Function
Function method_ListPage_assignAdmin(byref this_object)
	Dim this
	if bValue(this_object.isAdminTable()) then
		this_object.xt.assign_p2 "html_attrs","lang=""en"""
		this_object.xt.assign_p2 "exitaalink_attrs","href=""menu.asp"" onclick=""window.location.href='menu.asp';return false;"""
		this_object.xt.assign_p2 "exitaalink_href","href=""menu.asp"""
		this_object.xt.assign_p2 "exitadminarea_link",true
		this_object.xt.assign_p2 "admin_rights_tablelink",true
		this_object.xt.assign_p2 "admin_rights_tablelink_attrs","href=""admin_rights_list.asp"""
		this_object.xt.assign_p2 "admin_members_tablelink",true
		this_object.xt.assign_p2 "admin_members_tablelink_attrs","href=""admin_members_list.asp"""
		this_object.xt.assign_p2 "admin_users_tablelink",true
		this_object.xt.assign_p2 "admin_users_tablelink_attrs","href=""admin_users_list.asp"""
		this_object.xt.assign_p2 "admin_rights_optionattrs","value=""admin_rights_list.asp"""
		this_object.xt.assign_p2 "admin_members_optionattrs","value=""admin_members_list.asp"""
		this_object.xt.assign_p2 "admin_users_optionattrs","value=""admin_users_list.asp"""
	end if
	if bValue(this_object.isDynamicPerm) and bValue(IsAdmin()) then
		this_object.xt.assign_p2 "adminarea_link",true
		this_object.xt.assign_p2 "adminarealink_attrs","href=""admin_rights_list.asp"" onclick=""window.location.href='admin_rights_list.asp';return false;"""
	end if
End Function
Function method_ListPage_commonAssign(byref this_object)
	Dim this
	this_object.xt.assign_p2 "id",this_object.id
	this_object.xt.assignbyref_p2 "body",this_object.body
	this_object.xt.enable_section_p1 "style_block"
	this_object.xt.enable_section_p1 "iestyle_block"
	this_object.xt.assign_p2 "recordcontrols_block",bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"add")) or bValue(this_object.isDispGrid())
	this_object.xt.assign_p2 "newrecord_controls",ArrayElement(ArrayElement(this_object.permis,this_object.tName),"add")
	this_object.xt.assign_p2 "grid_controls",this_object.isDispGrid()
	this_object.xt.enable_section_p1 "usermessage"
	this_object.xt.enable_section_p1 "message_block"
	this_object.importLinksAttrs 
	this_object.xt.assign_p2 "changepwd_link",not IsEqual(Session("AccessLevel"),ACCESS_LEVEL_GUEST)
	this_object.xt.assign_p2 "changepwdlink_attrs","href=""changepwd.asp"" onclick=""window.location.href='changepwd.asp';return false;"""
	if bValue(this_object.isCreateMenu()) or bValue(this_object.isAdminTable()) then
		this_object.xt.assign_p2 "quickjump_attrs","onfocus =""window.selectcurrent = this.selectedIndex;"" " & vbcrlf & _
			"							   onchange=""if(this.options[this.selectedIndex].value){if($(this.options[this.selectedIndex]).attr('link') == 'External'){window.open(this.options[this.selectedIndex].value);}else{window.location.href = this.options[this.selectedIndex].value;}}else{this.selectedIndex = window.selectcurrent;}"""
	end if
	if bValue(this_object.isAddWebRep) then
		this_object.xt.assign_p2 "webreport_link",true
	end if
	if bValue(this_object.createLoginPage) then
		this_object.xt.assign_p2 "security_block",true
		this_object.xt.assign_p2 "username",htmlspecialchars(Session("UserID"))
		this_object.xt.assign_p2 "logoutlink_attrs","onclick=""window.location.href='login.asp?a=logout';return false;"""
	end if
	GetCollectionBounds ArrayElement(this_object.googleMapCfg,"mainMapIds"),c_listpage_loopIdx22,c_listpage_loopMax22
	do while c_listpage_loopIdx22<=c_listpage_loopMax22
		c_listpage_arrKey22 = GetCollectionKey(ArrayElement(this_object.googleMapCfg,"mainMapIds"),c_listpage_loopIdx22)
		doAssignment mapId,ArrayElement(ArrayElement(this_object.googleMapCfg,"mainMapIds"),c_listpage_arrKey22)
		this_object.xt.assign_event_p4 mapId,this_object,"createMapDiv",CreateDictionary3("mapId",mapId,"width",ArrayElement(ArrayElement(ArrayElement(this_object.googleMapCfg,"mapsData"),mapId),"width"),"height",ArrayElement(ArrayElement(ArrayElement(this_object.googleMapCfg,"mapsData"),mapId),"height"))
		c_listpage_loopIdx22=c_listpage_loopIdx22+1
	loop
	this_object.addAssignForGrid 
End Function
Function method_ListPage_addAssignForGrid(byref this_object)
	Dim this,i,colsonpage,record_header,record_footer,rheader,rfooter,gridTableStyle
	if bValue(this_object.is508) then
		this_object.xt.assign_section_p3 "grid_header","<caption style=""display:none"">Table data</caption>",""
	end if
	this_object.xt.assign_p2 "endrecordblock_attrs","colid=""endrecord"""
	this_object.inlineAddLinksAttrs 
	i = 0
	do while IsLess(i,asp_count(this_object.listFields))
		this_object.xt.assign_p2 CSmartStr(GoodFieldName(ArrayElement(ArrayElement(this_object.listFields,i),"fName"))) & "_fieldheadercolumn",true
		this_object.xt.assign_p2 CSmartStr(GoodFieldName(ArrayElement(ArrayElement(this_object.listFields,i),"fName"))) & "_fieldcolumn",true
		this_object.xt.assign_p2 CSmartStr(GoodFieldName(ArrayElement(ArrayElement(this_object.listFields,i),"fName"))) & "_fieldfootercolumn",true
		i = CSmartDbl(i)+1
	loop
	if bValue(this_object.isDispGrid()) then
		doAssignment colsonpage,this_object.recsPerRowList
		Set record_header = (CreateDictionary1("data",CreateDictionary()))
		Set record_footer = (CreateDictionary1("data",CreateDictionary()))
		i = 0
		do while IsLess(i,colsonpage)
			Set rheader = (CreateDictionary())
			Set rfooter = (CreateDictionary())
			if IsLess(i,CSmartDbl(colsonpage)-1) then
				setArrElement rheader,"endrecordheader_block",true
				setArrElement rfooter,"endrecordfooter_block",true
			end if
			setArrElementN record_header,CreateArray2("data",empty),rheader
			setArrElementN record_footer,CreateArray2("data",empty),rfooter
			i = CSmartDbl(i)+1
		loop
		this_object.xt.assignbyref_p2 "record_header",record_header
		this_object.xt.assignbyref_p2 "record_footer",record_footer
		this_object.xt.assign_p2 "grid_header",true
		if not bValue(this_object.numRowsFromSQL) then
			this_object.xt.assign_p2 "gridHeader_attrs",("id=""gridHeaderTr" & CSmartStr(this_object.id)) & """ style=""display: none;"""
		end if
		this_object.xt.assign_p2 "grid_footer",true
		this_object.xt.assign_p2 "record_controls",true
		gridTableStyle = ""
		gridTableStyle = "style="""
		gridTableStyle = CSmartStr(gridTableStyle) & CSmartStr(IIF(IsLess(0,this_object.recordsOnPage),"""","width: 50%;"""))
		this_object.xt.assign_p2 "gridTable_attrs",gridTableStyle
	end if
End Function
Function method_ListPage_createMapDiv(byref this_object,ByRef params)
	ResponseWrite ((((("<div id=""" & CSmartStr(ArrayElement(params,"mapId"))) & """ style=""width: ") & CSmartStr(ArrayElement(params,"width"))) & "px; height: ") & CSmartStr(ArrayElement(params,"height"))) & "px;""></div>"
End Function
Function method_ListPage_importLinksAttrs(byref this_object)
	this_object.xt.assign_p2 "import_link",ArrayElement(ArrayElement(this_object.permis,this_object.tName),"import")
	this_object.xt.assign_p2 "importlink_attrs",((((((("id=""import_" & CSmartStr(this_object.id)) & """ name=""import_") & CSmartStr(this_object.id)) & """ href='") & CSmartStr(this_object.shortTableName)) & "_import.asp' onclick=""window.location.href='") & CSmartStr(this_object.shortTableName)) & "_import.asp';return false;"""
End Function
Function method_ListPage_inlineAddLinksAttrs(byref this_object)
	this_object.xt.assign_p2 "inlineadd_link",ArrayElement(ArrayElement(this_object.permis,this_object.tName),"add")
	this_object.xt.assign_p2 "inlineaddlink_attrs",((((("name=""inlineAdd_" & CSmartStr(this_object.id)) & """ href='") & CSmartStr(this_object.shortTableName)) & "_add.asp' id=""inlineAdd") & CSmartStr(this_object.id)) & """"
End Function
Function method_ListPage_selectAllLinkAttrs(byref this_object)
	this_object.xt.assign_p2 "selectall_link",(bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"delete")) or bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"export"))) or bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"edit"))
	this_object.xt.assign_p2 "selectalllink_span",this_object.buttonShowHideStyle()
	this_object.xt.assign_p2 "selectalllink_attrs",((("name=""select_all" & CSmartStr(this_object.id)) & """ " & vbcrlf & _
		"			id=""select_all") & CSmartStr(this_object.id)) & """ " & vbcrlf & _
		"			href=""#"""
End Function
Function method_ListPage_checkboxColumnAttrs(byref this_object)
	this_object.xt.assign_p2 "checkbox_column",(bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"delete")) or bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"export"))) or bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"edit"))
	this_object.xt.assign_p2 "checkbox_header",true
	this_object.xt.assign_p2 "checkboxheader_attrs",("id=""chooseAll_" & CSmartStr(this_object.id)) & """"
End Function
Function method_ListPage_getPrintExportLinkAttrs(byref this_object,ByVal page)
	if not bValue(page) then
		method_ListPage_getPrintExportLinkAttrs = ""
		Exit Function
	end if
	method_ListPage_getPrintExportLinkAttrs = ((((((((((("name=""" & CSmartStr(page)) & "_selected") & CSmartStr(this_object.id)) & """ " & vbcrlf & _
		"				id=""") & CSmartStr(page)) & "_selected") & CSmartStr(this_object.id)) & """				" & vbcrlf & _
		"				href = '") & CSmartStr(this_object.shortTableName)) & "_") & CSmartStr(page)) & ".asp'"
	Exit Function
End Function
Function method_ListPage_buttonShowHideStyle(byref this_object,ByVal link)
	if IsEqual(link,"saveall") or IsEqual(link,"cancelall") then
		method_ListPage_buttonShowHideStyle = " style=""display:none;"" "
		Exit Function
	else
		doAssignmentByRef method_ListPage_buttonShowHideStyle,IIF(IsLess(0,this_object.numRowsFromSQL),""," style=""display:none;"" ")
		Exit Function
	end if
End Function
Function method_ListPage_editSelectedLinkAttrs(byref this_object)
	this_object.xt.assign_p2 "editselected_link",ArrayElement(ArrayElement(this_object.permis,this_object.tName),"edit")
	this_object.xt.assign_p2 "editselectedlink_span",this_object.buttonShowHideStyle()
	this_object.xt.assign_p2 "editselectedlink_attrs",(((((vbcrlf & _
		"					href='" & CSmartStr(this_object.shortTableName)) & "_edit.asp' " & vbcrlf & _
		"					name=""edit_selected") & CSmartStr(this_object.id)) & """ " & vbcrlf & _
		"					id=""edit_selected") & CSmartStr(this_object.id)) & """"
End Function
Function method_ListPage_saveAllLinkAttrs(byref this_object)
	this_object.xt.assign_p2 "saveall_link",ArrayElement(ArrayElement(this_object.permis,this_object.tName),"edit")
	this_object.xt.assign_p2 "savealllink_span",this_object.buttonShowHideStyle_p1("saveall")
	this_object.xt.assign_p2 "savealllink_attrs",((("name=""saveall_edited" & CSmartStr(this_object.id)) & """ id=""saveall_edited") & CSmartStr(this_object.id)) & """"
End Function
Function method_ListPage_cancelAllLinkAttrs(byref this_object)
	this_object.xt.assign_p2 "cancelall_link",bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"edit")) or bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"edit"))
	this_object.xt.assign_p2 "cancelalllink_span",this_object.buttonShowHideStyle_p1("cancelall")
	this_object.xt.assign_p2 "cancelalllink_attrs",((("name=""revertall_edited" & CSmartStr(this_object.id)) & """ id=""revertall_edited") & CSmartStr(this_object.id)) & """"
End Function
Function method_ListPage_deleteSelectedLink(byref this_object)
	Dim this
	this_object.xt.assign_p2 "deleteselected_link",ArrayElement(ArrayElement(this_object.permis,this_object.tName),"delete")
	this_object.xt.assign_p2 "deleteselectedlink_span",this_object.buttonShowHideStyle()
	this_object.deleteSelectedAttrs 
End Function
Function method_ListPage_deleteSelectedAttrs(byref this_object)
	this_object.xt.assign_p2 "deleteselectedlink_attrs",((("id=""delete_selected" & CSmartStr(this_object.id)) & """ name=""delete_selected") & CSmartStr(this_object.id)) & """ "
End Function
Function method_ListPage_getFormInputsHTML(byref this_object)
	method_ListPage_getFormInputsHTML = ""
	Exit Function
End Function
Function method_ListPage_getFormTargetHTML(byref this_object)
	method_ListPage_getFormTargetHTML = ""
	Exit Function
End Function
Function method_ListPage_buildPagination(byref this_object)
	Dim message,message_block,maxRecords,pagination,counterstart,counterend,this,counter
	this_object.recordsOnPage = CSmartDbl(this_object.numRowsFromSQL)-(CSmartDbl(this_object.myPage)-1)*CSmartDbl(this_object.pageSize)
	if IsLess(this_object.pageSize,this_object.recordsOnPage) then
		doClassAssignment this_object,"recordsOnPage",this_object.pageSize
	end if
	doClassAssignment this_object,"colsOnPage",this_object.recsPerRowList
	if IsLess(this_object.recordsOnPage,this_object.colsOnPage) then
		doClassAssignment this_object,"colsOnPage",this_object.recordsOnPage
	end if
	if IsLess(this_object.colsOnPage,1) then
		this_object.colsOnPage = 1
	end if
	if not bValue(this_object.numRowsFromSQL) then
		this_object.rowsFound = false
		message = "No records found"
		Set message_block = (CreateDictionary())
		setArrElement message_block,"begin",("<span name=""notfound_message" & CSmartStr(this_object.id)) & """>"
		setArrElement message_block,"end","</span>"
		this_object.xt.assignbyref_p2 "message_block",message_block
		this_object.xt.assign_p2 "message",CSmartStr(IIF(IsEqual(this_object.is508,true),"<a name=""skipdata""></a>","")) & CSmartStr(message)
	else
		this_object.rowsFound = true
		doAssignment maxRecords,this_object.numRowsFromSQL
		this_object.xt.assign_p2 "records_found",this_object.numRowsFromSQL
		if bValue(this_object.pageSize) and not IsEqual(this_object.pageSize,-1) then
			doClassAssignment this_object,"maxPages",asp_ceil(CSmartDbl(maxRecords)/CSmartDbl(this_object.pageSize))
		end if
		if IsLess(this_object.maxPages,this_object.myPage) then
			doClassAssignment this_object,"myPage",this_object.maxPages
		end if
		if IsLess(this_object.myPage,1) then
			this_object.myPage = 1
		end if
		doClassAssignment this_object,"maxRecs",this_object.pageSize
		this_object.xt.assign_p2 "page",this_object.myPage
		this_object.xt.assign_p2 "maxpages",this_object.maxPages
		this_object.xt.assign_p2 "pagination_block",true
		if IsLess(1,this_object.maxPages) then
			pagination = ("<table rows='1' cols='1' align='center' width='auto' border='0' name='paginationTable" & CSmartStr(this_object.id)) & "'>"
			pagination = CSmartStr(pagination) & "<tr valign='center'><td align='center'>"
			counterstart = CSmartDbl(this_object.myPage)-9
			if bValue(this_object.myPage mod 10) then
				counterstart = (CSmartDbl(this_object.myPage)-this_object.myPage mod 10)+1
			end if
			counterend = CSmartDbl(counterstart)+9
			if IsLess(this_object.maxPages,counterend) then
				doAssignment counterend,this_object.maxPages
			end if
			if not IsEqual(counterstart,1) then
				pagination = CSmartStr(pagination) & (CSmartStr(this_object.getPaginationLink_p2(1,"First")) & "&nbsp;:&nbsp;")
				pagination = CSmartStr(pagination) & (CSmartStr(this_object.getPaginationLink_p2(CSmartDbl(counterstart)-1,"Previous")) & "&nbsp;")
			end if
			pagination = CSmartStr(pagination) & "<b>[</b>"
			doAssignment counter,counterstart
			do while IsLessOrEqual(counter,counterend)
				if not IsEqual(counter,this_object.myPage) then
					pagination = CSmartStr(pagination) & ("&nbsp;" & CSmartStr(this_object.getPaginationLink_p3(counter,counter,true)))
				else
					pagination = CSmartStr(pagination) & (("&nbsp;<b>" & CSmartStr(counter)) & "</b>")
				end if
				counter = CSmartDbl(counter)+1
			loop
			pagination = CSmartStr(pagination) & "&nbsp;<b>]</b>"
			if not IsEqual(counterend,this_object.maxPages) then
				pagination = CSmartStr(pagination) & (("&nbsp;" & CSmartStr(this_object.getPaginationLink_p2(CSmartDbl(counterend)+1,"Next"))) & "&nbsp;:&nbsp;")
				pagination = CSmartStr(pagination) & ("&nbsp;" & CSmartStr(this_object.getPaginationLink_p2(this_object.maxPages,"Last")))
			end if
			pagination = CSmartStr(pagination) & "</td></tr></table>"
			this_object.xt.assign_p2 "pagination",pagination
		end if
	end if
End Function
Function method_ListPage_getPaginationLink(byref this_object,ByVal pageNum,ByVal linkText,ByVal cls)
	method_ListPage_getPaginationLink = ((((("<a href=""#"" pageNum=""" & CSmartStr(pageNum)) & """ ") & CSmartStr(IIF(cls,"class=""pag_n""",""))) & " style=""TEXT-DECORATION: none;"">") & CSmartStr(linkText)) & "</a>"
	Exit Function
End Function
Function method_ListPage_addWhereWithMasterTable(byref this_object)
	Dim where,i,mValue,var_SESSION
	where = ""
	if bValue(asp_count(this_object.detailKeysByM)) then
		i = 0
		do while IsLess(i,asp_count(this_object.detailKeysByM))
			if not IsEqual(i,0) then
				where = CSmartStr(where) & " and "
			end if
			doAssignmentByRef mValue,make_db_value(ArrayElement(this_object.detailKeysByM,i),Session((CSmartStr(this_object.sessionPrefix) & "_masterkey") & CSmartStr(CSmartDbl(i)+1)),"","","")
			if not (not bValue(mValue)) then
				where = CSmartStr(where) & ((CSmartStr(GetFullFieldName(ArrayElement(this_object.detailKeysByM,i),"")) & "=") & CSmartStr(mValue))
			else
				where = CSmartStr(where) & "1=0"
			end if
			i = CSmartDbl(i)+1
		loop
	end if
	doAssignmentByRef method_ListPage_addWhereWithMasterTable,where
	Exit Function
End Function
Function method_ListPage_seekPageInRecSet(byref this_object,ByVal strSQL)
	Dim listarray,this,maxrecs
	listarray = false
	if bValue(this_object.eventExists_p1("ListQuery")) then
		doAssignmentByRef listarray,this_object.eventsObject.ListQuery_p8(this_object.searchClauseObj,this_object.arrFieldForSort,this_object.arrHowFieldSort,this_object.masterTable,this_object.masterKeysReq,null,this_object.pageSize,this_object.myPage)
	end if
	if not IsFalse(listarray) then
		doClassAssignment this_object,"recSet",listarray
	else
		if IsEqual(this_object.dbType,nDATABASE_MySQL) then
			if IsLess(1,this_object.maxPages) then
				strSQL = CSmartStr(strSQL) & (((" limit " & CSmartStr((CSmartDbl(this_object.myPage)-1)*CSmartDbl(this_object.pageSize))) & ",") & CSmartStr(this_object.pageSize))
			end if
			doClassAssignment this_object,"recSet",db_query(strSQL,this_object.conn)
		else
			if IsEqual(this_object.dbType,nDATABASE_MSSQLServer) then
				if IsLess(1,this_object.maxPages) then
					doAssignmentByRef strSQL,AddTop(strSQL,CSmartDbl(this_object.myPage)*CSmartDbl(this_object.pageSize))
				end if
				doClassAssignment this_object,"recSet",db_query(strSQL,this_object.conn)
				db_pageseek this_object.recSet,this_object.pageSize,this_object.myPage
			else
				if IsEqual(this_object.dbType,nDATABASE_Access) then
					if IsLess(1,this_object.maxPages) then
						doAssignmentByRef strSQL,AddTop(strSQL,CSmartDbl(this_object.myPage)*CSmartDbl(this_object.pageSize))
					end if
					doClassAssignment this_object,"recSet",db_query(strSQL,this_object.conn)
					db_pageseek this_object.recSet,this_object.pageSize,this_object.myPage
				else
					if IsEqual(this_object.dbType,nDATABASE_Oracle) then
						if IsLess(1,this_object.maxPages) then
							doAssignmentByRef strSQL,AddRowNumber(strSQL,CSmartDbl(this_object.myPage)*CSmartDbl(this_object.pageSize))
						end if
						doClassAssignment this_object,"recSet",db_query(strSQL,this_object.conn)
						db_pageseek this_object.recSet,this_object.pageSize,this_object.myPage
					else
						if IsEqual(this_object.dbType,nDATABASE_PostgreSQL) then
							if IsLess(1,this_object.maxPages) then
								doAssignment maxrecs,this_object.pageSize
								strSQL = CSmartStr(strSQL) & (((" limit " & CSmartStr(this_object.pageSize)) & " offset ") & CSmartStr((CSmartDbl(this_object.myPage)-1)*CSmartDbl(this_object.pageSize)))
							end if
							doClassAssignment this_object,"recSet",db_query(strSQL,this_object.conn)
						else
							if IsEqual(this_object.dbType,nDATABASE_DB2) then
								if IsLess(1,this_object.maxPages) then
									strSQL = (((("with DB2_QUERY as (" & CSmartStr(strSQL)) & ") select * from DB2_QUERY where DB2_ROW_NUMBER between ") & CSmartStr((CSmartDbl(this_object.myPage)-1)*CSmartDbl(this_object.pageSize)+1)) & " and ") & CSmartStr(CSmartDbl(this_object.myPage)*CSmartDbl(this_object.pageSize))
								end if
								doClassAssignment this_object,"recSet",db_query(strSQL,this_object.conn)
							else
								if IsEqual(this_object.dbType,nDATABASE_Informix) then
									if IsLess(1,this_object.maxPages) then
										doAssignmentByRef strSQL,AddTopIfx(strSQL,CSmartDbl(this_object.myPage)*CSmartDbl(this_object.pageSize))
									end if
									doClassAssignment this_object,"recSet",db_query(strSQL,this_object.conn)
									db_pageseek this_object.recSet,this_object.pageSize,this_object.myPage
								else
									if IsEqual(this_object.dbType,nDATABASE_SQLite3) then
										if IsLess(1,this_object.maxPages) then
											strSQL = CSmartStr(strSQL) & (((" limit " & CSmartStr((CSmartDbl(this_object.myPage)-1)*CSmartDbl(this_object.pageSize))) & ",") & CSmartStr(this_object.pageSize))
										end if
										doClassAssignment this_object,"recSet",db_query(strSQL,this_object.conn)
									else
										doClassAssignment this_object,"recSet",db_query(strSQL,this_object.conn)
										db_pageseek this_object.recSet,this_object.pageSize,this_object.myPage
									end if
								end if
							end if
						end if
					end if
				end if
			end if
		end if
	end if
End Function
Function method_ListPage_buildSQL(byref this_object)
	Dim searchWhereClause,searchHavingClause,strSecuritySql,var_GET,var_POST,where,this,strSQL,var_SESSION,strSQLbak,tstrWhereClause,tstrOrderBy,rowcount
	doAssignmentByRef searchWhereClause,this_object.searchClauseObj.getWhere_p1(GetListOfFieldsByExprType(false,this_object.tName))
	doAssignmentByRef searchHavingClause,this_object.searchClauseObj.getWhere_p1(GetListOfFieldsByExprType(true,this_object.tName))
	doClassAssignment this_object,"strWhereClause",whereAdd(this_object.strWhereClause,searchWhereClause)
	doClassAssignment this_object,"strHavingClause",whereAdd(this_object.strHavingClause,searchHavingClause)
	doAssignmentByRef strSecuritySql,SecuritySQL("Search",this_object.tName)
	doClassAssignment this_object,"strWhereClause",whereAdd(this_object.strWhereClause,strSecuritySql)
	if (bValue(this_object.noRecordsFirstPage) and not bValue(asp_count(Request.QueryString))) and not bValue(asp_count(RequestForm())) then
		doClassAssignment this_object,"strWhereClause",whereAdd(this_object.strWhereClause,"1=0")
	end if
	doAssignmentByRef where,this_object.addWhereWithMasterTable()
	doClassAssignment this_object,"strWhereClause",whereAdd(this_object.strWhereClause,where)
	if IsEqual(this_object.dbType,nDATABASE_DB2) then
		this_object.gsqlHead = CSmartStr(this_object.gsqlHead) & ", ROW_NUMBER() over () as DB2_ROW_NUMBER  "
	end if
	doAssignmentByRef strSQL,gSQLWhere_having(this_object.gsqlHead,this_object.gsqlFrom,this_object.gsqlWhereExpr,this_object.gsqlGroupBy,this_object.gsqlHaving,this_object.strWhereClause,this_object.strHavingClause)
	strSQL = CSmartStr(strSQL) & (" " & CSmartStr(trim(this_object.strOrderBy)))
	setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_sql",strSQL
	setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_where",this_object.strWhereClause
	setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_having",this_object.strHavingClause
	setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_order",this_object.strOrderBy
	setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_arrFieldForSort",this_object.arrFieldForSort
	setArrElement Session,CSmartStr(this_object.sessionPrefix) & "_arrHowFieldSort",this_object.arrHowFieldSort
	this_object.addMasterDetailSubQuery 
	doAssignment strSQLbak,strSQL
	if bValue(this_object.eventExists_p1("BeforeQueryList")) then
		doAssignment tstrWhereClause,this_object.strWhereClause
		doAssignment tstrOrderBy,this_object.strOrderBy
		this_object.eventsObject.BeforeQueryList_p3 strSQL,tstrWhereClause,tstrOrderBy
		doClassAssignment this_object,"strWhereClause",tstrWhereClause
		doClassAssignment this_object,"strOrderBy",tstrOrderBy
	end if
	if not IsEqual(strSQL,strSQLbak) then
		doClassAssignment this_object,"numRowsFromSQL",GetRowCount(strSQL)
	else
		doAssignmentByRef strSQL,gSQLWhere_having(this_object.gsqlHead,this_object.gsqlFrom,this_object.gsqlWhereExpr,this_object.gsqlGroupBy,this_object.gsqlHaving,this_object.strWhereClause,this_object.strHavingClause)
		strSQL = CSmartStr(strSQL) & (" " & CSmartStr(trim(this_object.strOrderBy)))
		rowcount = false
		if bValue(this_object.eventExists_p1("ListGetRowCount")) then
			doAssignmentByRef rowcount,this_object.eventsObject.ListGetRowCount_p4(this_object.searchClauseObj,this_object.masterTable,this_object.masterKeysReq,null)
		end if
		if not IsFalse(rowcount) then
			doClassAssignment this_object,"numRowsFromSQL",rowcount
		else
			doClassAssignment this_object,"numRowsFromSQL",gSQLRowCount_int(this_object.gsqlHead,this_object.gsqlFrom,this_object.gsqlWhereExpr,this_object.gsqlGroupBy,this_object.gsqlHaving,this_object.strWhereClause,this_object.strHavingClause)
		end if
	end if
	LogInfo strSQL
	doClassAssignment this_object,"querySQL",strSQL
End Function
Function method_ListPage_addMasterDetailSubQuery(byref this_object)
	Dim i,detailsSqlWhere,detailsTableFrom,origTName,dataSourceTName,shortTName,masterWhere,idx,subQ,k,securityClause,countsql
	if (bValue(this_object.subQueriesSupp) and bValue(this_object.subQueriesSupAccess)) and not bValue(this_object.theSameFieldsType) then
		i = 0
		do while IsLess(i,asp_count(this_object.allDetailsTablesArr))
			if bValue(ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dispChildCount")) or bValue(ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"hideChild")) then
				doAssignment detailsSqlWhere,ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"sqlWhere")
				doAssignment detailsTableFrom,ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"sqlFrom")
				doAssignment origTName,ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dOriginalTable")
				doAssignment dataSourceTName,ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dDataSourceTable")
				doAssignment shortTName,ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dShortTable")
				masterWhere = ""
				GetCollectionBounds ArrayElement(this_object.masterKeysByD,i),c_listpage_loopIdx28,c_listpage_loopMax28
				do while c_listpage_loopIdx28<=c_listpage_loopMax28
					idx = GetCollectionKey(ArrayElement(this_object.masterKeysByD,i),c_listpage_loopIdx28)
					doAssignment val,ArrayElement(ArrayElement(this_object.masterKeysByD,i),idx)
					if bValue(masterWhere) then
						masterWhere = CSmartStr(masterWhere) & " and "
					end if
					masterWhere = CSmartStr(masterWhere) & ((((((CSmartStr(AddTableWrappers("subQuery_cnt")) & ".") & CSmartStr(AddFieldWrappers(ArrayElement(ArrayElement(this_object.detailKeysByD,i),idx)))) & "=") & CSmartStr(AddTableWrappers(this_object.origTName))) & ".") & CSmartStr(AddFieldWrappers(ArrayElement(ArrayElement(this_object.masterKeysByD,i),idx))))
					c_listpage_loopIdx28=c_listpage_loopIdx28+1
				loop
				subQ = ""
				GetCollectionBounds ArrayElement(this_object.detailKeysByD,i),c_listpage_loopIdx29,c_listpage_loopMax29
				do while c_listpage_loopIdx29<=c_listpage_loopMax29
					c_listpage_arrKey29 = GetCollectionKey(ArrayElement(this_object.detailKeysByD,i),c_listpage_loopIdx29)
					doAssignment k,ArrayElement(ArrayElement(this_object.detailKeysByD,i),c_listpage_arrKey29)
					if bValue(asp_strlen(subQ)) then
						subQ = CSmartStr(subQ) & ","
					end if
					subQ = CSmartStr(subQ) & CSmartStr(GetFullFieldName(k,dataSourceTName))
					c_listpage_loopIdx29=c_listpage_loopIdx29+1
				loop
				subQ = (("SELECT " & CSmartStr(subQ)) & " ") & CSmartStr(detailsTableFrom)
				doAssignmentByRef securityClause,SecuritySQL("Search",dataSourceTName)
				if bValue(asp_strlen(securityClause)) then
					subQ = CSmartStr(subQ) & (" WHERE " & CSmartStr(whereAdd(detailsSqlWhere,securityClause)))
				else
					if bValue(asp_strlen(detailsSqlWhere)) then
						subQ = CSmartStr(subQ) & (" WHERE " & CSmartStr(whereAdd("",detailsSqlWhere)))
					end if
				end if
				subQ = CSmartStr(subQ) & (" " & CSmartStr(ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"sqlTail")))
				countsql = (((("SELECT count(*) FROM (" & CSmartStr(subQ)) & ") ") & CSmartStr(AddTableWrappers("subQuery_cnt"))) & " WHERE ") & CSmartStr(masterWhere)
				this_object.gsqlHead = CSmartStr(this_object.gsqlHead) & ((((",(" & CSmartStr(countsql)) & ") as ") & CSmartStr(AddFieldWrappers(CSmartStr(dataSourceTName) & "_cnt"))) & " ")
			end if
			i = CSmartDbl(i)+1
		loop
	end if
End Function
Function method_ListPage_fillGridShowInfo(byref this_object,ByRef rowInfoArr)
	Dim editlink,copylink,row,record,this,addedInlineAddContainer,i,dDataSourceTable,dShortTable,rec
	setArrElement rowInfoArr,"data",CreateDictionary()
	editlink = ""
	copylink = ""
	if (bValue(this_object.isUseInlineAdd) or bValue(this_object.showAddInPopup)) and bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"add")) then
		Set row = (CreateDictionary())
		setArrElement row,"rowattrs",("id=""addarea" & CSmartStr(this_object.id)) & """ style=""display: none;"""
		if bValue(this_object.isVerLayout) then
			setArrElement row,"rowattrs",CSmartStr(ArrayElement(row,"rowattrs")) & "vertical=""1"""
		end if
		setArrElement row,"rowspace_attrs",("id=""addarea" & CSmartStr(this_object.id)) & """"
		Set record = (CreateDictionary())
		setArrElement record,"edit_link",true
		setArrElement record,"inlineedit_link",true
		setArrElement record,"view_link",true
		setArrElement record,"copy_link",true
		setArrElement record,"checkbox",true
		setArrElement record,"checkbox",true
		setArrElement record,"editlink_attrs",("id=""editLink_add" & CSmartStr(this_object.id)) & """"
		this_object.countWidthListIcons_p1 "add"
		setArrElement record,"copylink_attrs",((("id=""copyLink_add" & CSmartStr(this_object.id)) & """ name=""copyLink_add") & CSmartStr(this_object.id)) & """"
		setArrElement record,"viewlink_attrs",((("id=""viewLink_add" & CSmartStr(this_object.id)) & """ name=""viewLink_add") & CSmartStr(this_object.id)) & """"
		setArrElement record,"checkbox_attrs",("id=""check_add" & CSmartStr(this_object.id)) & """ name=""selection[]"""
		addedInlineAddContainer = false
		if bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"edit")) and bValue(this_object.isUseInlineEdit) then
			setArrElement record,"inlineeditlink_attrs",("id=""inlineEdit_add" & CSmartStr(this_object.id)) & """ "
			addedInlineAddContainer = true
		end if
		i = 0
		do while IsLess(i,asp_count(this_object.allDetailsTablesArr))
			doAssignment dDataSourceTable,ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dDataSourceTable")
			doAssignment dShortTable,ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dShortTable")
			setArrElement record,CSmartStr(dShortTable) & "_dtable_link",bValue(ArrayElement(ArrayElement(this_object.permis,dDataSourceTable),"add")) or bValue(ArrayElement(ArrayElement(this_object.permis,dDataSourceTable),"search"))
			setArrElement record,CSmartStr(dShortTable) & "_dtablelink_attrs",(((((" href=""" & CSmartStr(dShortTable)) & "_list.asp?"" id=""master_") & CSmartStr(dShortTable)) & "_add") & CSmartStr(this_object.id)) & """"
			if IsEqual(ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"previewOnList"),DP_POPUP) then
				setArrElement record,CSmartStr(dShortTable) & "_dtablelink_attrs",CSmartStr(ArrayElement(record,CSmartStr(dShortTable) & "_dtablelink_attrs")) & ((" onmouseover=""RollDetailsLink.showPopup(this,'" & CSmartStr(dShortTable)) & "_detailspreview.asp'+this.href.substr(this.href.indexOf('?')));"" onmouseout=""RollDetailsLink.hidePopup();""")
			else
				if IsEqual(ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"previewOnList"),DP_INLINE) then
					setArrElement record,CSmartStr(dShortTable) & "_dtablelink_attrs",((((((("id = """ & CSmartStr(dShortTable)) & "_preview") & CSmartStr(this_object.id)) & """" & vbcrlf & _
						"						caption = """) & CSmartStr(GetTableCaption(GoodFieldName(dDataSourceTable)))) & """ " & vbcrlf & _
						"						href = """) & CSmartStr(dShortTable)) & "_list.asp?""" & vbcrlf & _
						"						style = ""display:none;"""
				end if
			end if
			i = CSmartDbl(i)+1
		loop
		i = 0
		do while IsLess(i,asp_count(this_object.listFields))
			setArrElement record,CSmartStr(GoodFieldName(ArrayElement(ArrayElement(this_object.listFields,i),"fName"))) & "_value",((("<span id=""add" & CSmartStr(this_object.id)) & "_") & CSmartStr(GoodFieldName(ArrayElement(ArrayElement(this_object.listFields,i),"fName")))) & """>&nbsp;</span>"
			if not bValue(addedInlineAddContainer) then
				if IsEqual(i,0) and not (bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"edit")) and bValue(this_object.isUseInlineEdit)) then
					setArrElement record,CSmartStr(GoodFieldName(ArrayElement(ArrayElement(this_object.listFields,i),"fName"))) & "_value",CSmartStr(ArrayElement(record,CSmartStr(GoodFieldName(ArrayElement(ArrayElement(this_object.listFields,i),"fName"))) & "_value")) & (("<span id=""inlineEdit_add" & CSmartStr(this_object.id)) & """></span>")
				end if
			end if
			setArrElement record,CSmartStr(GoodFieldName(ArrayElement(ArrayElement(this_object.listFields,i),"fName"))) & "_style",this_object.setAttrAlign_p1(i)
			i = CSmartDbl(i)+1
		loop
		if IsLess(1,this_object.colsOnPage) then
			setArrElement record,"endrecord_block",true
		end if
		setArrElement record,"grid_recordheader",true
		setArrElement record,"grid_vrecord",true
		setArrElement row,"grid_record",CreateDictionary1("data",CreateDictionary())
		setArrElementN row,CreateArray3("grid_record","data",empty),record
		i = 1
		do while IsLess(i,this_object.colsOnPage)
			Set rec = (CreateDictionary())
			if IsLess(i,CSmartDbl(this_object.colsOnPage)-1) then
				setArrElement rec,"endrecord_block",true
			end if
			setArrElementN row,CreateArray3("grid_record","data",empty),rec
			i = CSmartDbl(i)+1
		loop
		setArrElement row,"grid_rowspace",true
		setArrElement row,"grid_recordspace",CreateDictionary1("data",CreateDictionary())
		i = 0
		do while IsLess(i,CSmartDbl(this_object.colsOnPage)*2-1)
			setArrElementN row,CreateArray3("grid_recordspace","data",empty),true
			i = CSmartDbl(i)+1
		loop
		setArrElementN rowInfoArr,CreateArray2("data",empty),row
	end if
End Function
Function method_ListPage_beforeProccessRow(byref this_object)
	Dim this,data
	if bValue(this_object.eventExists_p1("ListFetchArray")) then
		doAssignmentByRef data,this_object.eventsObject.ListFetchArray_p1(this_object.recSet)
	else
		doAssignmentByRef data,db_fetch_array(this_object.recSet)
	end if
	c_listpage_exitLoop34=false
	do while bValue(data)
		c_listpage_exitLoop34=false
		do
			if bValue(this_object.eventExists_p1("BeforeProcessRowList")) then
				if not bValue(this_object.eventsObject.BeforeProcessRowList_p1(data)) then
					if bValue(this_object.eventExists_p1("ListFetchArray")) then
						doAssignmentByRef data,this_object.eventsObject.ListFetchArray_p1(this_object.recSet)
					else
						doAssignmentByRef data,db_fetch_array(this_object.recSet)
					end if
					exit do
				end if
			end if
			doAssignmentByRef method_ListPage_beforeProccessRow,data
			Exit Function
		loop while false
		if c_listpage_exitLoop34 then _
			exit do
	loop
End Function
Function method_ListPage_countWidthListIcons(byref this_object,ByVal row)
	Dim editAble,addAble
	if not bValue(this_object.listIcons) or bValue(this_object.widthListIcons) then
		Exit Function
	end if
	if ((not bValue(this_object.edit) and not bValue(this_object.inlineEdit)) and not bValue(this_object.iCopy)) and not bValue(this_object.iView) then
		Exit Function
	end if
	doAssignment editAble,ArrayElement(ArrayElement(this_object.permis,this_object.tName),"edit")
	doAssignment addAble,ArrayElement(ArrayElement(this_object.permis,this_object.tName),"add")
	if IsEqual(row,"add") then
		editAble = true
		addAble = true
	end if
	if bValue(this_object.inlineEdit) and bValue(editAble) then
		this_object.widthListIcons = CSmartDbl(this_object.widthListIcons)+25
	end if
	if IsEqual(this_object.mode,LIST_SIMPLE) or IsEqual(this_object.mode,LIST_AJAX) then
		if bValue(this_object.edit) and bValue(editAble) then
			this_object.widthListIcons = CSmartDbl(this_object.widthListIcons)+25
		end if
		if bValue(this_object.iCopy) and bValue(addAble) then
			this_object.widthListIcons = CSmartDbl(this_object.widthListIcons)+25
		end if
		if bValue(this_object.iView) then
			this_object.widthListIcons = CSmartDbl(this_object.widthListIcons)+25
		end if
	end if
End Function
Function method_ListPage_assignListIconsColumn(byref this_object,ByVal editPermis,ByVal addPermis,ByVal searchPermis)
	if not bValue(this_object.listIcons) then
		Exit Function
	end if
	if not bValue(this_object.widthListIcons) or (not bValue(editPermis) and not bValue(addPermis)) and not bValue(searchPermis) then
		this_object.xt.assign_p2 "listIcons_column",false
		Exit Function
	end if
	if IsLess(0,this_object.widthListIcons) then
		this_object.xt.assign_p2 "listIcons_column",true
		this_object.xt.assign_p2 "widthListIcons",("width=""" & CSmartStr(this_object.widthListIcons)) & """"
	else
		this_object.xt.assign_p2 "listIcons_column",false
	end if
End Function
Function method_ListPage_fillGridData(byref this_object)
	Dim totals,rowinfo,this,shade,data,lockRecIds,tKeys,row,col,record,gridRowInd,i,var_SESSION,lockDelRec,key,val,keyblock,editlink,copylink,keylink,viewLink
	Set totals = (CreateDictionary())
	Set rowinfo = (CreateDictionary())
	this_object.fillGridShowInfo_p1 rowinfo
	shade = false
	doAssignmentByRef data,this_object.beforeProccessRow()
	Set lockRecIds = (CreateDictionary())
	setArrElement this_object.googleMapCfg,"viewLinkBase",CSmartStr(this_object.shortTableName) & "_view.asp?"
	doAssignmentByRef tKeys,GetTableKeys(this_object.tName)
	setArrElement this_object.controlsMap,"gridRows",CreateDictionary()
	do while bValue(data) and (IsLessOrEqual(this_object.recNo,this_object.pageSize) or IsEqual(this_object.pageSize,-1))
		Set row = (CreateDictionary())
		if not bValue(this_object.isVerLayout) then
			setArrElement row,"rowattrs",""
			if not bValue(shade) then
				setArrElement row,"rowattrs",CSmartStr(ArrayElement(row,"rowattrs")) & "class='shade'"
				shade = true
			else
				shade = false
			end if
		end if
		setArrElement row,"grid_record",CreateDictionary()
		setArrElementN row,CreateArray2("grid_record","data"),CreateDictionary()
		this_object.rowId = CSmartDbl(this_object.rowId)+1
		col = 1
		do while (bValue(data) and (IsLessOrEqual(this_object.recNo,this_object.pageSize) or IsEqual(this_object.pageSize,-1))) and IsLessOrEqual(col,this_object.colsOnPage)
			this_object.countTotals_p2 totals,data
			Set record = (CreateDictionary())
			this_object.genId 
			if (IsEqual(col,1) and IsEqual(this_object.recNo,1)) and bValue(this_object.isResizeColumns) then
				setArrElementN this_object.jsSettings,CreateArray3("tableSettings",this_object.tName,"recIdStart"),this_object.recId
			end if
			setArrElement row,"rowattrs",CSmartStr(ArrayElement(row,"rowattrs")) & ((" id=""gridRow" & CSmartStr(this_object.recId)) & """")
			doAssignmentByRef gridRowInd,asp_count(ArrayElement(this_object.controlsMap,"gridRows"))
			setArrElementN this_object.controlsMap,CreateArray3("gridRows",gridRowInd,"id"),this_object.recId
			setArrElementN this_object.controlsMap,CreateArray3("gridRows",gridRowInd,"rowInd"),gridRowInd
			i = 0
			do while IsLess(i,asp_count(tKeys))
				setArrElementN this_object.controlsMap,CreateArray4("gridRows",gridRowInd,"keys",ArrayElement(tKeys,i)),ArrayElement(data,ArrayElement(tKeys,i))
				i = CSmartDbl(i)+1
			loop
			setArrElement record,"edit_link",bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"edit")) and bValue(CheckSecurity(ArrayElement(data,this_object.mainTableOwnerID),"Edit"))
			setArrElement record,"inlineedit_link",bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"edit")) and bValue(CheckSecurity(ArrayElement(data,this_object.mainTableOwnerID),"Edit"))
			setArrElement record,"view_link",ArrayElement(ArrayElement(this_object.permis,this_object.tName),"search")
			setArrElement record,"copy_link",ArrayElement(ArrayElement(this_object.permis,this_object.tName),"add")
			if IsEqual(col,1) then
				this_object.countWidthListIcons_p1 ""
			end if
			if bValue(this_object.lockingObj) then
				if not bValue(asp_count(this_object.lockDelRec)) and not IsEmpty(Session(CSmartStr(this_object.sessionPrefix) & "_lockDelRec")) then
					doClassAssignment this_object,"lockDelRec",Session(CSmartStr(this_object.sessionPrefix) & "_lockDelRec")
					asp_unsetElement Session,CSmartStr(this_object.sessionPrefix) & "_lockDelRec"
				end if
				i = 0
				do while IsLess(i,asp_count(this_object.lockDelRec))
					lockDelRec = true
					GetCollectionBounds ArrayElement(this_object.lockDelRec,i),c_listpage_loopIdx39,c_listpage_loopMax39
					do while c_listpage_loopIdx39<=c_listpage_loopMax39
						key = GetCollectionKey(ArrayElement(this_object.lockDelRec,i),c_listpage_loopIdx39)
						doAssignment val,ArrayElement(ArrayElement(this_object.lockDelRec,i),key)
						if not IsEqual(ArrayElement(data,key),val) then
							lockDelRec = false
							exit do
						end if
						c_listpage_loopIdx39=c_listpage_loopIdx39+1
					loop
					if bValue(lockDelRec) then
						setArrElement lockRecIds,asp_count(lockRecIds),this_object.recId
						exit do
					end if
					i = CSmartDbl(i)+1
				loop
			end if
			this_object.proccessDetailGridInfo_p2 record,data
			keyblock = ""
			editlink = ""
			copylink = ""
			keylink = ""
			i = 0
			do while IsLess(i,asp_count(tKeys))
				if not IsEqual(i,0) then
					keyblock = CSmartStr(keyblock) & "&"
					editlink = CSmartStr(editlink) & "&"
					copylink = CSmartStr(copylink) & "&"
				end if
				keyblock = CSmartStr(keyblock) & CSmartStr(asp_rawurlencode(ArrayElement(data,ArrayElement(tKeys,i))))
				editlink = CSmartStr(editlink) & ((("editid" & CSmartStr(CSmartDbl(i)+1)) & "=") & CSmartStr(htmlspecialchars(asp_rawurlencode(ArrayElement(data,ArrayElement(tKeys,i))))))
				copylink = CSmartStr(copylink) & ((("copyid" & CSmartStr(CSmartDbl(i)+1)) & "=") & CSmartStr(htmlspecialchars(asp_rawurlencode(ArrayElement(data,ArrayElement(tKeys,i))))))
				keylink = CSmartStr(keylink) & ((("&key" & CSmartStr(CSmartDbl(i)+1)) & "=") & CSmartStr(htmlspecialchars(asp_rawurlencode(ArrayElement(data,ArrayElement(tKeys,i))))))
				i = CSmartDbl(i)+1
			loop
			setArrElement record,"editlink_attrs",((((((("id=""editLink" & CSmartStr(this_object.recId)) & """ name=""editLink") & CSmartStr(this_object.recId)) & """ href='") & CSmartStr(this_object.shortTableName)) & "_edit.asp?") & CSmartStr(editlink)) & "'"
			setArrElement record,"inlineeditlink_attrs",((((((("id=""iEditLink" & CSmartStr(this_object.recId)) & """ name=""iEditLink") & CSmartStr(this_object.recId)) & """ href='") & CSmartStr(this_object.shortTableName)) & "_edit.asp?") & CSmartStr(editlink)) & "'"
			setArrElement record,"copylink_attrs",((((((("id=""copyLink" & CSmartStr(this_object.recId)) & """ name=""copyLink") & CSmartStr(this_object.recId)) & """ href='") & CSmartStr(this_object.shortTableName)) & "_add.asp?") & CSmartStr(copylink)) & "'"
			setArrElement record,"viewlink_attrs",((((((("id=""viewLink" & CSmartStr(this_object.recId)) & """ name=""viewLink") & CSmartStr(this_object.recId)) & """ href='") & CSmartStr(this_object.shortTableName)) & "_view.asp?") & CSmartStr(editlink)) & "'"
			viewLink = (CSmartStr(this_object.shortTableName) & "_view.asp?") & CSmartStr(editlink)
			this_object.fillCheckAttr_p3 record,data,keyblock
			if bValue(ArrayElement(this_object.googleMapCfg,"isUseMainMaps")) then
				this_object.addBigGoogleMapMarkers_p2 data,viewLink
			end if
			i = 0
			do while IsLess(i,asp_count(this_object.listFields))
				if bValue(asp_in_array(i,this_object.gMapFields,false)) then
					this_object.addGoogleMapData_p3 ArrayElement(ArrayElement(this_object.listFields,i),"fName"),data,viewLink
				end if
				setArrElement record,ArrayElement(ArrayElement(this_object.listFields,i),"valueFieldName"),this_object.proccessRecordValue_p3(data,keylink,ArrayElement(this_object.listFields,i))
				setArrElement record,ArrayElement(ArrayElement(this_object.listFields,i),"styleFieldName"),this_object.setAttrAlign_p1(i)
				i = CSmartDbl(i)+1
			loop
			if bValue(this_object.eventExists_p1("BeforeMoveNextList")) then
				this_object.eventsObject.BeforeMoveNextList_p3 data,row,record
			end if
			setArrElement this_object.recIds,asp_count(this_object.recIds),this_object.recId
			this_object.addSpansForGridCells_p2 record,data
			if IsLess(col,this_object.colsOnPage) then
				setArrElement record,"endrecord_block",true
			end if
			setArrElement record,"grid_recordheader",true
			setArrElement record,"grid_vrecord",true
			setArrElementN row,CreateArray3("grid_record","data",empty),record
			doAssignmentByRef data,this_object.beforeProccessRow()
			this_object.recNo = CSmartDbl(this_object.recNo)+1
			col = CSmartDbl(col)+1
		loop
		do while IsLessOrEqual(col,this_object.colsOnPage)
			Set record = (CreateDictionary())
			if IsLess(col,this_object.colsOnPage) then
				setArrElement record,"endrecord_block",true
			end if
			setArrElementN row,CreateArray3("grid_record","data",empty),record
			col = CSmartDbl(col)+1
		loop
		setArrElement row,"grid_rowspace",true
		setArrElement row,"grid_recordspace",CreateDictionary1("data",CreateDictionary())
		i = 0
		do while IsLess(i,CSmartDbl(this_object.colsOnPage)*2-1)
			setArrElementN row,CreateArray3("grid_recordspace","data",empty),true
			i = CSmartDbl(i)+1
		loop
		setArrElementN rowinfo,CreateArray2("data",empty),row
	loop
	if bValue(this_object.lockingObj) and bValue(asp_count(lockRecIds)) then
		setArrElementN this_object.jsSettings,CreateArray3("tableSettings",this_object.tName,"lockRecIds"),lockRecIds
	end if
	if bValue(asp_count(ArrayElement(rowinfo,"data"))) then
		setArrElementN rowinfo,CreateArray3("data",CSmartDbl(asp_count(ArrayElement(rowinfo,"data")))-1,"grid_rowspace"),false
		if bValue(this_object.isVerLayout) and bValue(this_object.is508) then
			setArrElement rowinfo,"begin","<caption style=""display:none"">Table data</caption>"
		end if
		this_object.xt.assignbyref_p2 "grid_row",rowinfo
	end if
	this_object.buildTotals_p1 totals
End Function
Function method_ListPage_setAttrAlign(byref this_object,ByVal i)
	Dim var_type
	doAssignmentByRef var_type,GetFieldType(ArrayElement(ArrayElement(this_object.listFields,i),"fName"),"")
	if IsEqual(ArrayElement(ArrayElement(this_object.listFields,i),"editFormat"),FORMAT_LOOKUP_WIZARD) then
		method_ListPage_setAttrAlign = "align=""left"""
		Exit Function
	else
		if IsEqual(ArrayElement(ArrayElement(this_object.listFields,i),"viewFormat"),FORMAT_CHECKBOX) then
			method_ListPage_setAttrAlign = "align=""center"""
			Exit Function
		else
			if IsEqual(ArrayElement(ArrayElement(this_object.listFields,i),"viewFormat"),FORMAT_NUMBER) or bValue(IsNumberType(var_type)) then
				method_ListPage_setAttrAlign = "align=""right"""
				Exit Function
			else
				if IsEqual(ArrayElement(ArrayElement(this_object.listFields,i),"viewFormat"),FORMAT_AUDIO) then
					method_ListPage_setAttrAlign = "align=""center"" style=""vertical-align:top;"""
					Exit Function
				else
					method_ListPage_setAttrAlign = "align=""left"""
					Exit Function
				end if
			end if
		end if
	end if
End Function
Function method_ListPage_countTotals(byref this_object,ByRef totals,ByRef data)
	Dim i,time
	i = 0
	do while IsLess(i,asp_count(this_object.totalsFields))
		if IsEqual(ArrayElement(ArrayElement(this_object.totalsFields,i),"totalsType"),"COUNT") then
			setArrElement totals,ArrayElement(ArrayElement(this_object.totalsFields,i),"fName"),CSmartDbl(ArrayElement(totals,ArrayElement(ArrayElement(this_object.totalsFields,i),"fName")))+CSmartDbl(not IsEqual(ArrayElement(data,ArrayElement(ArrayElement(this_object.totalsFields,i),"fName")),""))
		else
			if IsEqual(ArrayElement(ArrayElement(this_object.totalsFields,i),"viewFormat"),"Time") then
				doAssignmentByRef time,GetTotalsForTime(ArrayElement(data,ArrayElement(ArrayElement(this_object.totalsFields,i),"fName")))
				setArrElement totals,ArrayElement(ArrayElement(this_object.totalsFields,i),"fName"),CSmartDbl(ArrayElement(totals,ArrayElement(ArrayElement(this_object.totalsFields,i),"fName")))+((CSmartDbl(ArrayElement(time,2))+CSmartDbl(ArrayElement(time,1))*60)+CSmartDbl(ArrayElement(time,0))*3600)
			else
				setArrElement totals,ArrayElement(ArrayElement(this_object.totalsFields,i),"fName"),CSmartDbl(ArrayElement(totals,ArrayElement(ArrayElement(this_object.totalsFields,i),"fName")))+(CSmartDbl(ArrayElement(data,ArrayElement(ArrayElement(this_object.totalsFields,i),"fName")))+0)
			end if
		end if
		i = CSmartDbl(i)+1
	loop
End Function
Function method_ListPage_buildTotals(byref this_object,ByRef totals)
	Dim totals_records,i,record,total
	if bValue(asp_count(this_object.totalsFields)) then
		this_object.xt.assign_p2 "totals_row",true
		Set totals_records = (CreateDictionary1("data",CreateDictionary()))
		i = 0
		do while IsLess(i,this_object.colsOnPage)
			Set record = (CreateDictionary())
			if IsEqual(i,0) then
				i = 0
				do while IsLess(i,asp_count(this_object.totalsFields))
					doAssignmentByRef total,GetTotals(ArrayElement(ArrayElement(this_object.totalsFields,i),"fName"),ArrayElement(totals,ArrayElement(ArrayElement(this_object.totalsFields,i),"fName")),ArrayElement(ArrayElement(this_object.totalsFields,i),"totalsType"),CSmartDbl(this_object.recNo)-1,ArrayElement(ArrayElement(this_object.totalsFields,i),"viewFormat"))
					if bValue(this_object.isUseInlineJs) then
						total = ((((("<span id=""total" & CSmartStr(this_object.id)) & "_") & CSmartStr(GoodFieldName(ArrayElement(ArrayElement(this_object.totalsFields,i),"fName")))) & """>") & CSmartStr(total)) & "</span>"
					end if
					this_object.xt.assign_p2 CSmartStr(GoodFieldName(ArrayElement(ArrayElement(this_object.totalsFields,i),"fName"))) & "_total",total
					setArrElement record,CSmartStr(GoodFieldName(ArrayElement(ArrayElement(this_object.totalsFields,i),"fName"))) & "_showtotal",true
					i = CSmartDbl(i)+1
				loop
			end if
			if IsLess(i,CSmartDbl(this_object.colsOnPage)-1) then
				setArrElement record,"endrecordtotals_block",true
			end if
			setArrElementN totals_records,CreateArray2("data",empty),record
			i = CSmartDbl(i)+1
		loop
		this_object.xt.assignbyref_p2 "totals_record",totals_records
	end if
End Function
Function method_ListPage_outputFieldValue(byref this_object,ByVal field,ByVal state)
	setArrElement this_object.arrFieldSpanVal,field,state
End Function
Function method_ListPage_addSpanVal(byref this_object,ByVal fName,ByRef data)
	if IsEqual(ArrayElement(this_object.arrFieldSpanVal,fName),2) then
		method_ListPage_addSpanVal = ("val=""" & CSmartStr(htmlspecialchars(ArrayElement(data,fName)))) & """ "
		Exit Function
	else
		if IsEqual(ArrayElement(this_object.arrFieldSpanVal,fName),1) then
			method_ListPage_addSpanVal = "val=""1"" "
			Exit Function
		end if
	end if
End Function
Function method_ListPage_addSpansForGridCells(byref this_object,ByRef record,ByRef data)
	Dim i,fName,span,this
	i = 0
	do while IsLess(i,asp_count(this_object.listFields))
		doAssignment fName,ArrayElement(ArrayElement(this_object.listFields,i),"goodFieldName")
		span = "<span "
		span = CSmartStr(span) & (((("id=""edit" & CSmartStr(this_object.recId)) & "_") & CSmartStr(fName)) & """ ")
		span = CSmartStr(span) & CSmartStr(this_object.addSpanVal_p2(ArrayElement(ArrayElement(this_object.listFields,i),"fName"),data))
		span = CSmartStr(span) & ">"
		setArrElement record,ArrayElement(ArrayElement(this_object.listFields,i),"valueFieldName"),(CSmartStr(span) & CSmartStr(ArrayElement(record,ArrayElement(ArrayElement(this_object.listFields,i),"valueFieldName")))) & "</span>"
		i = CSmartDbl(i)+1
	loop
End Function
Function method_ListPage_proccessRecordValue(byref this_object,ByRef data,ByRef keylink,ByVal listFieldInfo)
	Dim value,thumbPref,imgWidth,imgHeight,thumbname,fileNameF,fileName,fieldIsUrl,absFileName,titleField,title,href,videoId,this
	value = ""
	if IsEqual(ArrayElement(listFieldInfo,"viewFormat"),FORMAT_DATABASE_IMAGE) then
		if bValue(ShowThumbnail(ArrayElement(listFieldInfo,"fName"),this_object.tName)) then
			doAssignmentByRef thumbPref,GetThumbnailPrefix(ArrayElement(listFieldInfo,"fName"),this_object.tName)
			value = CSmartStr(value) & "<a"
			if bValue(IsUseiBox(ArrayElement(listFieldInfo,"fName"),this_object.tName)) then
				value = CSmartStr(value) & " rel='ibox'"
			else
				value = CSmartStr(value) & " target=_blank"
			end if
			value = CSmartStr(value) & (((((" href='imager.asp?table=" & CSmartStr(this_object.shortTableName)) & "&field=") & CSmartStr(asp_rawurlencode(ArrayElement(listFieldInfo,"fName")))) & CSmartStr(keylink)) & "'>")
			value = CSmartStr(value) & "<img border=0"
			if bValue(this_object.is508) then
				value = CSmartStr(value) & " alt=""Image from DB"""
			end if
			value = CSmartStr(value) & (((((((" src='imager.asp?table=" & CSmartStr(this_object.shortTableName)) & "&field=") & CSmartStr(asp_rawurlencode(thumbPref))) & "&alt=") & CSmartStr(asp_rawurlencode(ArrayElement(listFieldInfo,"fName")))) & CSmartStr(keylink)) & "'>")
			value = CSmartStr(value) & "</a>"
		else
			value = "<img"
			if bValue(this_object.is508) then
				value = CSmartStr(value) & " alt=""Image from DB"""
			end if
			doAssignmentByRef imgWidth,GetImageWidth(ArrayElement(listFieldInfo,"fName"),this_object.tName)
			value = CSmartStr(value) & CSmartStr(IIF(imgWidth," width=" & CSmartStr(imgWidth),""))
			doAssignmentByRef imgHeight,GetImageHeight(ArrayElement(listFieldInfo,"fName"),this_object.tName)
			value = CSmartStr(value) & CSmartStr(IIF(imgHeight," height=" & CSmartStr(imgHeight),""))
			value = CSmartStr(value) & " border=0"
			value = CSmartStr(value) & (((((" src='imager.asp?table=" & CSmartStr(this_object.shortTableName)) & "&field=") & CSmartStr(asp_rawurlencode(ArrayElement(listFieldInfo,"fName")))) & CSmartStr(keylink)) & "'>")
		end if
	else
		if IsEqual(ArrayElement(listFieldInfo,"viewFormat"),FORMAT_FILE_IMAGE) then
			if bValue(CheckImageExtension(ArrayElement(data,ArrayElement(listFieldInfo,"fName")))) then
				if bValue(ShowThumbnail(ArrayElement(listFieldInfo,"fName"),this_object.tName)) then
					doAssignmentByRef thumbPref,GetThumbnailPrefix(ArrayElement(listFieldInfo,"fName"),this_object.tName)
					thumbname = CSmartStr(thumbPref) & CSmartStr(ArrayElement(data,ArrayElement(listFieldInfo,"fName")))
					if not IsEqual(asp_substr(GetLinkPrefix(ArrayElement(listFieldInfo,"fName"),this_object.tName),0,7),"http://") and not bValue(myfile_exists(CSmartStr(GetUploadFolder(ArrayElement(listFieldInfo,"fName"),"")) & CSmartStr(thumbname))) then
						doAssignment thumbname,ArrayElement(data,ArrayElement(listFieldInfo,"fName"))
					end if
					value = "<a"
					if bValue(IsUseiBox(ArrayElement(listFieldInfo,"fName"),this_object.tName)) then
						value = CSmartStr(value) & " rel='ibox'"
					else
						value = CSmartStr(value) & " target=_blank"
					end if
					value = CSmartStr(value) & ((" href=""" & CSmartStr(htmlspecialchars(AddLinkPrefix(ArrayElement(listFieldInfo,"fName"),ArrayElement(data,ArrayElement(listFieldInfo,"fName")),"")))) & """>")
					value = CSmartStr(value) & "<img"
					if IsEqual(thumbname,ArrayElement(data,ArrayElement(listFieldInfo,"fName"))) then
						doAssignmentByRef imgWidth,GetImageWidth(ArrayElement(listFieldInfo,"fName"),this_object.tName)
						value = CSmartStr(value) & CSmartStr(IIF(imgWidth," width=" & CSmartStr(imgWidth),""))
						doAssignmentByRef imgHeight,GetImageHeight(ArrayElement(listFieldInfo,"fName"),this_object.tName)
						value = CSmartStr(value) & CSmartStr(IIF(imgHeight," height=" & CSmartStr(imgHeight),""))
					end if
					value = CSmartStr(value) & " border=0"
					if bValue(this_object.is508) then
						value = CSmartStr(value) & ((" alt=""" & CSmartStr(htmlspecialchars(ArrayElement(data,ArrayElement(listFieldInfo,"fName"))))) & """")
					end if
					value = CSmartStr(value) & ((" src=""" & CSmartStr(htmlspecialchars(AddLinkPrefix(ArrayElement(listFieldInfo,"fName"),thumbname,"")))) & """></a>")
				else
					value = "<img"
					doAssignmentByRef imgWidth,GetImageWidth(ArrayElement(listFieldInfo,"fName"),this_object.tName)
					value = CSmartStr(value) & CSmartStr(IIF(imgWidth," width=" & CSmartStr(imgWidth),""))
					doAssignmentByRef imgHeight,GetImageHeight(ArrayElement(listFieldInfo,"fName"),this_object.tName)
					value = CSmartStr(value) & CSmartStr(IIF(imgHeight," height=" & CSmartStr(imgHeight),""))
					value = CSmartStr(value) & " border=0"
					if bValue(this_object.is508) then
						value = CSmartStr(value) & ((" alt=""" & CSmartStr(htmlspecialchars(ArrayElement(data,ArrayElement(listFieldInfo,"fName"))))) & """")
					end if
					value = CSmartStr(value) & ((" src=""" & CSmartStr(htmlspecialchars(AddLinkPrefix(ArrayElement(listFieldInfo,"fName"),ArrayElement(data,ArrayElement(listFieldInfo,"fName")),"")))) & """>")
				end if
			end if
		else
			if IsEqual(ArrayElement(listFieldInfo,"viewFormat"),FORMAT_DATABASE_FILE) then
				doAssignmentByRef fileNameF,GetFilenameField(ArrayElement(listFieldInfo,"fName"),this_object.tName)
				if bValue(fileNameF) then
					doAssignment fileName,ArrayElement(data,fileNameF)
					if not bValue(fileName) then
						fileName = "file.bin"
					end if
				else
					fileName = "file.bin"
				end if
				if bValue(asp_strlen(ArrayElement(data,ArrayElement(listFieldInfo,"fName")))) then
					value = (((((("<a href='getfile.asp?table=" & CSmartStr(this_object.shortTableName)) & "&filename=") & CSmartStr(asp_rawurlencode(fileName))) & "&field=") & CSmartStr(asp_rawurlencode(ArrayElement(listFieldInfo,"fName")))) & CSmartStr(keylink)) & "'>"
					value = CSmartStr(value) & CSmartStr(htmlspecialchars(fileName))
					value = CSmartStr(value) & "</a>"
				end if
			else
				if IsEqual(ArrayElement(listFieldInfo,"viewFormat"),FORMAT_AUDIO) then
					doAssignmentByRef fileName,GetData(data,ArrayElement(listFieldInfo,"fName"),FORMAT_NONE)
					doAssignmentByRef fieldIsUrl,GetFieldData(this_object.tName,ArrayElement(listFieldInfo,"fName"),"fieldIsVideoUrl",false)
					if bValue(asp_strlen(fileName)) then
						absFileName = ""
						if not bValue(fieldIsUrl) and bValue(GetFieldData(this_object.tName,ArrayElement(listFieldInfo,"fName"),"Absolute",false)) then
							absFileName = CSmartStr(GetUploadFolder(ArrayElement(listFieldInfo,"fName"),"")) & CSmartStr(fileName)
						else
							if not bValue(fieldIsUrl) then
								doAssignmentByRef absFileName,getabspath(CSmartStr(GetUploadFolder(ArrayElement(listFieldInfo,"fName"),"")) & CSmartStr(fileName))
							end if
						end if
						if bValue(fieldIsUrl) or bValue(asp_file_exists(absFileName)) then
							doAssignmentByRef titleField,GetFieldData(this_object.tName,ArrayElement(listFieldInfo,"fName"),"audioTitleField","")
							title = ""
							if bValue(titleField) then
								doAssignmentByRef title,htmlspecialchars(GetData(data,titleField,ViewFormat(titleField,titleField)))
							end if
							if bValue(fieldIsUrl) then
								doAssignment href,fileName
							else
								href = ((("download.asp?table=" & CSmartStr(this_object.shortTableName)) & "&field=") & CSmartStr(asp_rawurlencode(ArrayElement(listFieldInfo,"fName")))) & CSmartStr(keylink)
							end if
							value = ((((("<a class=""htrack"" type=""audio/mpeg"" title=""" & CSmartStr(title)) & """ href=""") & CSmartStr(href)) & """>") & CSmartStr(title)) & "</a>"
						end if
					end if
				else
					if IsEqual(ArrayElement(listFieldInfo,"viewFormat"),FORMAT_DATABASE_AUDIO) then
						doAssignmentByRef titleField,GetFieldData(this_object.tName,ArrayElement(listFieldInfo,"fName"),"audioTitleField","")
						title = ""
						if bValue(titleField) then
							doAssignmentByRef title,htmlspecialchars(GetData(data,titleField,ViewFormat(titleField,titleField)))
						end if
						if not IsNull(ArrayElement(data,ArrayElement(listFieldInfo,"fName"))) then
							value = (((((((("<a class=""htrack"" type=""audio/mpeg"" title=""" & CSmartStr(title)) & """ href=""getfile.asp?table=") & CSmartStr(this_object.shortTableName)) & "&field=") & CSmartStr(asp_rawurlencode(ArrayElement(listFieldInfo,"fName")))) & CSmartStr(keylink)) & """>") & CSmartStr(title)) & "</a>"
						else
							doAssignment value,title
						end if
					else
						if IsEqual(ArrayElement(listFieldInfo,"viewFormat"),FORMAT_VIDEO) then
							value = ""
							doAssignmentByRef fieldIsUrl,GetFieldData(this_object.tName,ArrayElement(listFieldInfo,"fName"),"fieldIsVideoUrl",false)
							doAssignmentByRef fileName,GetData(data,ArrayElement(listFieldInfo,"fName"),FORMAT_NONE)
							if bValue(asp_strlen(fileName)) then
								absFileName = ""
								if not bValue(fieldIsUrl) and bValue(GetFieldData(this_object.tName,ArrayElement(listFieldInfo,"fName"),"Absolute",false)) then
									absFileName = CSmartStr(GetUploadFolder(ArrayElement(listFieldInfo,"fName"),"")) & CSmartStr(fileName)
								else
									if not bValue(fieldIsUrl) then
										doAssignmentByRef absFileName,getabspath(CSmartStr(GetUploadFolder(ArrayElement(listFieldInfo,"fName"),"")) & CSmartStr(fileName))
									end if
								end if
								if bValue(fieldIsUrl) or bValue(asp_file_exists(absFileName)) then
									videoId = (("video_" & CSmartStr(htmlspecialchars(ArrayElement(listFieldInfo,"fName")))) & "_") & CSmartStr(this_object.recId)
									if bValue(fieldIsUrl) then
										doAssignment href,fileName
									else
										href = ((("download.asp?table=" & CSmartStr(this_object.shortTableName)) & "&field=") & CSmartStr(asp_rawurlencode(ArrayElement(listFieldInfo,"fName")))) & CSmartStr(keylink)
									end if
									value = ((((((("<a href=""" & CSmartStr(href)) & """ style=""display:block;width:") & CSmartStr(GetFieldData(this_object.tName,ArrayElement(listFieldInfo,"fName"),"videoWidth",""))) & "px;height:") & CSmartStr(GetFieldData(this_object.tName,ArrayElement(listFieldInfo,"fName"),"videoHeight",""))) & "px;"" id=""") & CSmartStr(videoId)) & """></a>"
									setArrElementN this_object.controlsMap,CreateArray2("video",empty),videoId
								end if
							end if
						else
							if IsEqual(ArrayElement(listFieldInfo,"viewFormat"),FORMAT_DATABASE_VIDEO) then
								if not IsNull(ArrayElement(data,ArrayElement(listFieldInfo,"fName"))) then
									videoId = (("video_" & CSmartStr(htmlspecialchars(ArrayElement(listFieldInfo,"fName")))) & "_") & CSmartStr(this_object.recId)
									value = (((((((((("<a href=""getfile.asp?table=" & CSmartStr(this_object.shortTableName)) & "&field=") & CSmartStr(asp_rawurlencode(ArrayElement(listFieldInfo,"fName")))) & CSmartStr(keylink)) & """ style=""display:block;width:") & CSmartStr(GetFieldData(this_object.tName,ArrayElement(listFieldInfo,"fName"),"videoWidth",0))) & "px;height:") & CSmartStr(GetFieldData(this_object.tName,ArrayElement(listFieldInfo,"fName"),"videoHeight",0))) & "px;"" id=""") & CSmartStr(videoId)) & """></a>"
									setArrElementN this_object.controlsMap,CreateArray2("video",empty),videoId
								end if
							else
								if IsEqual(ArrayElement(listFieldInfo,"viewFormat"),FORMAT_MAP) then
									value = ((((((("<div id=""littleMap_" & CSmartStr(GoodFieldName(ArrayElement(listFieldInfo,"fName")))) & "_") & CSmartStr(this_object.recId)) & """ style=""width: ") & CSmartStr(ArrayElement(ArrayElement(ArrayElement(this_object.googleMapCfg,"fieldsAsMap"),ArrayElement(listFieldInfo,"fName")),"width"))) & "px; height: ") & CSmartStr(ArrayElement(ArrayElement(ArrayElement(this_object.googleMapCfg,"fieldsAsMap"),ArrayElement(listFieldInfo,"fName")),"height"))) & "px;""></div>"
								else
									if ((IsEqual(ArrayElement(listFieldInfo,"editFormat"),EDIT_FORMAT_LOOKUP_WIZARD) or IsEqual(ArrayElement(listFieldInfo,"editFormat"),EDIT_FORMAT_RADIO)) and IsEqual(GetLookupType(ArrayElement(listFieldInfo,"fName"),this_object.tName),LT_LOOKUPTABLE)) and not IsEqual(GetLWLinkField(ArrayElement(listFieldInfo,"fName"),this_object.tName,true),GetLWDisplayField(ArrayElement(listFieldInfo,"fName"),this_object.tName,true)) then
										doAssignmentByRef value,DisplayLookupWizard(ArrayElement(listFieldInfo,"fName"),ArrayElement(data,ArrayElement(listFieldInfo,"fName")),data,keylink,MODE_LIST)
									else
										if bValue(NeedEncode(ArrayElement(listFieldInfo,"fName"),this_object.tName)) then
											doAssignmentByRef value,ProcessLargeText(GetData(data,ArrayElement(listFieldInfo,"fName"),ArrayElement(listFieldInfo,"viewFormat")),("field=" & CSmartStr(asp_rawurlencode(ArrayElement(listFieldInfo,"fName")))) & CSmartStr(keylink),"",MODE_LIST,"")
										else
											doAssignmentByRef value,GetData(data,ArrayElement(listFieldInfo,"fName"),ArrayElement(listFieldInfo,"viewFormat"))
										end if
									end if
								end if
							end if
						end if
					end if
				end if
			end if
		end if
	end if
	doAssignmentByRef value,this_object.addCenterLink_p2(value,ArrayElement(listFieldInfo,"fName"))
	doAssignmentByRef method_ListPage_proccessRecordValue,value
	Exit Function
End Function
Function method_ListPage_proccessDetailGridInfo(byref this_object,ByRef record,ByRef data)
	Dim i,dDataSourceTable,dShortTable,masterquery,detailid,d,idx,m,this,dpreview
	i = 0
	do while IsLess(i,asp_count(this_object.allDetailsTablesArr))
		doAssignment dDataSourceTable,ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dDataSourceTable")
		doAssignment dShortTable,ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dShortTable")
		masterquery = "mastertable=" & CSmartStr(asp_rawurlencode(this_object.tName))
		Set detailid = (CreateDictionary())
		GetCollectionBounds ArrayElement(this_object.masterKeysByD,i),c_listpage_loopIdx49,c_listpage_loopMax49
		do while c_listpage_loopIdx49<=c_listpage_loopMax49
			idx = GetCollectionKey(ArrayElement(this_object.masterKeysByD,i),c_listpage_loopIdx49)
			doAssignment m,ArrayElement(ArrayElement(this_object.masterKeysByD,i),idx)
			doAssignment d,ArrayElement(ArrayElement(this_object.detailKeysByD,i),idx)
			masterquery = CSmartStr(masterquery) & ((("&masterkey" & CSmartStr(CSmartDbl(idx)+1)) & "=") & CSmartStr(asp_rawurlencode(ArrayElement(data,m))))
			setArrElement detailid,asp_count(detailid),make_db_value(d,ArrayElement(data,m),"","",dDataSourceTable)
			c_listpage_loopIdx49=c_listpage_loopIdx49+1
		loop
		if (bValue(ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dispChildCount")) or bValue(ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"hideChild"))) and ((not bValue(this_object.subQueriesSupp) or not bValue(this_object.subQueriesSupAccess)) or bValue(this_object.theSameFieldsType)) then
			this_object.countDetailsRecsNoSubQ_p3 i,data,detailid
		end if
		dpreview = true
		setArrElement record,CSmartStr(dShortTable) & "_dtable_link",bValue(ArrayElement(ArrayElement(this_object.permis,dDataSourceTable),"add")) or bValue(ArrayElement(ArrayElement(this_object.permis,dDataSourceTable),"search"))
		setArrElement record,CSmartStr(dShortTable) & "_dtablelink_attrs",(((((("href=""" & CSmartStr(dShortTable)) & "_list.asp?") & CSmartStr(masterquery)) & """ id=""master_") & CSmartStr(dShortTable)) & CSmartStr(this_object.recId)) & """"
		if bValue(ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dispChildCount")) then
			if bValue(CSmartDbl(ArrayElement(data,CSmartStr(dDataSourceTable) & "_cnt"))+0) then
				setArrElement record,CSmartStr(dShortTable) & "_childcount",true
			end if
			setArrElement record,CSmartStr(dShortTable) & "_childnumber",ArrayElement(data,CSmartStr(dDataSourceTable) & "_cnt")
			setArrElement record,CSmartStr(dShortTable) & "_childnumber_attr",(((" id='cntDet_" & CSmartStr(dShortTable)) & "_") & CSmartStr(this_object.recId)) & "'"
		end if
		if bValue(dpreview) then
			if IsEqual(GetDPType(dDataSourceTable,this_object.tName),DP_POPUP) then
				setArrElement record,CSmartStr(dShortTable) & "_dtablelink_attrs",CSmartStr(ArrayElement(record,CSmartStr(dShortTable) & "_dtablelink_attrs")) & ((" onmouseover=""RollDetailsLink.showPopup(this,'" & CSmartStr(dShortTable)) & "_detailspreview.asp'+this.href.substr(this.href.indexOf('?')));"" onmouseout=""RollDetailsLink.hidePopup();""")
			end if
			if IsEqual(GetDPType(dDataSourceTable,this_object.tName),DP_INLINE) then
				setArrElement record,CSmartStr(dShortTable) & "_dtablelink_attrs",(((((((((((vbcrlf & _
					"						id = """ & CSmartStr(dShortTable)) & "_preview") & CSmartStr(this_object.recId)) & """ " & vbcrlf & _
					"						caption = """) & CSmartStr(htmlspecialchars(GetTableCaption(GoodFieldName(dDataSourceTable))))) & """") & CSmartStr(IIF(not IsEqual(ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"previewOnList"),DP_NONE),((((("onclick = ""window['dpInline" & CSmartStr(this_object.id)) & "'].showDPInline('") & CSmartStr(htmlspecialchars(jsreplace(dDataSourceTable)))) & "',") & CSmartStr(this_object.recId)) & ",this); return false;""",""))) & "href = """) & CSmartStr(dShortTable)) & "_list.asp?") & CSmartStr(masterquery)) & """"
			end if
		end if
		if bValue(ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"hideChild")) then
			if not bValue(CSmartDbl(ArrayElement(data,CSmartStr(dDataSourceTable) & "_cnt"))+0) then
				setArrElement record,CSmartStr(dShortTable) & "_dtablelink_attrs",CSmartStr(ArrayElement(record,CSmartStr(dShortTable) & "_dtablelink_attrs")) & "style='display:none;'"
			end if
		end if
		i = CSmartDbl(i)+1
	loop
End Function
Function method_ListPage_countDetailsRecsNoSubQ(byref this_object,ByVal i,ByRef data,ByRef detailid)
	Dim dDataSourceTable,gQuery,dObjHaving,dSqlHaving,dSqlGroupBy,dSqlHead,dSqlFrom,dSqlWhere,detailKeys,securityClause,masterwhere,idx
	doAssignment dDataSourceTable,ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dDataSourceTable")
	doAssignmentByRef gQuery,GetTableData(dDataSourceTable,".sqlquery",null)
	doAssignmentByRef dObjHaving,gQuery.Having()
	doAssignmentByRef dSqlHaving,dObjHaving.toSql_p1(gQuery)
	doAssignmentByRef dSqlGroupBy,gQuery.GroupByToSql()
	doAssignment dSqlHead,ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"sqlHead")
	doAssignment dSqlFrom,ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"sqlFrom")
	doAssignment dSqlWhere,ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"sqlWhere")
	doAssignmentByRef detailKeys,GetDetailKeysByMasterTable(this_object.tName,dDataSourceTable,CreateDictionary())
	doAssignmentByRef securityClause,SecuritySQL("Search",dDataSourceTable)
	if bValue(asp_strlen(securityClause)) then
		doAssignmentByRef dSqlWhere,whereAdd(dSqlWhere,securityClause)
	end if
	masterwhere = ""
	GetCollectionBounds ArrayElement(this_object.masterKeysByD,i),c_listpage_loopIdx50,c_listpage_loopMax50
	do while c_listpage_loopIdx50<=c_listpage_loopMax50
		idx = GetCollectionKey(ArrayElement(this_object.masterKeysByD,i),c_listpage_loopIdx50)
		doAssignment val,ArrayElement(ArrayElement(this_object.masterKeysByD,i),idx)
		if bValue(masterwhere) then
			masterwhere = CSmartStr(masterwhere) & " and "
		end if
		masterwhere = CSmartStr(masterwhere) & ((CSmartStr(GetFullFieldName(ArrayElement(detailKeys,idx),dDataSourceTable)) & "=") & CSmartStr(ArrayElement(detailid,idx)))
		c_listpage_loopIdx50=c_listpage_loopIdx50+1
	loop
	setArrElement data,CSmartStr(dDataSourceTable) & "_cnt",gSQLRowCount_int(dSqlHead,dSqlFrom,dSqlWhere,dSqlGroupBy,dSqlHaving,masterwhere,"")
End Function
Function method_ListPage_checkDetailAndMasterFieldTypes(byref this_object)
	Dim i,masterFieldType,idx,detailsFieldType
	if IsEqual(this_object.dbType,nDATABASE_MySQL) then
		method_ListPage_checkDetailAndMasterFieldTypes = false
		Exit Function
	else
		i = 0
		do while IsLess(i,asp_count(this_object.allDetailsTablesArr))
			GetCollectionBounds ArrayElement(this_object.masterKeysByD,i),c_listpage_loopIdx52,c_listpage_loopMax52
			do while c_listpage_loopIdx52<=c_listpage_loopMax52
				idx = GetCollectionKey(ArrayElement(this_object.masterKeysByD,i),c_listpage_loopIdx52)
				doAssignment val,ArrayElement(ArrayElement(this_object.masterKeysByD,i),idx)
				doAssignmentByRef masterFieldType,GetFieldType(ArrayElement(ArrayElement(this_object.masterKeysByD,i),idx),"")
				doAssignmentByRef detailsFieldType,GetFieldType(ArrayElement(ArrayElement(this_object.detailKeysByD,i),idx),ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dDataSourceTable"))
				if not IsEqual(masterFieldType,detailsFieldType) then
					method_ListPage_checkDetailAndMasterFieldTypes = true
					Exit Function
				end if
				c_listpage_loopIdx52=c_listpage_loopIdx52+1
			loop
			i = CSmartDbl(i)+1
		loop
		method_ListPage_checkDetailAndMasterFieldTypes = false
		Exit Function
	end if
End Function
Function method_ListPage_isDispGrid(byref this_object)
	if not bValue(this_object.isUseInlineAdd) and not bValue(this_object.showAddInPopup) then
		method_ListPage_isDispGrid = bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"search")) and bValue(this_object.rowsFound)
		Exit Function
	else
		method_ListPage_isDispGrid = bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"add")) or bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"search")) and bValue(this_object.rowsFound)
		Exit Function
	end if
End Function
Function method_ListPage_fillCheckAttr(byref this_object,ByRef record,ByVal data,ByVal keyblock)
	setArrElement record,"checkbox",ArrayElement(ArrayElement(this_object.permis,this_object.tName),"Edit")
	if (bValue(this_object.exportTo) or bValue(this_object.printFriendly)) or bValue(this_object.deleteRecs) then
		if bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"export")) or bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"delete")) then
			setArrElement record,"checkbox",true
		end if
	end if
	setArrElement record,"checkbox_attrs",((((("name=""selection[]"" value=""" & CSmartStr(keyblock)) & """ id=""check") & CSmartStr(this_object.id)) & "_") & CSmartStr(this_object.recId)) & """"
End Function
Function method_ListPage_callJSCodeAfterRecordEdited(byref this_object)
	method_ListPage_callJSCodeAfterRecordEdited = true
	Exit Function
End Function
Function method_ListPage_addDivSearchWin(byref this_object)
	method_ListPage_addDivSearchWin = ("<div id=""searchWin" & CSmartStr(this_object.id)) & """ class=""searchWin""></div>"
	Exit Function
End Function
Function method_ListPage_prepareForBuildPage(byref this_object)
	Dim this
	this_object.buildOrderParams 
	this_object.deleteRecords 
	this_object.rulePRG 
	this_object.buildSQL 
	this_object.buildPagination 
	this_object.seekPageInRecSet_p1 this_object.querySQL
	this_object.setGoogleMapsParams_p1 this_object.listFields
	this_object.fillGridData 
	if bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"search")) then
		this_object.buildSearchPanel_p1 "adv_search_panel"
	end if
	this_object.addCommonJs 
	this_object.addCommonHtml 
	this_object.commonAssign 
	this_object.assignAdmin 
	this_object.createOldMenu 
End Function
Function method_ListPage_showPage(byref this_object)
	Dim this
	this_object.BeforeShowList 
	this_object.xt.display_p1 this_object.templatefile
End Function
Function method_ListPage_createListPage(byref this_object,ByVal table,ByVal options)
	Dim gQuery,params,oHaving,allfields,f,pageObject
	doAssignmentByRef gQuery,GetTableData(table,".sqlquery",null)
	Set params = (CreateDictionary())
	doAssignment params,options
	setArrElement params,"origTName",GetTableData(table,".OriginalTable","")
	setArrElement params,"sessionPrefix",strTableName
	setArrElement params,"tName",table
	setArrElementByRef params,"conn",conn
	setArrElement params,"gPageSize",GetTableData(table,".pageSize",0)
	setArrElement params,"gOrderIndexes",GetTableData(table,".orderindexes",CreateDictionary())
	setArrElement params,"gstrOrderBy",GetTableData(table,".strOrderBy","")
	setArrElement params,"gsqlHead",GetTableData(table,".sqlHead","")
	setArrElement params,"gsqlFrom",GetTableData(table,".sqlFrom","")
	setArrElement params,"gsqlWhereExpr",GetTableData(table,".sqlWhereExpr","")
	setArrElement params,"gsqlGroupBy",gQuery.GroupByToSql()
	doAssignmentByRef oHaving,gQuery.Having()
	setArrElement params,"gsqlHaving",oHaving.toSql_p1(gQuery)
	setArrElementByRef params,"locale_info",locale_info
	setArrElement params,"subQueriesSupp",bSubqueriesSupported
	setArrElement params,"shortTableName",GetTableData(table,".shortTableName","")
	setArrElement params,"nSecOptions",GetTableData(table,".nSecOptions",0)
	setArrElement params,"nLoginMethod",GetGlobalData("nLoginMethod",0)
	setArrElement params,"recsPerRowList",GetTableData(table,".recsPerRowList",0)
	setArrElement params,"tableGroupBy",GetTableData(table,".tableGroupBy","")
	setArrElement params,"dbType",GetGlobalData("dbType",0)
	setArrElement params,"mainTableOwnerID",GetTableData(table,".mainTableOwnerID","")
	setArrElement params,"moveNext",GetTableData(table,".moveNext",0)
	setArrElement params,"exportTo",GetTableData(table,".exportTo",false)
	setArrElement params,"printFriendly",GetTableData(table,".printFriendly",false)
	setArrElement params,"deleteRecs",GetTableData(table,".delete",false)
	setArrElement params,"rowHighlite",GetTableData(table,".rowHighlite",false)
	setArrElement params,"delFile",GetGlobalData("delFile",false)
	setArrElement params,"isGroupSecurity",isGroupSecurity
	setArrElement params,"arrKeyFields",GetTableData(table,".arrKeyFields",CreateDictionary())
	setArrElement params,"useIbox",GetTableData(table,".useIbox",false)
	setArrElement params,"isUseInlineAdd",GetTableData(table,".isUseInlineAdd",false)
	setArrElement params,"isUseInlineEdit",GetTableData(table,".isUseInlineEdit",false)
	setArrElement params,"isUseInlineJs",bValue(ArrayElement(params,"isUseInlineAdd")) or bValue(ArrayElement(params,"isUseInlineEdit"))
	setArrElement params,"globSearchFields",GetTableData(table,".globSearchFields",CreateDictionary())
	setArrElement params,"panelSearchFields",GetTableData(table,".panelSearchFields",CreateDictionary())
	setArrElement params,"isDynamicPerm",GetGlobalData("isDynamicPerm",false)
	setArrElement params,"isAddWebRep",GetGlobalData("isAddWebRep",false)
	setArrElement params,"isVerLayout",GetTableData(table,".isVerLayout",false)
	setArrElement params,"isDisplayLoading",GetTableData(table,".isDisplayLoading",false)
	setArrElement params,"createLoginPage",GetGlobalData("createLoginPage",false)
	setArrElement params,"menuTablesArr",menuTablesArr
	setArrElement params,"subQueriesSupAccess",GetTableData(table,".subQueriesSupAccess",false)
	setArrElement params,"noRecordsFirstPage",GetTableData(table,".noRecordsFirstPage",false)
	setArrElement params,"totalsFields",GetTableData(table,".totalsFields",CreateDictionary())
	setArrElement params,"listIcons",GetTableData(table,".listIcons",false)
	setArrElement params,"edit",GetTableData(table,".edit",false)
	setArrElement params,"inlineEdit",GetTableData(table,".inlineEdit",false)
	setArrElement params,"iCopy",GetTableData(table,".copy",false)
	setArrElement params,"iView",GetTableData(table,".view",false)
	setArrElement params,"listAjax",GetTableData(table,".listAjax",false)
	setArrElement params,"arrRecsPerPage",GetTableData(table,".arrRecsPerPage",CreateDictionary())
	setArrElement params,"audit",GetAuditObject(table)
	setArrElement params,"listFields",CreateDictionary()
	doAssignmentByRef allfields,GetFieldsList(table)
	GetCollectionBounds allfields,c_listpage_loopIdx53,c_listpage_loopMax53
	c_listpage_exitLoop53=false
	do while c_listpage_loopIdx53<=c_listpage_loopMax53
		c_listpage_exitLoop53=false
		do
			c_listpage_arrKey53 = GetCollectionKey(allfields,c_listpage_loopIdx53)
			doAssignment f,ArrayElement(allfields,c_listpage_arrKey53)
			if not bValue(AppearOnListPage(f,table)) then
				exit do
			end if
			setArrElementN params,CreateArray2("listFields",empty),CreateDictionary6("fName",f,"goodFieldName",GoodFieldName(f),"valueFieldName",CSmartStr(GoodFieldName(f)) & "_value","styleFieldName",CSmartStr(GoodFieldName(f)) & "_style","viewFormat",GetFieldData(table,f,"ViewFormat",""),"editFormat",GetFieldData(table,f,"EditFormat",""))
		loop while false
		if c_listpage_exitLoop53 then _
			exit do
		c_listpage_loopIdx53=c_listpage_loopIdx53+1
	loop
	if IsEqual(ArrayElement(params,"mode"),LIST_SIMPLE) then
		Set pageObject = (CreateClass("ListPage_Simple",1,params,Empty,Empty,Empty,Empty,Empty,Empty))
	else
		if IsEqual(ArrayElement(params,"mode"),LIST_AJAX) then
			Set pageObject = (CreateClass("ListPage_Ajax",1,params,Empty,Empty,Empty,Empty,Empty,Empty))
		else
			if IsEqual(ArrayElement(params,"mode"),LIST_LOOKUP) then
				Set pageObject = (CreateClass("ListPage_Lookup",1,params,Empty,Empty,Empty,Empty,Empty,Empty))
			else
				if IsEqual(ArrayElement(params,"mode"),LIST_DETAILS) then
					Set pageObject = (CreateClass("ListPage_DPInline",1,params,Empty,Empty,Empty,Empty,Empty,Empty))
				else
					if IsEqual(ArrayElement(params,"mode"),RIGHTS_PAGE) then
						Set pageObject = (CreateClass("RightsPage",1,params,Empty,Empty,Empty,Empty,Empty,Empty))
					else
						if IsEqual(ArrayElement(params,"mode"),MEMBERS_PAGE) then
							Set pageObject = (CreateClass("MembersPage",1,params,Empty,Empty,Empty,Empty,Empty,Empty))
						end if
					end if
				end if
			end if
		end if
	end if
	doAssignmentByRef method_ListPage_createListPage,pageObject
	Exit Function
End Function
Function method_ListPage_isAdminTable(byref this_object)
	method_ListPage_isAdminTable = (IsIdentical(this_object.tName,"admin_rights") or IsIdentical(this_object.tName,"admin_members")) or IsIdentical(this_object.tName,"admin_users")
	Exit Function
End Function
%>
