<%

'------ Class ListPage_Simple extends ListPage ------
Class ListPage_Simple
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
	Public Function init_ListPage_Simple_p1(ByRef params)
		DoAssignmentByRef init_ListPage_Simple_p1,method_ListPage_Simple_ListPage_Simple(me,params)
	End Function
	Public Function commonAssign()
		DoAssignmentByRef commonAssign,method_ListPage_Simple_commonAssign(me)
	End Function
	Public Function addAssignForGrid()
		DoAssignmentByRef addAssignForGrid,method_ListPage_Simple_addAssignForGrid(me)
	End Function
	Public Function addCommonHtml()
		DoAssignmentByRef addCommonHtml,method_ListPage_Simple_addCommonHtml(me)
	End Function
	Public Function addJsForGrid()
		DoAssignmentByRef addJsForGrid,method_ListPage_Simple_addJsForGrid(me)
	End Function
	Public Function buildSearchPanel_p1(ByVal xtVarName)
		DoAssignmentByRef buildSearchPanel_p1,method_ListPage_Simple_buildSearchPanel(me,xtVarName)
	End Function
	Public Function displayMasterTableInfo()
		DoAssignmentByRef displayMasterTableInfo,method_ListPage_displayMasterTableInfo(me)
	End Function
	Public Function addCommonJs()
		DoAssignmentByRef addCommonJs,method_ListPage_addCommonJs(me)
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
	Public Function assignAdmin()
		DoAssignmentByRef assignAdmin,method_ListPage_assignAdmin(me)
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
' ListPage_Simple implementation 
Function method_ListPage_Simple_ListPage_Simple(byref this_object,ByRef params)
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
	method_ListPage_ListPage this_object,params
End Function
Function method_ListPage_Simple_commonAssign(byref this_object)
	Dim searchPermis,exportPermis,this
	method_ListPage_commonAssign this_object
	doAssignment searchPermis,ArrayElement(ArrayElement(this_object.permis,this_object.tName),"search")
	doAssignment exportPermis,ArrayElement(ArrayElement(this_object.permis,this_object.tName),"export")
	this_object.xt.assign_p2 "details_block",bValue(searchPermis) and bValue(this_object.rowsFound)
	this_object.xt.assign_p2 "details_attrs",((("id=""detFound" & CSmartStr(this_object.id)) & """ name=""detFound") & CSmartStr(this_object.id)) & """"
	if (bValue(searchPermis) and bValue(asp_count(this_object.arrRecsPerPage))) and (bValue(this_object.rowsFound) or IsEqual(this_object.listAjax,LIST_AJAX)) then
		this_object.xt.assign_p2 "recordspp_block",true
		this_object.createPerPage 
	end if
	this_object.xt.assign_p2 "pages_block",bValue(searchPermis) and bValue(this_object.rowsFound)
	this_object.xt.assign_p2 "pages_attrs",((("id=""pageOf" & CSmartStr(this_object.id)) & """ name=""pageOf") & CSmartStr(this_object.id)) & """"
	this_object.xt.assign_p2 "shiftstyle_block",true
	this_object.xt.assign_p2 "left_block",true
	this_object.xt.assign_p2 "toplinks_block",true
	this_object.xt.assign_p2 "exportselected_link",exportPermis
	this_object.xt.assign_p2 "exportselectedlink_span",this_object.buttonShowHideStyle()
	this_object.xt.assign_p2 "exportselectedlink_attrs",this_object.getPrintExportLinkAttrs_p1("export")
	this_object.xt.assign_p2 "print_link",exportPermis
	this_object.xt.assign_p2 "printlink_attrs",((((((("id='print_" & CSmartStr(this_object.id)) & "'" & vbcrlf & _
		"						   name='print_") & CSmartStr(this_object.id)) & "'" & vbcrlf & _
		"						   href='") & CSmartStr(this_object.shortTableName)) & "_print.asp' " & vbcrlf & _
		"						   onclick=""window.open('") & CSmartStr(this_object.shortTableName)) & "_print.asp','wPrint');return false;"""
	this_object.xt.assign_p2 "printselected_link",exportPermis
	this_object.xt.assign_p2 "printselectedlink_attrs",this_object.getPrintExportLinkAttrs_p1("print")
	this_object.xt.assign_p2 "printselectedlink_span",this_object.buttonShowHideStyle()
	this_object.xt.assign_p2 "printall_link",exportPermis
	this_object.xt.assign_p2 "printalllink_attrs",((((((("id='printAll_" & CSmartStr(this_object.id)) & "'" & vbcrlf & _
		"						   name='printAll_") & CSmartStr(this_object.id)) & "'" & vbcrlf & _
		"						   href='") & CSmartStr(this_object.shortTableName)) & "_print.asp?all=1' " & vbcrlf & _
		"						   onclick=""window.open('") & CSmartStr(this_object.shortTableName)) & "_print.asp?all=1','wPrint');return false;"""
	this_object.xt.assign_p2 "export_link",exportPermis
	this_object.xt.assign_p2 "exportlink_attrs",((((((("id = 'export_" & CSmartStr(this_object.id)) & "'" & vbcrlf & _
		"						   name = 'export_") & CSmartStr(this_object.id)) & "'" & vbcrlf & _
		"						   href='") & CSmartStr(this_object.shortTableName)) & "_export.asp' " & vbcrlf & _
		"						   onclick=""window.open('") & CSmartStr(this_object.shortTableName)) & "_export.asp','wExport');return false;"""
	this_object.xt.assign_p2 "add_link",ArrayElement(ArrayElement(this_object.permis,this_object.tName),"add")
	this_object.xt.assign_p2 "addlink_attrs",((("href='" & CSmartStr(this_object.shortTableName)) & "_add.asp' id=""addButton") & CSmartStr(this_object.id)) & """"
	this_object.selectAllLinkAttrs 
	this_object.editSelectedLinkAttrs 
	this_object.saveAllLinkAttrs 
	this_object.cancelAllLinkAttrs 
	this_object.xt.assign_p2 "grid_block",true
	this_object.xt.assign_p2 "menu_block",this_object.isCreateMenu()
	this_object.xt.assign_p2 "languages_block",true
End Function
Function method_ListPage_Simple_addAssignForGrid(byref this_object)
	Dim editPermis,addPermis,searchPermis,this,i,permis
	method_ListPage_addAssignForGrid this_object
	doAssignment editPermis,ArrayElement(ArrayElement(this_object.permis,this_object.tName),"edit")
	doAssignment addPermis,ArrayElement(ArrayElement(this_object.permis,this_object.tName),"add")
	doAssignment searchPermis,ArrayElement(ArrayElement(this_object.permis,this_object.tName),"search")
	this_object.checkboxColumnAttrs 
	this_object.xt.assign_p2 "edit_column",editPermis
	this_object.xt.assign_p2 "edit_headercolumn",editPermis
	this_object.xt.assign_p2 "edit_footercolumn",editPermis
	this_object.xt.assign_p2 "inlineedit_column",editPermis
	this_object.xt.assign_p2 "inlineedit_headercolumn",editPermis
	this_object.xt.assign_p2 "inlineedit_footercolumn",editPermis
	this_object.xt.assign_p2 "copy_column",addPermis
	this_object.xt.assign_p2 "view_column",searchPermis
	this_object.assignListIconsColumn_p3 editPermis,addPermis,searchPermis
	i = 0
	do while IsLess(i,asp_count(this_object.allDetailsTablesArr))
		permis = bValue(ArrayElement(ArrayElement(this_object.permis,ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dDataSourceTable")),"add")) or bValue(ArrayElement(ArrayElement(this_object.permis,ArrayElement(ArrayElement(this_object.allDetailsTablesArr,i),"dDataSourceTable")),"search"))
		if bValue(permis) then
			this_object.xt.assign_p2 CSmartStr(GoodFieldName(this_object.tName)) & "_dtable_column",permis
			exit do
		end if
		i = CSmartDbl(i)+1
	loop
	this_object.deleteSelectedLink 
End Function
Function method_ListPage_Simple_addCommonHtml(byref this_object)
	Dim this
	setArrElement this_object.body,"begin",CSmartStr(ArrayElement(this_object.body,"begin")) & "<script type=""text/javascript"" src=""include/jquery.js""></script>"
	if IsIdentical(this_object.debugJSMode,true) then
		setArrElement this_object.body,"begin",CSmartStr(ArrayElement(this_object.body,"begin")) & "<script language=""JavaScript"" src=""include/runnerJS/RunnerBase.js""></script>" & vbcrlf
	else
		setArrElement this_object.body,"begin",CSmartStr(ArrayElement(this_object.body,"begin")) & "<script type=""text/javascript"" src=""include/runnerJS/RunnerBase.js""></script>"
	end if
	setArrElement this_object.body,"begin",CSmartStr(ArrayElement(this_object.body,"begin")) & "<script type=""text/javascript"" src=""include/jsfunctions.js""></script>" & vbcrlf
	if bValue(this_object.isDisplayLoading) then
		setArrElement this_object.body,"begin",CSmartStr(ArrayElement(this_object.body,"begin")) & (("<script type=""text/javascript"">runLoading(" & CSmartStr(this_object.id)) & ",document.body,0);</script>")
	end if
	this_object.AddJSFile_p1 "include/customlabels"
	method_ListPage_addCommonHtml this_object
	setArrElement this_object.body,"end",CreateDictionary()
	setArrElementN this_object.body,CreateArray2("end","method"),"assignBodyEnd"
	setArrElementByRefN this_object.body,CreateArray2("end","object"),this_object
End Function
Function method_ListPage_Simple_addJsForGrid(byref this_object)
	Dim this
	method_ListPage_addJsForGrid this_object
	if bValue(this_object.is508) then
		this_object.AddJSCode_p1 (vblf & _
			"s508jumpto(" & CSmartStr(this_object.id)) & ");"
		if bValue(this_object.isUseInlineEdit) and bValue(ArrayElement(ArrayElement(this_object.permis,this_object.tName),"edit")) then
			this_object.AddJSCode_p1 (vblf & _
				"s508inlineEdit(" & CSmartStr(this_object.id)) & ");"
		end if
	end if
End Function
Function method_ListPage_Simple_buildSearchPanel(byref this_object,ByVal xtVarName)
	Dim params,this,panelSearchFields,allSearchFields,i
	Set params = (CreateDictionary())
	setArrElementByRef params,"pageObj",this_object
	setArrElement params,"globSearchFields",this_object.globSearchFields
	setArrElement params,"panelSearchFields",this_object.panelSearchFields
	Set panelSearchFields = (CreateDictionary())
	doAssignmentByRef allSearchFields,GetTableData(this_object.tName,".allSearchFields",CreateDictionary())
	i = 0
	do while IsLess(i,asp_count(allSearchFields))
		if not bValue(this_object.matchWithDetailKeys_p1(ArrayElement(allSearchFields,i))) then
			setArrElement panelSearchFields,asp_count(panelSearchFields),ArrayElement(allSearchFields,i)
		end if
		i = CSmartDbl(i)+1
	loop
	setArrElement params,"allSearchFields",panelSearchFields
	doClassAssignment this_object,"searchPanel",CreateClass("SearchPanelSimple",1,params,Empty,Empty,Empty,Empty,Empty,Empty)
	this_object.searchPanel.buildSearchPanel_p1 xtVarName
End Function
%>
