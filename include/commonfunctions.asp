<%
Function getLangFileName(ByVal langName)
	Dim langArr
	Set langArr = (CreateDictionary())
	setArrElement langArr,"English","English"
	doAssignmentByRef getLangFileName,ArrayElement(langArr,langName)
	Exit Function
End Function
Function GetGlobalData(ByVal name,ByVal defValue)
	if not bValue(asp_array_key_exists(name,globalSettings)) then
		doAssignmentByRef GetGlobalData,defValue
		Exit Function
	end if
	doAssignmentByRef GetGlobalData,ArrayElement(globalSettings,name)
	Exit Function
End Function
Function DisplayMap(ByVal params)
	setArrElementN pageObject.googleMapCfg,CreateArray3("mapsData",ArrayElement(params,"id"),"addressField"),IIF(ArrayElement(params,"addressField"),ArrayElement(params,"addressField"),"")
	setArrElementN pageObject.googleMapCfg,CreateArray3("mapsData",ArrayElement(params,"id"),"latField"),IIF(ArrayElement(params,"latField"),ArrayElement(params,"latField"),"")
	setArrElementN pageObject.googleMapCfg,CreateArray3("mapsData",ArrayElement(params,"id"),"lngField"),IIF(ArrayElement(params,"lngField"),ArrayElement(params,"lngField"),"")
	setArrElementN pageObject.googleMapCfg,CreateArray3("mapsData",ArrayElement(params,"id"),"width"),IIF(ArrayElement(params,"width"),ArrayElement(params,"width"),0)
	setArrElementN pageObject.googleMapCfg,CreateArray3("mapsData",ArrayElement(params,"id"),"height"),IIF(ArrayElement(params,"height"),ArrayElement(params,"height"),0)
	setArrElementN pageObject.googleMapCfg,CreateArray3("mapsData",ArrayElement(params,"id"),"type"),"BIG_MAP"
	setArrElementN pageObject.googleMapCfg,CreateArray3("mapsData",ArrayElement(params,"id"),"showCenterLink"),IIF(ArrayElement(params,"showCenterLink"),ArrayElement(params,"showCenterLink"),0)
	setArrElementN pageObject.googleMapCfg,CreateArray3("mapsData",ArrayElement(params,"id"),"descField"),IIF(ArrayElement(params,"descField"),ArrayElement(params,"descField"),ArrayElement(ArrayElement(ArrayElement(pageObject.googleMapCfg,"mapsData"),ArrayElement(params,"id")),"addressField"))
	if not IsEmpty(ArrayElement(params,"zoom")) then
		setArrElementN pageObject.googleMapCfg,CreateArray3("mapsData",ArrayElement(params,"id"),"zoom"),ArrayElement(params,"zoom")
	end if
	if bValue(ArrayElement(ArrayElement(ArrayElement(pageObject.googleMapCfg,"mapsData"),ArrayElement(params,"id")),"showCenterLink")) then
		setArrElementN pageObject.googleMapCfg,CreateArray3("mapsData",ArrayElement(params,"id"),"centerLinkText"),IIF(ArrayElement(params,"centerLinkText"),ArrayElement(params,"centerLinkText"),"")
	end if
	setArrElementN pageObject.googleMapCfg,CreateArray2("mainMapIds",empty),ArrayElement(params,"id")
	if not IsEmpty(ArrayElement(params,"APIkey")) then
		setArrElement pageObject.googleMapCfg,"APIcode",ArrayElement(params,"APIkey")
	end if
End Function
Function DisplayCAPTCHA()
	pageObject.xt.assign_event_p4 pageObject.captchaId,pageObject,"createCaptcha",CreateDictionary()
End Function
Function checkTableName(ByVal shortTName,ByVal var_type)
	if not bValue(shortTName) then
		checkTableName = false
		Exit Function
	end if
	if IsEqual("payment",shortTName) and (IsFalse(var_type) or not IsFalse(var_type) and IsEqual(var_type,0)) then
		checkTableName = true
		Exit Function
	end if
	checkTableName = false
	Exit Function
End Function
Function GetEditTabs(ByVal table)
	doAssignmentByRef GetEditTabs,GetTableData(table,".arrEditTabs",CreateDictionary())
	Exit Function
End Function
Function useTabsOnEdit(ByVal table)
	if bValue(asp_count(GetEditTabs(table))) then
		useTabsOnEdit = true
		Exit Function
	else
		useTabsOnEdit = false
		Exit Function
	end if
End Function
Function GetAddTabs(ByVal table)
	doAssignmentByRef GetAddTabs,GetTableData(table,".arrAddTabs",CreateDictionary())
	Exit Function
End Function
Function useTabsOnAdd(ByVal table)
	if bValue(asp_count(GetAddTabs(table))) then
		useTabsOnAdd = true
		Exit Function
	else
		useTabsOnAdd = false
		Exit Function
	end if
End Function
Function GetViewTabs(ByVal table)
	doAssignmentByRef GetViewTabs,GetTableData(table,".arrViewTabs",CreateDictionary())
	Exit Function
End Function
Function useTabsOnView(ByVal table)
	if bValue(asp_count(GetViewTabs(table))) then
		useTabsOnView = true
		Exit Function
	else
		useTabsOnView = false
		Exit Function
	end if
End Function
Function GetMasterTablesArr(ByVal tName)
	doAssignmentByRef GetMasterTablesArr,ArrayElement(masterTablesData,tName)
	Exit Function
End Function
Function GetDetailTablesArr(ByVal tName)
	doAssignmentByRef GetDetailTablesArr,ArrayElement(detailsTablesData,tName)
	Exit Function
End Function
Function GetDetailKeysByMasterTable(ByVal mTableName,ByVal tName,ByVal default)
	Dim strTableName,mTableDataArr
	if not bValue(mTableName) then
		doAssignmentByRef GetDetailKeysByMasterTable,default
		Exit Function
	end if
	if not bValue(tName) then
		doAssignment tName,strTableName
	end if
	GetCollectionBounds ArrayElement(masterTablesData,tName),i_commonfunctions_loopIdx2,i_commonfunctions_loopMax2
	do while i_commonfunctions_loopIdx2<=i_commonfunctions_loopMax2
		i_commonfunctions_arrKey2 = GetCollectionKey(ArrayElement(masterTablesData,tName),i_commonfunctions_loopIdx2)
		doAssignment mTableDataArr,ArrayElement(ArrayElement(masterTablesData,tName),i_commonfunctions_arrKey2)
		if IsEqual(ArrayElement(mTableDataArr,"mDataSourceTable"),mTableName) then
			doAssignmentByRef GetDetailKeysByMasterTable,ArrayElement(mTableDataArr,"detailKeys")
			Exit Function
		end if
		i_commonfunctions_loopIdx2=i_commonfunctions_loopIdx2+1
	loop
	doAssignmentByRef GetDetailKeysByMasterTable,default
	Exit Function
End Function
Function GetMasterKeysByDetailTable(ByVal dTableName,ByVal tName,ByVal default)
	Dim strTableName,dTableDataArr
	if not bValue(dTableName) then
		doAssignmentByRef GetMasterKeysByDetailTable,default
		Exit Function
	end if
	if not bValue(tName) then
		doAssignment tName,strTableName
	end if
	GetCollectionBounds ArrayElement(detailsTablesData,tName),i_commonfunctions_loopIdx3,i_commonfunctions_loopMax3
	do while i_commonfunctions_loopIdx3<=i_commonfunctions_loopMax3
		i_commonfunctions_arrKey3 = GetCollectionKey(ArrayElement(detailsTablesData,tName),i_commonfunctions_loopIdx3)
		doAssignment dTableDataArr,ArrayElement(ArrayElement(detailsTablesData,tName),i_commonfunctions_arrKey3)
		if IsEqual(ArrayElement(dTableDataArr,"dDataSourceTable"),dTableName) then
			doAssignmentByRef GetMasterKeysByDetailTable,ArrayElement(dTableDataArr,"masterKeys")
			Exit Function
		end if
		i_commonfunctions_loopIdx3=i_commonfunctions_loopIdx3+1
	loop
	doAssignmentByRef GetMasterKeysByDetailTable,default
	Exit Function
End Function
Function GetDetailKeysByDetailTable(ByVal dTableName,ByVal tName,ByVal default)
	Dim dTableDataArr
	GetCollectionBounds ArrayElement(detailsTablesData,tName),i_commonfunctions_loopIdx4,i_commonfunctions_loopMax4
	do while i_commonfunctions_loopIdx4<=i_commonfunctions_loopMax4
		i_commonfunctions_arrKey4 = GetCollectionKey(ArrayElement(detailsTablesData,tName),i_commonfunctions_loopIdx4)
		doAssignment dTableDataArr,ArrayElement(ArrayElement(detailsTablesData,tName),i_commonfunctions_arrKey4)
		if IsEqual(ArrayElement(dTableDataArr,"dDataSourceTable"),dTableName) then
			doAssignmentByRef GetDetailKeysByDetailTable,ArrayElement(dTableDataArr,"detailKeys")
			Exit Function
		end if
		i_commonfunctions_loopIdx4=i_commonfunctions_loopIdx4+1
	loop
	doAssignmentByRef GetDetailKeysByDetailTable,default
	Exit Function
End Function
Function GetDPType(ByVal dTableName,ByVal tName)
	Dim strTableName,dTableDataArr
	if not bValue(dTableName) then
		GetDPType = false
		Exit Function
	end if
	if not bValue(tName) then
		doAssignment tName,strTableName
	end if
	GetCollectionBounds ArrayElement(detailsTablesData,tName),i_commonfunctions_loopIdx5,i_commonfunctions_loopMax5
	do while i_commonfunctions_loopIdx5<=i_commonfunctions_loopMax5
		i_commonfunctions_arrKey5 = GetCollectionKey(ArrayElement(detailsTablesData,tName),i_commonfunctions_loopIdx5)
		doAssignment dTableDataArr,ArrayElement(ArrayElement(detailsTablesData,tName),i_commonfunctions_arrKey5)
		if IsEqual(ArrayElement(dTableDataArr,"dDataSourceTable"),dTableName) then
			doAssignmentByRef GetDPType,ArrayElement(dTableDataArr,"previewOnList")
			Exit Function
		end if
		i_commonfunctions_loopIdx5=i_commonfunctions_loopIdx5+1
	loop
	GetDPType = false
	Exit Function
End Function
Function GetFieldByIndex(ByVal index,ByVal table)
	Dim value,key
	if not bValue(table) then
		doAssignment table,strTableName
	end if
	if not bValue(asp_array_key_exists(table,tables_data)) then
		GetFieldByIndex = null
		Exit Function
	end if
	GetCollectionBounds ArrayElement(tables_data,table),i_commonfunctions_loopIdx6,i_commonfunctions_loopMax6
	i_commonfunctions_exitLoop6=false
	do while i_commonfunctions_loopIdx6<=i_commonfunctions_loopMax6
		i_commonfunctions_exitLoop6=false
		do
			key = GetCollectionKey(ArrayElement(tables_data,table),i_commonfunctions_loopIdx6)
			doAssignment value,ArrayElement(ArrayElement(tables_data,table),key)
			if not bValue(asp_is_array(value)) or not bValue(asp_array_key_exists("Index",value)) then
				exit do
			end if
			if IsEqual(ArrayElement(value,"Index"),index) and bValue(GetFieldIndex(key,"")) then
				doAssignmentByRef GetFieldByIndex,key
				Exit Function
			end if
		loop while false
		if i_commonfunctions_exitLoop6 then _
			exit do
		i_commonfunctions_loopIdx6=i_commonfunctions_loopIdx6+1
	loop
	GetFieldByIndex = null
	Exit Function
End Function
Function Label(ByVal field,ByVal table)
	doAssignmentByRef Label,GetFieldData(table,field,"Label",field)
	Exit Function
End Function
Function GetFilenameField(ByVal field,ByVal table)
	doAssignmentByRef GetFilenameField,GetFieldData(table,field,"Filename","")
	Exit Function
End Function
Function GetLinkPrefix(ByVal field,ByVal table)
	doAssignmentByRef GetLinkPrefix,GetFieldData(table,field,"LinkPrefix","")
	Exit Function
End Function
Function GetFieldType(ByVal field,ByVal table)
	doAssignmentByRef GetFieldType,GetFieldData(table,field,"FieldType","")
	Exit Function
End Function
Function IsAutoincField(ByVal field,ByVal table)
	doAssignmentByRef IsAutoincField,GetFieldData(table,field,"AutoInc",false)
	Exit Function
End Function
Function IsUseiBox(ByVal field,ByVal table)
	doAssignmentByRef IsUseiBox,GetFieldData(table,field,"UseiBox",false)
	Exit Function
End Function
Function GetEditFormat(ByVal field,ByVal table)
	doAssignmentByRef GetEditFormat,GetFieldData(table,field,"EditFormat","")
	Exit Function
End Function
Function ViewFormat(ByVal field,ByVal table)
	doAssignmentByRef ViewFormat,GetFieldData(table,field,"ViewFormat","")
	Exit Function
End Function
Function DateEditShowTime(ByVal field,ByVal table)
	doAssignmentByRef DateEditShowTime,GetFieldData(table,field,"ShowTime",false)
	Exit Function
End Function
Function FastType(ByVal field,ByVal table)
	doAssignmentByRef FastType,GetFieldData(table,field,"FastType",false)
	Exit Function
End Function
Function LookupControlType(ByVal field,ByVal table)
	doAssignmentByRef LookupControlType,GetFieldData(table,field,"LCType",LCT_DROPDOWN)
	Exit Function
End Function
Function UseCategory(ByVal field,ByVal table)
	doAssignmentByRef UseCategory,GetFieldData(table,field,"UseCategory",false)
	Exit Function
End Function
Function Multiselect(ByVal field,ByVal table)
	doAssignmentByRef Multiselect,GetFieldData(table,field,"Multiselect",false)
	Exit Function
End Function
Function SelectSize(ByVal field,ByVal table)
	doAssignmentByRef SelectSize,GetFieldData(table,field,"SelectSize",1)
	Exit Function
End Function
Function ShowThumbnail(ByVal field,ByVal table)
	doAssignmentByRef ShowThumbnail,GetFieldData(table,field,"ShowThumbnail",false)
	Exit Function
End Function
Function GetImageWidth(ByVal field,ByVal table)
	doAssignmentByRef GetImageWidth,GetFieldData(table,field,"ImageWidth",0)
	Exit Function
End Function
Function GetImageHeight(ByVal field,ByVal table)
	doAssignmentByRef GetImageHeight,GetFieldData(table,field,"ImageHeight",0)
	Exit Function
End Function
Function GetLWWhere(ByVal field,ByVal table)
	if not bValue(table) then
		doAssignment table,strTableName
	end if
	GetLWWhere = ""
	Exit Function
End Function
Function GetLookupType(ByVal field,ByVal table)
	doAssignmentByRef GetLookupType,GetFieldData(table,field,"LookupType",0)
	Exit Function
End Function
Function GetpLookupType(ByVal field,ByVal table)
	doAssignmentByRef GetpLookupType,GetFieldData(table,field,"pLookupType",0)
	Exit Function
End Function
Function GetLookupTable(ByVal field,ByVal table)
	doAssignmentByRef GetLookupTable,GetFieldData(table,field,"LookupTable","")
	Exit Function
End Function
Function GetLWLinkField(ByVal field,ByVal table,ByVal addWrap)
	if bValue(addWrap) then
		doAssignmentByRef GetLWLinkField,AddFieldWrappers(GetFieldData(table,field,"LinkField",""))
		Exit Function
	else
		doAssignmentByRef GetLWLinkField,GetFieldData(table,field,"LinkField","")
		Exit Function
	end if
End Function
Function GetLWLinkFieldType(ByVal field,ByVal table)
	doAssignmentByRef GetLWLinkFieldType,GetFieldData(table,field,"LinkFieldType",0)
	Exit Function
End Function
Function GetLWDisplayField(ByVal field,ByVal table,ByVal addWrap)
	if bValue(addWrap) and not bValue(GetFieldData(table,field,"CustomDisplay",false)) then
		doAssignmentByRef GetLWDisplayField,AddFieldWrappers(GetFieldData(table,field,"DisplayField",""))
		Exit Function
	else
		doAssignmentByRef GetLWDisplayField,GetFieldData(table,field,"DisplayField","")
		Exit Function
	end if
End Function
Function NeedEncode(ByVal field,ByVal table)
	doAssignmentByRef NeedEncode,GetFieldData(table,field,"NeedEncode",false)
	Exit Function
End Function
Function getValidation(ByVal field,ByVal table)
	doAssignmentByRef getValidation,GetFieldData(table,field,"validateAs",CreateDictionary())
	Exit Function
End Function
Function AppearOnListPage(ByVal field,ByVal table)
	doAssignmentByRef AppearOnListPage,GetFieldData(table,field,"bListPage",false)
	Exit Function
End Function
Function AppearOnAddPage(ByVal field,ByVal table)
	doAssignmentByRef AppearOnAddPage,GetFieldData(table,field,"bAddPage",false)
	Exit Function
End Function
Function AppearOnInlineAdd(ByVal field,ByVal table)
	doAssignmentByRef AppearOnInlineAdd,GetFieldData(table,field,"bInlineAdd",false)
	Exit Function
End Function
Function AppearOnEditPage(ByVal field,ByVal table)
	doAssignmentByRef AppearOnEditPage,GetFieldData(table,field,"bEditPage",false)
	Exit Function
End Function
Function AppearOnInlineEdit(ByVal field,ByVal table)
	doAssignmentByRef AppearOnInlineEdit,GetFieldData(table,field,"bInlineEdit",false)
	Exit Function
End Function
Function AppearOnViewPage(ByVal field,ByVal table)
	doAssignmentByRef AppearOnViewPage,GetFieldData(table,field,"bViewPage",false)
	Exit Function
End Function
Function AppearOnPrinterPage(ByVal field,ByVal table)
	doAssignmentByRef AppearOnPrinterPage,GetFieldData(table,field,"bPrinterPage",false)
	Exit Function
End Function
Function AppearOnExportPage(ByVal field,ByVal table)
	doAssignmentByRef AppearOnExportPage,GetFieldData(table,field,"bExportPage",false)
	Exit Function
End Function
Function AppearOnRegisterOrSearchPage(ByVal field,ByVal pageType,ByVal table)
	Dim arrFields,match,i
	Set arrFields = (CreateDictionary())
	if IsEqual(pageType,PAGE_REGISTER) then
		doAssignmentByRef arrFields,GetTableData(table,".fieldsForRegister",CreateDictionary())
	else
		if IsEqual(pageType,PAGE_SEARCH) then
			doAssignmentByRef arrFields,GetTableData(table,".allSearchFields",CreateDictionary())
		end if
	end if
	if not bValue(asp_count(arrFields)) then
		AppearOnRegisterOrSearchPage = "break"
		Exit Function
	end if
	match = false
	i = 0
	do while IsLess(i,asp_count(arrFields))
		if IsEqual(ArrayElement(arrFields,i),field) then
			match = true
			exit do
		end if
		i = CSmartDbl(i)+1
	loop
	doAssignmentByRef AppearOnRegisterOrSearchPage,match
	Exit Function
End Function
Function AppearOnCurrentPage(ByVal fName,ByVal pageType,ByVal pageLikeInline)
	if IsEqual(pageType,PAGE_LIST) then
		if bValue(AppearOnListPage(fName,"")) then
			AppearOnCurrentPage = true
			Exit Function
		else
			doAssignmentByRef AppearOnCurrentPage,AppearOnRegisterOrSearchPage(fName,PAGE_SEARCH,"")
			Exit Function
		end if
	else
		if IsEqual(pageType,PAGE_ADD) then
			if bValue(pageLikeInline) then
				if bValue(AppearOnInlineAdd(fName,"")) and bValue(AppearOnListPage(fName,"")) then
					AppearOnCurrentPage = true
					Exit Function
				end if
			else
				if bValue(AppearOnAddPage(fName,"")) then
					AppearOnCurrentPage = true
					Exit Function
				end if
			end if
		else
			if IsEqual(pageType,PAGE_EDIT) then
				if bValue(pageLikeInline) then
					if bValue(AppearOnInlineEdit(fName,"")) and bValue(AppearOnListPage(fName,"")) then
						AppearOnCurrentPage = true
						Exit Function
					end if
				else
					if bValue(AppearOnEditPage(fName,"")) then
						AppearOnCurrentPage = true
						Exit Function
					end if
				end if
			else
				if (IsEqual(pageType,PAGE_SEARCH) or IsEqual(pageType,PAGE_REPORT)) or IsEqual(pageType,PAGE_CHART) then
					doAssignmentByRef AppearOnCurrentPage,AppearOnRegisterOrSearchPage(fName,PAGE_SEARCH,"")
					Exit Function
				else
					if IsEqual(pageType,PAGE_REGISTER) then
						doAssignmentByRef AppearOnCurrentPage,AppearOnRegisterOrSearchPage(fName,PAGE_REGISTER,"")
						Exit Function
					else
						AppearOnCurrentPage = "break"
						Exit Function
					end if
				end if
			end if
		end if
	end if
	AppearOnCurrentPage = false
	Exit Function
End Function
Function GetPasswordField(ByVal table)
	doAssignmentByRef GetPasswordField,GetTableData(table,".PasswordField","")
	Exit Function
End Function
Function GetUserNameField(ByVal table)
	doAssignmentByRef GetUserNameField,GetTableData(table,".UserNameField","")
	Exit Function
End Function
Function GetTablesList(ByVal pdfMode)
	Dim arr,strPerm
	Set arr = (CreateDictionary())
		setArrElement arr,asp_count(arr),"payment"
	doAssignmentByRef GetTablesList,arr
	Exit Function
End Function
Function GetFieldsList(ByVal table)
	Dim t,arr,f
	if not bValue(table) then
		doAssignment table,strTableName
	end if
	if not bValue(asp_array_key_exists(table,tables_data)) then
		doAssignmentByRef GetFieldsList,CreateDictionary()
		Exit Function
	end if
	doAssignmentByRef t,asp_array_keys(ArrayElement(tables_data,table),empty)
	Set arr = (CreateDictionary())
	GetCollectionBounds t,i_commonfunctions_loopIdx8,i_commonfunctions_loopMax8
	do while i_commonfunctions_loopIdx8<=i_commonfunctions_loopMax8
		i_commonfunctions_arrKey8 = GetCollectionKey(t,i_commonfunctions_loopIdx8)
		doAssignment f,ArrayElement(t,i_commonfunctions_arrKey8)
		if not IsEqual(asp_substr(f,0,1),".") then
			setArrElement arr,asp_count(arr),f
		end if
		i_commonfunctions_loopIdx8=i_commonfunctions_loopIdx8+1
	loop
	doAssignmentByRef GetFieldsList,arr
	Exit Function
End Function
Function GetBinaryFieldsIndices(ByVal table)
	Dim fields,out,f,idx
	doAssignmentByRef fields,GetFieldsList(table)
	Set out = (CreateDictionary())
	GetCollectionBounds fields,i_commonfunctions_loopIdx9,i_commonfunctions_loopMax9
	do while i_commonfunctions_loopIdx9<=i_commonfunctions_loopMax9
		idx = GetCollectionKey(fields,i_commonfunctions_loopIdx9)
		doAssignment f,ArrayElement(fields,idx)
		if bValue(IsBinaryType(GetFieldType(f,table))) then
			setArrElement out,asp_count(out),CSmartDbl(idx)+1
		end if
		i_commonfunctions_loopIdx9=i_commonfunctions_loopIdx9+1
	loop
	doAssignmentByRef GetBinaryFieldsIndices,out
	Exit Function
End Function
Function GetNBFieldsList(ByVal table)
	Dim t,arr,f
	doAssignmentByRef t,GetFieldsList(table)
	Set arr = (CreateDictionary())
	GetCollectionBounds t,i_commonfunctions_loopIdx10,i_commonfunctions_loopMax10
	do while i_commonfunctions_loopIdx10<=i_commonfunctions_loopMax10
		i_commonfunctions_arrKey10 = GetCollectionKey(t,i_commonfunctions_loopIdx10)
		doAssignment f,ArrayElement(t,i_commonfunctions_arrKey10)
		if not bValue(IsBinaryType(GetFieldType(f,table))) then
			setArrElement arr,asp_count(arr),f
		end if
		i_commonfunctions_loopIdx10=i_commonfunctions_loopIdx10+1
	loop
	doAssignmentByRef GetNBFieldsList,arr
	Exit Function
End Function
Function CategoryControl(ByVal field,ByVal table)
	doAssignmentByRef CategoryControl,GetFieldData(table,field,"CategoryControl","")
	Exit Function
End Function
Function GetCreateThumbnail(ByVal field,ByVal table)
	doAssignmentByRef GetCreateThumbnail,GetFieldData(table,field,"CreateThumbnail",false)
	Exit Function
End Function
Function GetThumbnailPrefix(ByVal field,ByVal table)
	doAssignmentByRef GetThumbnailPrefix,GetFieldData(table,field,"ThumbnailPrefix","")
	Exit Function
End Function
Function ResizeOnUpload(ByVal field,ByVal table)
	doAssignmentByRef ResizeOnUpload,GetFieldData(table,field,"ResizeImage",false)
	Exit Function
End Function
Function GetNewImageSize(ByVal field,ByVal table)
	doAssignmentByRef GetNewImageSize,GetFieldData(table,field,"NewSize",0)
	Exit Function
End Function
Function GetFieldByGoodFieldName(ByVal field,ByVal table)
	Dim value,key
	if not bValue(table) then
		doAssignment table,strTableName
	end if
	if not bValue(asp_array_key_exists(table,tables_data)) then
		GetFieldByGoodFieldName = ""
		Exit Function
	end if
	GetCollectionBounds ArrayElement(tables_data,table),i_commonfunctions_loopIdx11,i_commonfunctions_loopMax11
	do while i_commonfunctions_loopIdx11<=i_commonfunctions_loopMax11
		key = GetCollectionKey(ArrayElement(tables_data,table),i_commonfunctions_loopIdx11)
		doAssignment value,ArrayElement(ArrayElement(tables_data,table),key)
		if IsLess(1,asp_count(value)) then
			if IsEqual(ArrayElement(value,"GoodName"),field) then
				doAssignmentByRef GetFieldByGoodFieldName,key
				Exit Function
			end if
		end if
		i_commonfunctions_loopIdx11=i_commonfunctions_loopIdx11+1
	loop
	GetFieldByGoodFieldName = ""
	Exit Function
End Function
Function GetFullFieldName(ByVal field,ByVal table)
	Dim fname
	fname = (CSmartStr(AddTableWrappers(GetOriginalTableName(table))) & ".") & CSmartStr(AddFieldWrappers(field))
	doAssignmentByRef GetFullFieldName,GetFieldData(table,field,"FullName",fname)
	Exit Function
End Function
Function GetNRows(ByVal field,ByVal table)
	doAssignmentByRef GetNRows,GetFieldData(table,field,"nRows",field)
	Exit Function
End Function
Function GetNCols(ByVal field,ByVal table)
	doAssignmentByRef GetNCols,GetFieldData(table,field,"nCols",field)
	Exit Function
End Function
Function GetOriginalTableName(ByVal table)
	if not bValue(table) then
		doAssignment table,strTableName
	end if
	doAssignmentByRef GetOriginalTableName,GetTableData(table,".OriginalTable",table)
	Exit Function
End Function
Function GetTableKeys(ByVal table)
	doAssignmentByRef GetTableKeys,GetTableData(table,".Keys",CreateDictionary())
	Exit Function
End Function
Function GetNumberOfChars(ByVal table)
	doAssignmentByRef GetNumberOfChars,GetTableData(table,".NumberOfChars",0)
	Exit Function
End Function
Function GetTableURL(ByVal table)
	if not bValue(table) then
		doAssignment table,strTableName
	end if
	if IsEqual("payment",table) then
		GetTableURL = "payment"
		Exit Function
	end if
End Function
Function GetTableByShort(ByVal shortTName)
	if not bValue(shortTName) then
		GetTableByShort = false
		Exit Function
	end if
	if IsEqual("payment",shortTName) then
		GetTableByShort = "payment"
		Exit Function
	end if
End Function
Function GetTableOwnerID(ByVal table)
	doAssignmentByRef GetTableOwnerID,GetTableData(table,".OwnerID",0)
	Exit Function
End Function
Function IsRequired(ByVal field,ByVal table)
	doAssignmentByRef IsRequired,GetFieldData(table,field,"IsRequired",false)
	Exit Function
End Function
Function UseRTE(ByVal field,ByVal table)
	doAssignmentByRef UseRTE,GetFieldData(table,field,"UseRTE",false)
	Exit Function
End Function
Function UseRTEBasic(ByVal field,ByVal table)
	doAssignmentByRef UseRTEBasic,GetFieldData(table,field,"UseRTEBasic",false)
	Exit Function
End Function
Function UseRTEFCK(ByVal field,ByVal table)
	doAssignmentByRef UseRTEFCK,GetFieldData(table,field,"UseRTEFCK",false)
	Exit Function
End Function
Function UseRTEInnova(ByVal field,ByVal table)
	doAssignmentByRef UseRTEInnova,GetFieldData(table,field,"UseRTEInnova",false)
	Exit Function
End Function
Function UseTimestamp(ByVal field,ByVal table)
	doAssignmentByRef UseTimestamp,GetFieldData(table,field,"UseTimestamp",false)
	Exit Function
End Function
Function GetUploadFolder(ByVal field,ByVal table)
	Dim path
	doAssignmentByRef path,GetFieldData(table,field,"UploadFolder","")
	if bValue(asp_strlen(path)) and not IsEqual(asp_substr(path,CSmartDbl(asp_strlen(path))-1,empty),"/") then
		path = CSmartStr(path) & "/"
	end if
	doAssignmentByRef GetUploadFolder,path
	Exit Function
End Function
Function GetFieldIndex(ByVal field,ByVal table)
	doAssignmentByRef GetFieldIndex,GetFieldData(table,field,"Index",0)
	Exit Function
End Function
Function GetQueryObject(ByVal table)
	Dim ret
	ret = null
	if not bValue(asp_array_key_exists(table,tables_data)) then
		doAssignmentByRef GetQueryObject,ret
		Exit Function
	end if
	doAssignmentByRef GetQueryObject,ArrayElement(ArrayElement(tables_data,table),".sqlquery")
	Exit Function
End Function
Function GetListOfFieldsByExprType(ByVal needaggregate,ByVal table)
	Dim query,fields,out,aggr,f
	if not bValue(asp_strlen(table)) then
		doAssignment table,strTableName
	end if
	if not bValue(asp_array_key_exists(table,tables_data)) then
		doAssignmentByRef GetListOfFieldsByExprType,CreateDictionary()
		Exit Function
	end if
	doAssignmentByRef query,ArrayElement(ArrayElement(tables_data,table),".sqlquery")
	doAssignmentByRef fields,GetFieldsList(table)
	Set out = (CreateDictionary())
	GetCollectionBounds fields,i_commonfunctions_loopIdx12,i_commonfunctions_loopMax12
	do while i_commonfunctions_loopIdx12<=i_commonfunctions_loopMax12
		idx = GetCollectionKey(fields,i_commonfunctions_loopIdx12)
		doAssignment f,ArrayElement(fields,idx)
		doAssignmentByRef aggr,query.IsAggrFuncField_p1(idx)
		if bValue(needaggregate) and bValue(aggr) or not bValue(needaggregate) and not bValue(aggr) then
			setArrElement out,asp_count(out),f
		end if
		i_commonfunctions_loopIdx12=i_commonfunctions_loopIdx12+1
	loop
	doAssignmentByRef GetListOfFieldsByExprType,out
	Exit Function
End Function
Function DateEditType(ByVal field,ByVal table)
	doAssignmentByRef DateEditType,GetFieldData(table,field,"DateEditType",0)
	Exit Function
End Function
Function GetEditParams(ByVal field,ByVal table)
	doAssignmentByRef GetEditParams,GetFieldData(table,field,"EditParams","")
	Exit Function
End Function
Function GetChartType(ByVal shorttable)
	GetChartType = ""
	Exit Function
End Function
Function GetData(ByVal data,ByVal field,ByVal format)
	doAssignmentByRef GetData,GetDataInt(ArrayElement(data,field),data,field,format)
	Exit Function
End Function
Function GetDataInt(ByVal value,ByVal data,ByVal field,ByVal format)
	Dim ret,numbers,thumbnailed,thumbprefix,thumbname,link,title,iquery,arrKeys,keylink,j
	if IsEqual(format,FORMAT_CUSTOM) and bValue(data) then
		doAssignmentByRef GetDataInt,CustomExpression(value,data,field,"")
		Exit Function
	end if
	ret = ""
	if bValue(IsBinaryType(GetFieldType(field,""))) then
		ret = "LONG BINARY DATA - CANNOT BE DISPLAYED"
	else
		doAssignment ret,value
	end if
	if IsFalse(ret) then
		ret = ""
	end if
	if IsEqual(format,FORMAT_DATE_SHORT) then
		doAssignmentByRef ret,format_shortdate(db2time(value))
	else
		if IsEqual(format,FORMAT_DATE_LONG) then
			doAssignmentByRef ret,format_longdate(db2time(value))
		else
			if IsEqual(format,FORMAT_DATE_TIME) then
				doAssignmentByRef ret,str_format_datetime(db2time(value))
			else
				if IsEqual(format,FORMAT_TIME) then
					if bValue(IsDateFieldType(GetFieldType(field,""))) then
						doAssignmentByRef ret,str_format_time(db2time(value))
					else
						doAssignmentByRef numbers,parsenumbers(value)
						if not bValue(asp_count(numbers)) then
							GetDataInt = ""
							Exit Function
						end if
						do while IsLess(asp_count(numbers),3)
							setArrElement numbers,asp_count(numbers),0
						loop
						doAssignmentByRef ret,str_format_time(CreateDictionary6(Empty,0,Empty,0,Empty,0,Empty,ArrayElement(numbers,0),Empty,ArrayElement(numbers,1),Empty,ArrayElement(numbers,2)))
					end if
				else
					if IsEqual(format,FORMAT_NUMBER) then
						doAssignmentByRef ret,str_format_number(value,GetFieldData(strTableName,field,"DecimalDigits",false))
					else
						if IsEqual(format,FORMAT_CURRENCY) then
							doAssignmentByRef ret,str_format_currency(value)
						else
							if IsEqual(format,FORMAT_CHECKBOX) then
								ret = "<img src=""images/check_"
								if bValue(value) and not IsEqual(value,0) then
									ret = CSmartStr(ret) & "yes"
								else
									ret = CSmartStr(ret) & "no"
								end if
								ret = CSmartStr(ret) & ".gif"" border=0"
								if bValue(isEnableSection508()) then
									ret = CSmartStr(ret) & " alt="" """
								end if
								ret = CSmartStr(ret) & ">"
							else
								if IsEqual(format,FORMAT_PERCENT) then
									if not IsEqual(value,"") then
										ret = CSmartStr(CSmartDbl(value)*100) & "%"
									end if
								else
									if IsEqual(format,FORMAT_PHONE_NUMBER) then
										if IsEqual(asp_strlen(ret),7) then
											ret = (CSmartStr(asp_substr(ret,0,3)) & "-") & CSmartStr(asp_substr(ret,3,empty))
										else
											if IsEqual(asp_strlen(ret),10) then
												ret = (((("(" & CSmartStr(asp_substr(ret,0,3))) & ") ") & CSmartStr(asp_substr(ret,3,3))) & "-") & CSmartStr(asp_substr(ret,6,empty))
											end if
										end if
									else
										if IsEqual(format,FORMAT_FILE_IMAGE) then
											if not bValue(CheckImageExtension(ret)) then
												GetDataInt = ""
												Exit Function
											end if
											thumbnailed = false
											thumbprefix = ""
											if bValue(thumbnailed) then
												thumbname = CSmartStr(thumbprefix) & CSmartStr(ret)
												if not IsEqual(asp_substr(GetLinkPrefix(field,""),0,7),"http://") and not bValue(myfile_exists(getabspath(CSmartStr(GetLinkPrefix(field,"")) & CSmartStr(thumbname)))) then
													doAssignment thumbname,ret
												end if
												ret = ("<a target=_blank href=""" & CSmartStr(htmlspecialchars(AddLinkPrefix(field,ret,"")))) & """>"
												ret = CSmartStr(ret) & "<img"
												if bValue(isEnableSection508()) then
													ret = CSmartStr(ret) & ((" alt=""" & CSmartStr(htmlspecialchars(ArrayElement(data,field)))) & """")
												end if
												ret = CSmartStr(ret) & " border=0"
												ret = CSmartStr(ret) & ((" src=""" & CSmartStr(htmlspecialchars(AddLinkPrefix(field,thumbname,"")))) & """></a>")
											else
												if bValue(isEnableSection508()) then
													ret = ("<img alt=\"""".htmlspecialchars($data[$field]).""\"" src=""" & CSmartStr(AddLinkPrefix(field,ret,""))) & """ border=0>"
												else
													ret = ("<img src=""" & CSmartStr(htmlspecialchars(AddLinkPrefix(field,ret,"")))) & """ border=0>"
												end if
											end if
										else
											if IsEqual(format,FORMAT_HYPERLINK) then
												if bValue(data) then
													doAssignmentByRef ret,GetHyperlink(ret,field,data,"")
												end if
											else
												if IsEqual(format,FORMAT_EMAILHYPERLINK) then
													doAssignment link,ret
													doAssignment title,ret
													if IsEqual(asp_substr(ret,0,7),"mailto:") then
														doAssignmentByRef title,asp_substr(ret,8,empty)
													else
														link = "mailto:" & CSmartStr(link)
													end if
													ret = ((("<a href=""" & CSmartStr(link)) & """>") & CSmartStr(title)) & "</a>"
												else
													if IsEqual(format,FORMAT_FILE) then
														iquery = (("table=" & CSmartStr(GetTableURL(strTableName))) & "&field=") & CSmartStr(asp_rawurlencode(field))
														doAssignmentByRef arrKeys,GetTableKeys(strTableName)
														keylink = ""
														j = 0
														do while IsLess(j,asp_count(arrKeys))
															keylink = CSmartStr(keylink) & ((("&key" & CSmartStr(CSmartDbl(j)+1)) & "=") & CSmartStr(asp_rawurlencode(ArrayElement(data,ArrayElement(arrKeys,j)))))
															j = CSmartDbl(j)+1
														loop
														iquery = CSmartStr(iquery) & CSmartStr(keylink)
														GetDataInt = ((("<a href=""download.asp?" & CSmartStr(iquery)) & """>") & CSmartStr(htmlspecialchars(ret))) & "</a>"
														Exit Function
													else
														if IsEqual(GetEditFormat(field,""),EDIT_FORMAT_CHECKBOX) and IsEqual(format,FORMAT_NONE) then
															if bValue(ret) and not IsEqual(ret,0) then
																ret = "Yes"
															else
																ret = "No"
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
					end if
				end if
			end if
		end if
	end if
	doAssignmentByRef GetDataInt,ret
	Exit Function
End Function
Function ProcessLargeText(ByVal strValue,ByVal iquery,ByVal table,ByVal mode,ByVal format)
	Dim cNumberOfChars,useUTF8,ret
	if not bValue(table) then
		doAssignment table,strTableName
	end if
	doAssignmentByRef cNumberOfChars,GetNumberOfChars(table)
	if IsEqual(asp_substr(strValue,0,8),"<a href=") then
		doAssignmentByRef ProcessLargeText,strValue
		Exit Function
	end if
	if IsEqual(asp_substr(strValue,0,23),"<img src=""images/check_") then
		doAssignmentByRef ProcessLargeText,strValue
		Exit Function
	end if
	useUTF8 = false
	if ((not IsEqual(format,EDIT_FORMAT_LOOKUP_WIZARD) and IsLess(0,cNumberOfChars)) and IsLess(cNumberOfChars,asp_strlen(strValue))) and IsEqual(mode,MODE_LIST) then
		doAssignmentByRef table,GetTableURL(table)
		if bValue(useUTF8) then
			doAssignmentByRef ret,utf8_substr(strValue,0,cNumberOfChars)
		else
			doAssignmentByRef ret,asp_substr(strValue,0,cNumberOfChars)
		end if
		doAssignmentByRef ret,htmlspecialchars(ret)
		ret = CSmartStr(ret) & ((((((" <a href=""javascript:void(0);"" query=""fulltext.asp?table=" & CSmartStr(table)) & "&") & CSmartStr(iquery)) & """>") & CSmartStr("More")) & " ...</a>")
	else
		if ((not IsEqual(format,EDIT_FORMAT_LOOKUP_WIZARD) and IsLess(0,cNumberOfChars)) and IsLess(cNumberOfChars,asp_strlen(strValue))) and IsEqual(mode,MODE_PRINT) then
			if bValue(useUTF8) then
				doAssignmentByRef ret,utf8_substr(strValue,0,cNumberOfChars)
			else
				doAssignmentByRef ret,asp_substr(strValue,0,cNumberOfChars)
			end if
			doAssignmentByRef ret,htmlspecialchars(ret)
			if IsLess(cNumberOfChars,asp_strlen(strValue)) then
				ret = CSmartStr(ret) & " ..."
			end if
		else
			doAssignmentByRef ret,htmlspecialchars(strValue)
		end if
	end if
	doAssignmentByRef ProcessLargeText,asp_nl2br(ret)
	Exit Function
End Function
Function GetHyperlink(ByVal str,ByVal field,ByVal data,ByVal table)
	Dim ret,title,link,i,target,var_type,prefix
	if not bValue(asp_strlen(table)) then
		doAssignment table,strTableName
	end if
	if not bValue(asp_strlen(str)) then
		GetHyperlink = ""
		Exit Function
	end if
	doAssignment ret,str
	doAssignment title,ret
	doAssignment link,ret
	if IsEqual(asp_substr(ret,CSmartDbl(asp_strlen(ret))-1,empty),"#") then
		doAssignmentByRef i,asp_strpos(ret,"#",empty)
		doAssignmentByRef title,asp_substr(ret,0,i)
		doAssignmentByRef link,asp_substr(ret,CSmartDbl(i)+1,(CSmartDbl(asp_strlen(ret))-CSmartDbl(i))-2)
		if not bValue(title) then
			doAssignment title,link
		end if
	end if
	target = ""
	if IsFalse(asp_strpos(link,"://",empty)) and not IsEqual(asp_substr(link,0,7),"mailto:") then
		link = CSmartStr(prefix) & CSmartStr(link)
	end if
	ret = ((((("<a href=""" & CSmartStr(link)) & """") & CSmartStr(target)) & ">") & CSmartStr(title)) & "</a>"
	doAssignmentByRef GetHyperlink,ret
	Exit Function
End Function
Function AddLinkPrefix(ByVal field,ByVal link,ByVal table)
	if IsFalse(asp_strpos(link,"://",empty)) and not IsEqual(asp_substr(link,0,7),"mailto:") then
		AddLinkPrefix = CSmartStr(GetLinkPrefix(field,table)) & CSmartStr(link)
		Exit Function
	end if
	doAssignmentByRef AddLinkPrefix,link
	Exit Function
End Function
Function GetTotalsForTime(ByVal value)
	Dim time
	doAssignmentByRef time,parsenumbers(value)
	do while IsLess(asp_count(time),3)
		setArrElement time,asp_count(time),0
	loop
	doAssignmentByRef GetTotalsForTime,time
	Exit Function
End Function
Function GetTotals(ByVal field,ByVal value,ByVal stype,ByVal iNumberOfRows,ByVal sFormat)
	Dim days,s,m,h,d,sValue,data
	days = 0
	if IsEqual(stype,"AVERAGE") then
		if bValue(iNumberOfRows) then
			if IsEqual(sFormat,FORMAT_TIME) then
				if bValue(value) then
					doAssignmentByRef value,round(CSmartDbl(value)/CSmartDbl(iNumberOfRows),0)
					s = value mod 60
					value = CSmartDbl(value)-CSmartDbl(s)
					value = CSmartDbl(value)/60
					m = value mod 60
					value = CSmartDbl(value)-CSmartDbl(m)
					value = CSmartDbl(value)/60
					h = value mod 24
					value = CSmartDbl(value)-CSmartDbl(h)
					value = CSmartDbl(value)/24
					doAssignment d,value
					value = CSmartStr(IIF(not IsEqual(d,0),CSmartStr(d) & "d ","")) & CSmartStr(mysprintf("%02d:%02d:%02d",CreateDictionary3(Empty,h,Empty,m,Empty,s)))
				end if
			else
				doAssignmentByRef value,round(CSmartDbl(value)/CSmartDbl(iNumberOfRows),2)
			end if
		else
			GetTotals = ""
			Exit Function
		end if
	end if
	if IsEqual(stype,"TOTAL") then
		if IsEqual(sFormat,FORMAT_TIME) then
			if bValue(value) then
				s = value mod 60
				value = CSmartDbl(value)-CSmartDbl(s)
				value = CSmartDbl(value)/60
				m = value mod 60
				value = CSmartDbl(value)-CSmartDbl(m)
				value = CSmartDbl(value)/60
				h = value mod 24
				value = CSmartDbl(value)-CSmartDbl(h)
				value = CSmartDbl(value)/24
				doAssignment d,value
				value = CSmartStr(IIF(not IsEqual(d,0),CSmartStr(d) & "d ","")) & CSmartStr(mysprintf("%02d:%02d:%02d",CreateDictionary3(Empty,h,Empty,m,Empty,s)))
			end if
		end if
	end if
	sValue = ""
	Set data = (CreateDictionary1(field,value))
	if IsEqual(sFormat,FORMAT_CURRENCY) then
		doAssignmentByRef sValue,str_format_currency(value)
	else
		if IsEqual(sFormat,FORMAT_PERCENT) then
			sValue = CSmartStr(str_format_number(CSmartDbl(value)*100,false)) & "%"
		else
			if IsEqual(sFormat,FORMAT_NUMBER) then
				doAssignmentByRef sValue,str_format_number(value,GetFieldData(strTableName,field,"DecimalDigits",false))
			else
				if IsEqual(sFormat,FORMAT_CUSTOM) and not IsEqual(stype,"COUNT") then
					doAssignmentByRef sValue,GetData(data,field,sFormat)
				else
					doAssignment sValue,value
				end if
			end if
		end if
	end if
	if IsEqual(stype,"COUNT") then
		doAssignmentByRef GetTotals,value
		Exit Function
	end if
	if IsEqual(stype,"TOTAL") then
		doAssignmentByRef GetTotals,sValue
		Exit Function
	end if
	if IsEqual(stype,"AVERAGE") then
		doAssignmentByRef GetTotals,sValue
		Exit Function
	end if
	GetTotals = ""
	Exit Function
End Function
Function DisplayLookupWizard(ByVal field,ByVal value,ByVal data,ByVal keylink,ByVal mode)
	Dim LookupSQL,where,lookupvalue,iquery,out,arr,numeric,var_type,val,var_in,rsLookup,found,lookuprow,strdata
	if not bValue(asp_strlen(value)) then
		DisplayLookupWizard = ""
		Exit Function
	end if
	LookupSQL = "SELECT "
	LookupSQL = CSmartStr(LookupSQL) & CSmartStr(GetLWDisplayField(field,"",true))
	LookupSQL = CSmartStr(LookupSQL) & ((" FROM " & CSmartStr(AddTableWrappers(GetLookupTable(field,"")))) & " WHERE ")
	where = ""
	doAssignment lookupvalue,value
	iquery = ("field=" & CSmartStr(htmlspecialchars(asp_rawurlencode(field)))) & CSmartStr(keylink)
	out = ""
	if bValue(Multiselect(field,"")) then
		doAssignmentByRef arr,splitvalues(value)
		numeric = true
		doAssignmentByRef var_type,GetLWLinkFieldType(field,"")
		if not bValue(var_type) then
			GetCollectionBounds arr,i_commonfunctions_loopIdx16,i_commonfunctions_loopMax16
			do while i_commonfunctions_loopIdx16<=i_commonfunctions_loopMax16
				i_commonfunctions_arrKey16 = GetCollectionKey(arr,i_commonfunctions_loopIdx16)
				doAssignment val,ArrayElement(arr,i_commonfunctions_arrKey16)
				if bValue(asp_strlen(val)) and not bValue(asp_is_numeric(val)) then
					numeric = false
					exit do
				end if
				i_commonfunctions_loopIdx16=i_commonfunctions_loopIdx16+1
			loop
		else
			numeric = not bValue(NeedQuotes(var_type))
		end if
		var_in = ""
		GetCollectionBounds arr,i_commonfunctions_loopIdx17,i_commonfunctions_loopMax17
		i_commonfunctions_exitLoop17=false
		do while i_commonfunctions_loopIdx17<=i_commonfunctions_loopMax17
			i_commonfunctions_exitLoop17=false
			do
				i_commonfunctions_arrKey17 = GetCollectionKey(arr,i_commonfunctions_loopIdx17)
				doAssignment val,ArrayElement(arr,i_commonfunctions_arrKey17)
				if bValue(numeric) and not bValue(asp_strlen(val)) then
					exit do
				end if
				if bValue(asp_strlen(var_in)) then
					var_in = CSmartStr(var_in) & ","
				end if
				if bValue(numeric) then
					var_in = CSmartStr(var_in) & CSmartStr(CSmartDbl(val)+0)
				else
					var_in = CSmartStr(var_in) & (("'" & CSmartStr(db_addslashes(val))) & "'")
				end if
			loop while false
			if i_commonfunctions_exitLoop17 then _
				exit do
			i_commonfunctions_loopIdx17=i_commonfunctions_loopIdx17+1
		loop
		if bValue(asp_strlen(var_in)) then
			LookupSQL = CSmartStr(LookupSQL) & (((CSmartStr(GetLWLinkField(field,"",true)) & " in (") & CSmartStr(var_in)) & ")")
			doAssignmentByRef where,GetLWWhere(field,"")
			if bValue(asp_strlen(where)) then
				LookupSQL = CSmartStr(LookupSQL) & ((" and (" & CSmartStr(where)) & ")")
			end if
			LogInfo LookupSQL
			doAssignmentByRef rsLookup,db_query(LookupSQL,conn)
			found = false
			do while bValue(doAssignmentByRef(lookuprow,db_fetch_numarray(rsLookup)))
				doAssignment lookupvalue,ArrayElement(lookuprow,0)
				if bValue(found) then
					out = CSmartStr(out) & ","
				end if
				found = true
				out = CSmartStr(out) & CSmartStr(GetDataInt(lookupvalue,data,field,ViewFormat(field,"")))
			loop
			if bValue(found) then
				if bValue(NeedEncode(field,"")) and not IsEqual(mode,MODE_EXPORT) then
					doAssignmentByRef DisplayLookupWizard,ProcessLargeText(out,iquery,"",mode,GetEditFormat(field,""))
					Exit Function
				else
					doAssignmentByRef DisplayLookupWizard,out
					Exit Function
				end if
			end if
		end if
	else
		doAssignmentByRef strdata,make_db_value(field,value,"","","")
		LookupSQL = CSmartStr(LookupSQL) & ((CSmartStr(GetLWLinkField(field,"",true)) & " = ") & CSmartStr(strdata))
		doAssignmentByRef where,GetLWWhere(field,"")
		if bValue(asp_strlen(where)) then
			LookupSQL = CSmartStr(LookupSQL) & ((" and (" & CSmartStr(where)) & ")")
		end if
		LogInfo LookupSQL
		doAssignmentByRef rsLookup,db_query(LookupSQL,conn)
		if bValue(doAssignmentByRef(lookuprow,db_fetch_numarray(rsLookup))) then
			doAssignment lookupvalue,ArrayElement(lookuprow,0)
		end if
	end if
	if not bValue(out) then
		doAssignmentByRef out,GetDataInt(lookupvalue,data,field,ViewFormat(field,""))
	end if
	if bValue(NeedEncode(field,"")) and not IsEqual(mode,MODE_EXPORT) then
		doAssignmentByRef value,ProcessLargeText(out,iquery,"",mode,GetEditFormat(field,""))
	else
		doAssignment value,out
	end if
	doAssignmentByRef DisplayLookupWizard,value
	Exit Function
End Function
Function DisplayNoImage()
	Dim path
	doAssignmentByRef path,getabspath("images/no_image.gif")
	asp_header "Content-Type: image/gif"
	printfile path
End Function
Function DisplayFile()
	Dim path
	doAssignmentByRef path,getabspath("images/file.gif")
	asp_header "Content-Type: image/gif"
	printfile path
End Function
Function my_strrpos(ByVal haystack,ByVal needle)
	Dim index
	doAssignmentByRef index,asp_strpos(asp_strrev(haystack),asp_strrev(needle),empty)
	if IsFalse(index) then
		my_strrpos = false
		Exit Function
	end if
	index = (CSmartDbl(asp_strlen(haystack))-CSmartDbl(asp_strlen(needle)))-CSmartDbl(index)
	doAssignmentByRef my_strrpos,index
	Exit Function
End Function
Function strlen_utf8(ByVal str)
	Dim len,i,olen,c
	len = 0
	i = 0
	doAssignmentByRef olen,asp_strlen(str)
	do while IsLess(i,olen)
		doAssignmentByRef c,asc(asp_substr(str,i,1))
		if IsLess(c,128) then
			i = CSmartDbl(i)+1
		else
			if (IsLess(i,CSmartDbl(olen)-1) and IsLessOrEqual(192,c)) and IsLessOrEqual(c,223) then
				i = CSmartDbl(i)+2
			else
				if (IsLess(i,CSmartDbl(olen)-2) and IsLessOrEqual(224,c)) and IsLessOrEqual(c,239) then
					i = CSmartDbl(i)+3
				else
					if IsLess(i,CSmartDbl(olen)-3) and IsLessOrEqual(240,c) then
						i = CSmartDbl(i)+4
					else
						exit do
					end if
				end if
			end if
		end if
		len = CSmartDbl(len)+1
	loop
	doAssignmentByRef strlen_utf8,len
	Exit Function
End Function
Function substr_utf8(ByVal str,ByVal index,ByVal strlen)
	Dim len,i,olen,oindex,c
	if IsLessOrEqual(strlen,0) then
		substr_utf8 = ""
		Exit Function
	end if
	len = 0
	i = 0
	doAssignmentByRef olen,asp_strlen(str)
	oindex = -1
	do while IsLess(i,olen)
		if IsEqual(len,index) then
			doAssignment oindex,i
		end if
		doAssignmentByRef c,asc(asp_substr(str,i,1))
		if IsLess(c,128) then
			i = CSmartDbl(i)+1
		else
			if (IsLess(i,CSmartDbl(olen)-1) and IsLessOrEqual(192,c)) and IsLessOrEqual(c,223) then
				i = CSmartDbl(i)+2
			else
				if (IsLess(i,CSmartDbl(olen)-2) and IsLessOrEqual(224,c)) and IsLessOrEqual(c,239) then
					i = CSmartDbl(i)+3
				else
					if IsLess(i,CSmartDbl(olen)-3) and IsLessOrEqual(240,c) then
						i = CSmartDbl(i)+4
					else
						exit do
					end if
				end if
			end if
		end if
		len = CSmartDbl(len)+1
		if IsLessOrEqual(0,oindex) and IsEqual(len,CSmartDbl(index)+CSmartDbl(strlen)) then
			doAssignmentByRef substr_utf8,asp_substr(str,oindex,CSmartDbl(i)-CSmartDbl(oindex))
			Exit Function
		end if
	loop
	if IsLess(0,oindex) then
		doAssignmentByRef substr_utf8,asp_substr(str,oindex,CSmartDbl(olen)-CSmartDbl(oindex))
		Exit Function
	end if
	substr_utf8 = ""
	Exit Function
End Function
Function jsreplace(ByVal str)
	Dim ret
	doAssignmentByRef ret,asp_str_replace(CreateDictionary4(Empty,"\",Empty,"'",Empty,vbcr,Empty,vblf),CreateDictionary4(Empty,"\\",Empty,"\'",Empty,"\r",Empty,"\n"),str)
	doAssignmentByRef jsreplace,my_str_ireplace("</script>","</scr'+'ipt>",ret)
	Exit Function
End Function
Function LogInfo(ByVal SQL)
End Function
Function CheckImageExtension(ByVal filename)
	Dim ext
	if IsLess(asp_strlen(filename),4) then
		CheckImageExtension = false
		Exit Function
	end if
	doAssignmentByRef ext,asp_strtoupper(asp_substr(filename,CSmartDbl(asp_strlen(filename))-4,empty))
	if (((IsEqual(ext,".GIF") or IsEqual(ext,".JPG")) or IsEqual(ext,"JPEG")) or IsEqual(ext,".PNG")) or IsEqual(ext,".BMP") then
		doAssignmentByRef CheckImageExtension,ext
		Exit Function
	end if
	CheckImageExtension = false
	Exit Function
End Function
Function RTESafe(ByVal strText)
	Dim tmpString
	tmpString = ""
	doAssignmentByRef tmpString,trim(strText)
	if not bValue(tmpString) then
		RTESafe = ""
		Exit Function
	end if
	doAssignmentByRef tmpString,asp_str_replace(chr(145),chr(39),tmpString)
	doAssignmentByRef tmpString,asp_str_replace(chr(146),chr(39),tmpString)
	doAssignmentByRef tmpString,asp_str_replace("'","&#39;",tmpString)
	doAssignmentByRef tmpString,asp_str_replace(chr(147),chr(34),tmpString)
	doAssignmentByRef tmpString,asp_str_replace(chr(148),chr(34),tmpString)
	doAssignmentByRef tmpString,asp_str_replace(chr(10)," ",tmpString)
	doAssignmentByRef tmpString,asp_str_replace(chr(13)," ",tmpString)
	doAssignmentByRef RTESafe,tmpString
	Exit Function
End Function
Function html_special_decode(ByVal str)
	Dim ret
	doAssignment ret,str
	doAssignmentByRef ret,asp_str_replace("&gt;",">",ret)
	doAssignmentByRef ret,asp_str_replace("&lt;","<",ret)
	doAssignmentByRef ret,asp_str_replace("&quot;","""",ret)
	doAssignmentByRef ret,asp_str_replace("&#039;","'",ret)
	doAssignmentByRef ret,asp_str_replace("&#39;","'",ret)
	doAssignmentByRef ret,asp_str_replace("&amp;","&",ret)
	doAssignmentByRef html_special_decode,ret
	Exit Function
End Function
Function CalcSearchParameters()
	Dim sWhere,var_SESSION,strSearchFor,sfor,strSearchFor2,var_type,f,strSearchOption,where
	sWhere = ""
	if IsEqual(Session(CSmartStr(strTableName) & "_search"),2) then
		GetCollectionBounds Session(CSmartStr(strTableName) & "_asearchfor"),i_commonfunctions_loopIdx21,i_commonfunctions_loopMax21
		do while i_commonfunctions_loopIdx21<=i_commonfunctions_loopMax21
			f = GetCollectionKey(Session(CSmartStr(strTableName) & "_asearchfor"),i_commonfunctions_loopIdx21)
			doAssignment sfor,ArrayElement(Session(CSmartStr(strTableName) & "_asearchfor"),f)
			doAssignmentByRef strSearchFor,trim(sfor)
			strSearchFor2 = ""
			doAssignment var_type,ArrayElement(Session(CSmartStr(strTableName) & "_asearchfortype"),f)
			if bValue(asp_array_key_exists(f,Session(CSmartStr(strTableName) & "_asearchfor2"))) then
				doAssignmentByRef strSearchFor2,trim(ArrayElement(Session(CSmartStr(strTableName) & "_asearchfor2"),f))
			end if
			if not IsEqual(strSearchFor,"") or bValue(true) then
				if not bValue(sWhere) then
					if IsEqual(Session(CSmartStr(strTableName) & "_asearchtype"),"and") then
						sWhere = "1=1"
					else
						sWhere = "1=0"
					end if
				end if
				doAssignmentByRef strSearchOption,trim(ArrayElement(Session(CSmartStr(strTableName) & "_asearchopt"),f))
				if bValue(doAssignmentByRef(where,StrWhereAdv(f,strSearchFor,strSearchOption,strSearchFor2,var_type,false))) then
					if bValue(ArrayElement(Session(CSmartStr(strTableName) & "_asearchnot"),f)) then
						where = ("not (" & CSmartStr(where)) & ")"
					end if
					if IsEqual(Session(CSmartStr(strTableName) & "_asearchtype"),"and") then
						sWhere = CSmartStr(sWhere) & (" and " & CSmartStr(where))
					else
						sWhere = CSmartStr(sWhere) & (" or " & CSmartStr(where))
					end if
				end if
			end if
			i_commonfunctions_loopIdx21=i_commonfunctions_loopIdx21+1
		loop
	end if
	doAssignmentByRef CalcSearchParameters,sWhere
	Exit Function
End Function
Function gSQLWhere(ByVal where,ByVal having)
	Dim sqlHead,sqlGroupBy,oHaving,sqlHaving
	doAssignmentByRef sqlHead,gQuery.HeadToSql()
	doAssignmentByRef sqlGroupBy,gQuery.GroupByToSql()
	doAssignmentByRef oHaving,gQuery.Having()
	doAssignmentByRef sqlHaving,oHaving.toSql_p1(gQuery)
	doAssignmentByRef gSQLWhere,gSQLWhere_having(sqlHead,gsqlFrom,gsqlWhereExpr,sqlGroupBy,sqlHaving,where,having)
	Exit Function
End Function
Function gSQLWhere_having(ByVal sqlHead,ByVal sqlFrom,ByVal sqlWhere,ByVal sqlGroupBy,ByVal sqlHaving,ByVal where,ByVal having)
	Dim strWhere,strHaving
	doAssignmentByRef strWhere,whereAdd(sqlWhere,where)
	if bValue(asp_strlen(strWhere)) then
		strWhere = (" where " & CSmartStr(strWhere)) & " "
	end if
	doAssignmentByRef strHaving,whereAdd(sqlHaving,having)
	if bValue(asp_strlen(strHaving)) then
		strHaving = (" having " & CSmartStr(strHaving)) & " "
	end if
	gSQLWhere_having = (((((((CSmartStr(sqlHead) & " ") & CSmartStr(sqlFrom)) & " ") & CSmartStr(strWhere)) & " ") & CSmartStr(sqlGroupBy)) & " ") & CSmartStr(strHaving)
	Exit Function
End Function
Function whereAdd(ByVal where,ByVal clause)
	if not bValue(asp_strlen(clause)) then
		doAssignmentByRef whereAdd,where
		Exit Function
	end if
	if not bValue(asp_strlen(where)) then
		doAssignmentByRef whereAdd,clause
		Exit Function
	end if
	whereAdd = ((("(" & CSmartStr(where)) & ") and (") & CSmartStr(clause)) & ")"
	Exit Function
End Function
Function AddWhere(ByVal sql,ByVal where)
	Dim tsql,n,n1,n2
	if not bValue(asp_strlen(where)) then
		doAssignmentByRef AddWhere,sql
		Exit Function
	end if
	doAssignmentByRef sql,asp_str_replace(CreateDictionary3(Empty,vbcrlf,Empty,vblf,Empty,"	")," ",sql)
	doAssignmentByRef tsql,asp_strtolower(sql)
	doAssignmentByRef n,my_strrpos(tsql," where ")
	doAssignmentByRef n1,my_strrpos(tsql," group by ")
	doAssignmentByRef n2,my_strrpos(tsql," order by ")
	if IsFalse(n1) then
		doAssignmentByRef n1,asp_strlen(tsql)
	end if
	if IsFalse(n2) then
		doAssignmentByRef n2,asp_strlen(tsql)
	end if
	if IsLess(n2,n1) then
		doAssignment n1,n2
	end if
	if IsFalse(n) then
		AddWhere = ((CSmartStr(asp_substr(sql,0,n1)) & " where ") & CSmartStr(where)) & CSmartStr(asp_substr(sql,n1,empty))
		Exit Function
	else
		AddWhere = (((((CSmartStr(asp_substr(sql,0,CSmartDbl(n)+CSmartDbl(asp_strlen(" where ")))) & "(") & CSmartStr(asp_substr(sql,CSmartDbl(n)+CSmartDbl(asp_strlen(" where ")),(CSmartDbl(n1)-CSmartDbl(n))-CSmartDbl(asp_strlen(" where "))))) & ") and (") & CSmartStr(where)) & ")") & CSmartStr(asp_substr(sql,n1,empty))
		Exit Function
	end if
End Function
Function KeyWhere(ByRef keys,ByVal table)
	Dim strWhere,keyFields,value,kf,valueisnull
	if not bValue(table) then
		doAssignment table,strTableName
	end if
	strWhere = ""
	doAssignmentByRef keyFields,GetTableKeys(table)
	GetCollectionBounds keyFields,i_commonfunctions_loopIdx22,i_commonfunctions_loopMax22
	do while i_commonfunctions_loopIdx22<=i_commonfunctions_loopMax22
		i_commonfunctions_arrKey22 = GetCollectionKey(keyFields,i_commonfunctions_loopIdx22)
		doAssignment kf,ArrayElement(keyFields,i_commonfunctions_arrKey22)
		if bValue(asp_strlen(strWhere)) then
			strWhere = CSmartStr(strWhere) & " and "
		end if
		doAssignmentByRef value,make_db_value(kf,ArrayElement(keys,kf),"","","")
		valueisnull = IsIdentical(value,"null")
		if bValue(valueisnull) then
			strWhere = CSmartStr(strWhere) & (CSmartStr(GetFullFieldName(kf,table)) & " is null")
		else
			strWhere = CSmartStr(strWhere) & ((CSmartStr(GetFullFieldName(kf,table)) & "=") & CSmartStr(make_db_value(kf,ArrayElement(keys,kf),"","","")))
		end if
		i_commonfunctions_loopIdx22=i_commonfunctions_loopIdx22+1
	loop
	doAssignmentByRef KeyWhere,strWhere
	Exit Function
End Function
Function StrWhereExpression(ByVal strField,ByVal SearchFor,ByVal strSearchOption,ByVal SearchFor2)
	Dim var_type,ismssql,isdb2,isMysql,btexttype,strQuote,sSearchFor,sSearchFor2,time,ret,var_like,sSearchForMysql
	doAssignmentByRef var_type,GetFieldType(strField,"")
	ismssql = true
	isdb2 = false
	isMysql = false
	doAssignmentByRef btexttype,IsTextType(var_type)
	if IsEqual(strSearchOption,"Empty") then
		if bValue(IsCharType(var_type)) and (not bValue(ismssql) or not bValue(btexttype)) then
			StrWhereExpression = ((("(" & CSmartStr(GetFullFieldName(strField,""))) & " is null or ") & CSmartStr(GetFullFieldName(strField,""))) & "='')"
			Exit Function
		else
			if bValue(ismssql) and bValue(btexttype) then
				StrWhereExpression = ((("(" & CSmartStr(GetFullFieldName(strField,""))) & " is null or ") & CSmartStr(GetFullFieldName(strField,""))) & " LIKE '')"
				Exit Function
			else
				StrWhereExpression = CSmartStr(GetFullFieldName(strField,"")) & " is null"
				Exit Function
			end if
		end if
	end if
	strQuote = ""
	if bValue(NeedQuotes(var_type)) then
		strQuote = "'"
	end if
	doAssignment sSearchFor,SearchFor
	doAssignment sSearchFor2,SearchFor2
	if bValue(IsBinaryType(var_type)) then
		StrWhereExpression = ""
		Exit Function
	end if
	if (bValue(btexttype) and not IsEqual(strSearchOption,"Contains")) and not IsEqual(strSearchOption,"Starts with") then
		StrWhereExpression = ""
		Exit Function
	end if
	if (bValue(IsDateFieldType(var_type)) and not IsEqual(strSearchOption,"Contains")) and not IsEqual(strSearchOption,"Starts with") then
		doAssignmentByRef time,localdatetime2db(SearchFor,"")
		if IsEqual(time,"null") then
			StrWhereExpression = ""
			Exit Function
		end if
		doAssignmentByRef sSearchFor,db_datequotes(time)
		if IsEqual(strSearchOption,"Between") then
			doAssignmentByRef time,localdatetime2db(SearchFor2,"")
			if IsEqual(time,"null") then
				sSearchFor2 = ""
			else
				doAssignmentByRef sSearchFor2,db_datequotes(time)
			end if
		end if
	end if
	if (not bValue(strQuote) and not bValue(asp_is_numeric(sSearchFor))) and not bValue(asp_is_numeric(sSearchFor)) then
		StrWhereExpression = ""
		Exit Function
	else
		if (not bValue(strQuote) and not IsEqual(strSearchOption,"Contains")) and not IsEqual(strSearchOption,"Starts with") then
			sSearchFor = 0+CSmartDbl(sSearchFor)
			sSearchFor2 = 0+CSmartDbl(sSearchFor2)
		else
			if (not bValue(IsDateFieldType(var_type)) and not IsEqual(strSearchOption,"Contains")) and not IsEqual(strSearchOption,"Starts with") then
				if bValue(btexttype) then
					sSearchFor = (CSmartStr(strQuote) & CSmartStr(db_addslashes(sSearchFor))) & CSmartStr(strQuote)
					if IsEqual(strSearchOption,"Between") and bValue(sSearchFor2) then
						sSearchFor2 = (CSmartStr(strQuote) & CSmartStr(db_addslashes(sSearchFor2))) & CSmartStr(strQuote)
					end if
				else
					doAssignmentByRef sSearchFor,isEnableUpper((CSmartStr(strQuote) & CSmartStr(db_addslashes(sSearchFor))) & CSmartStr(strQuote))
					if IsEqual(strSearchOption,"Between") and bValue(sSearchFor2) then
						doAssignmentByRef sSearchFor2,isEnableUpper((CSmartStr(strQuote) & CSmartStr(db_addslashes(sSearchFor2))) & CSmartStr(strQuote))
					end if
				end if
			else
				if (not bValue(IsDateFieldType(var_type)) or IsEqual(strSearchOption,"Contains")) or IsEqual(strSearchOption,"Starts with") then
					doAssignmentByRef sSearchFor,db_addslashes(sSearchFor)
				end if
			end if
		end if
	end if
	if bValue(IsCharType(var_type)) and not bValue(btexttype) then
		doAssignmentByRef strField,isEnableUpper(GetFullFieldName(strField,""))
	else
		if IsEqual(strSearchOption,"Contains") or IsEqual(strSearchOption,"Starts with") then
			doAssignmentByRef strField,db_field2char(GetFullFieldName(strField,""),var_type)
		else
			if IsEqual(ViewFormat(strField,""),FORMAT_TIME) then
				doAssignmentByRef strField,db_field2time(GetFullFieldName(strField,""),var_type)
			else
				doAssignmentByRef strField,GetFullFieldName(strField,"")
			end if
		end if
	end if
	ret = ""
	var_like = "like"
	if bValue(isMysql) then
		doAssignmentByRef sSearchForMysql,asp_str_replace("\\","\\\\",sSearchFor)
	end if
	if IsEqual(strSearchOption,"Contains") then
		if bValue(isMysql) then
			doAssignment sSearchFor,sSearchForMysql
		end if
		if bValue(IsCharType(var_type)) and not bValue(btexttype) then
			StrWhereExpression = (((CSmartStr(strField) & " ") & CSmartStr(var_like)) & " ") & CSmartStr(isEnableUpper(("'%" & CSmartStr(sSearchFor)) & "%'"))
			Exit Function
		else
			StrWhereExpression = ((((CSmartStr(strField) & " ") & CSmartStr(var_like)) & " '%") & CSmartStr(sSearchFor)) & "%'"
			Exit Function
		end if
	else
		if IsEqual(strSearchOption,"Equals") then
			StrWhereExpression = (CSmartStr(strField) & "=") & CSmartStr(sSearchFor)
			Exit Function
		else
			if IsEqual(strSearchOption,"Starts with") then
				if bValue(isMysql) then
					doAssignment sSearchFor,sSearchForMysql
				end if
				if bValue(IsCharType(var_type)) and not bValue(btexttype) then
					StrWhereExpression = (((CSmartStr(strField) & " ") & CSmartStr(var_like)) & " ") & CSmartStr(isEnableUpper(("'" & CSmartStr(sSearchFor)) & "%'"))
					Exit Function
				else
					StrWhereExpression = ((((CSmartStr(strField) & " ") & CSmartStr(var_like)) & " '") & CSmartStr(sSearchFor)) & "%'"
					Exit Function
				end if
			else
				if IsEqual(strSearchOption,"More than") then
					StrWhereExpression = (CSmartStr(strField) & ">") & CSmartStr(sSearchFor)
					Exit Function
				else
					if IsEqual(strSearchOption,"Less than") then
						StrWhereExpression = (CSmartStr(strField) & "<") & CSmartStr(sSearchFor)
						Exit Function
					else
						if IsEqual(strSearchOption,"Between") then
							ret = (CSmartStr(strField) & ">=") & CSmartStr(sSearchFor)
							if bValue(sSearchFor2) then
								ret = CSmartStr(ret) & (((" and " & CSmartStr(strField)) & "<=") & CSmartStr(sSearchFor2))
							end if
							doAssignmentByRef StrWhereExpression,ret
							Exit Function
						end if
					end if
				end if
			end if
		end if
	end if
	StrWhereExpression = ""
	Exit Function
End Function
Function StrWhereAdv(ByVal strField,ByVal SearchFor,ByVal strSearchOption,ByVal SearchFor2,ByVal etype,ByVal isSuggest)
	Dim var_type,isOracle,ismssql,isdb2,btexttype,isMysql,var_like,ret,value,whereStr,value1,value2,cleanvalue2,gstrField,timeArr
	doAssignmentByRef var_type,GetFieldType(strField,"")
	isOracle = false
	ismssql = true
	isdb2 = false
	doAssignmentByRef btexttype,IsTextType(var_type)
	isMysql = false
	if bValue(IsBinaryType(var_type)) then
		StrWhereAdv = ""
		Exit Function
	end if
	if (bValue(btexttype) and not IsEqual(strSearchOption,"Contains")) and not IsEqual(strSearchOption,"Starts with") then
		StrWhereAdv = ""
		Exit Function
	end if
	if IsEqual(strSearchOption,"Empty") then
		if (bValue(IsCharType(var_type)) and (not bValue(ismssql) or not bValue(btexttype))) and not bValue(isOracle) then
			StrWhereAdv = ((("(" & CSmartStr(GetFullFieldName(strField,""))) & " is null or ") & CSmartStr(GetFullFieldName(strField,""))) & "='')"
			Exit Function
		else
			if bValue(ismssql) and bValue(btexttype) then
				StrWhereAdv = ((("(" & CSmartStr(GetFullFieldName(strField,""))) & " is null or ") & CSmartStr(GetFullFieldName(strField,""))) & " LIKE '')"
				Exit Function
			else
				StrWhereAdv = CSmartStr(GetFullFieldName(strField,"")) & " is null"
				Exit Function
			end if
		end if
	end if
	var_like = "like"
	if IsEqual(GetEditFormat(strField,""),EDIT_FORMAT_LOOKUP_WIZARD) then
		if bValue(Multiselect(strField,"")) then
			doAssignmentByRef SearchFor,splitvalues(SearchFor)
		else
			Set SearchFor = (CreateDictionary1(Empty,SearchFor))
		end if
		ret = ""
		GetCollectionBounds SearchFor,i_commonfunctions_loopIdx23,i_commonfunctions_loopMax23
		do while i_commonfunctions_loopIdx23<=i_commonfunctions_loopMax23
			i_commonfunctions_arrKey23 = GetCollectionKey(SearchFor,i_commonfunctions_loopIdx23)
			doAssignment value,ArrayElement(SearchFor,i_commonfunctions_arrKey23)
			if not ((IsEqual(value,"null") or IsEqual(value,"Null")) or IsEqual(value,"")) then
				if bValue(asp_strlen(ret)) then
					ret = CSmartStr(ret) & " or "
				end if
				if IsEqual(strSearchOption,"Equals") then
					doAssignmentByRef value,make_db_value(strField,value,"","","")
					if not (IsEqual(value,"null") or IsEqual(value,"Null")) then
						ret = CSmartStr(ret) & ((CSmartStr(GetFullFieldName(strField,"")) & "=") & CSmartStr(value))
					end if
				else
					if bValue(isSuggest) then
						ret = CSmartStr(ret) & ((((((" " & CSmartStr(GetFullFieldName(strField,""))) & " ") & CSmartStr(var_like)) & " '%") & CSmartStr(value)) & "%'")
					else
						if not IsFalse(asp_strpos(value,",",empty)) or not IsFalse(asp_strpos(value,"""",empty)) then
							value = ("""" & CSmartStr(asp_str_replace("""","""""",value))) & """"
						end if
						doAssignmentByRef value,db_addslashes(value)
						if bValue(isMysql) then
							doAssignmentByRef value,asp_str_replace("\\","\\\\",value)
						end if
						ret = CSmartStr(ret) & (((CSmartStr(GetFullFieldName(strField,"")) & " = '") & CSmartStr(value)) & "'")
						ret = CSmartStr(ret) & ((((((" or " & CSmartStr(GetFullFieldName(strField,""))) & " ") & CSmartStr(var_like)) & " '%,") & CSmartStr(value)) & ",%'")
						ret = CSmartStr(ret) & ((((((" or " & CSmartStr(GetFullFieldName(strField,""))) & " ") & CSmartStr(var_like)) & " '%,") & CSmartStr(value)) & "'")
						ret = CSmartStr(ret) & ((((((" or " & CSmartStr(GetFullFieldName(strField,""))) & " ") & CSmartStr(var_like)) & " '") & CSmartStr(value)) & ",%'")
					end if
				end if
			end if
			i_commonfunctions_loopIdx23=i_commonfunctions_loopIdx23+1
		loop
		if bValue(asp_strlen(ret)) then
			ret = ("(" & CSmartStr(ret)) & ")"
		end if
		doAssignmentByRef StrWhereAdv,ret
		Exit Function
	end if
	if IsEqual(GetEditFormat(strField,""),EDIT_FORMAT_CHECKBOX) then
		if IsEqual(SearchFor,"none") then
			StrWhereAdv = ""
			Exit Function
		end if
		if bValue(NeedQuotes(var_type)) then
			isOracle = false
			if IsEqual(SearchFor,"on") then
				whereStr = ("(" & CSmartStr(GetFullFieldName(strField,""))) & "<>'0' "
				if not bValue(isOracle) then
					whereStr = CSmartStr(whereStr) & ((" and " & CSmartStr(GetFullFieldName(strField,""))) & "<>'' ")
				end if
				whereStr = CSmartStr(whereStr) & ((" and " & CSmartStr(GetFullFieldName(strField,""))) & " is not null)")
				doAssignmentByRef StrWhereAdv,whereStr
				Exit Function
			else
				if IsEqual(SearchFor,"off") then
					whereStr = ("(" & CSmartStr(GetFullFieldName(strField,""))) & "='0' "
					if not bValue(isOracle) then
						whereStr = CSmartStr(whereStr) & ((" or " & CSmartStr(GetFullFieldName(strField,""))) & "='' ")
					end if
					whereStr = CSmartStr(whereStr) & ((" or " & CSmartStr(GetFullFieldName(strField,""))) & " is null)")
				end if
			end if
		else
			if IsEqual(SearchFor,"on") then
				StrWhereAdv = ((("(" & CSmartStr(GetFullFieldName(strField,""))) & "<>0 and ") & CSmartStr(GetFullFieldName(strField,""))) & " is not null)"
				Exit Function
			else
				if IsEqual(SearchFor,"off") then
					StrWhereAdv = ((("(" & CSmartStr(GetFullFieldName(strField,""))) & "=0 or ") & CSmartStr(GetFullFieldName(strField,""))) & " is null)"
					Exit Function
				end if
			end if
		end if
	end if
	doAssignmentByRef value1,make_db_value(strField,SearchFor,etype,"","")
	value2 = false
	cleanvalue2 = false
	if IsEqual(strSearchOption,"Between") then
		doAssignmentByRef cleanvalue2,prepare_for_db(strField,SearchFor2,etype,"","")
		doAssignmentByRef value2,make_db_value(strField,SearchFor2,etype,"","")
	end if
	if (not IsEqual(strSearchOption,"Contains") and not IsEqual(strSearchOption,"Starts with")) and (IsIdentical(value1,"null") or IsIdentical(value2,"null")) then
		StrWhereAdv = ""
		Exit Function
	end if
	if bValue(IsCharType(var_type)) and not bValue(btexttype) then
		doAssignmentByRef value1,isEnableUpper(value1)
		doAssignmentByRef value2,isEnableUpper(value2)
		doAssignmentByRef gstrField,isEnableUpper(GetFullFieldName(strField,""))
	else
		if IsEqual(strSearchOption,"Contains") or IsEqual(strSearchOption,"Starts with") then
			doAssignmentByRef gstrField,db_field2char(GetFullFieldName(strField,""),var_type)
		else
			if IsEqual(ViewFormat(strField,""),FORMAT_TIME) then
				doAssignmentByRef gstrField,db_field2time(GetFullFieldName(strField,""),var_type)
			else
				doAssignmentByRef gstrField,GetFullFieldName(strField,"")
			end if
		end if
	end if
	ret = ""
	if IsEqual(strSearchOption,"Contains") then
		doAssignmentByRef SearchFor,db_addslashes(SearchFor)
		if bValue(isMysql) then
			doAssignmentByRef SearchFor,asp_str_replace("\\","\\\\",SearchFor)
		end if
		if bValue(IsCharType(var_type)) and not bValue(btexttype) then
			StrWhereAdv = (((CSmartStr(gstrField) & " ") & CSmartStr(var_like)) & " ") & CSmartStr(isEnableUpper(("'%" & CSmartStr(SearchFor)) & "%'"))
			Exit Function
		else
			StrWhereAdv = ((((CSmartStr(gstrField) & " ") & CSmartStr(var_like)) & " '%") & CSmartStr(SearchFor)) & "%'"
			Exit Function
		end if
	else
		if IsEqual(strSearchOption,"Equals") then
			StrWhereAdv = (CSmartStr(gstrField) & "=") & CSmartStr(value1)
			Exit Function
		else
			if IsEqual(strSearchOption,"Starts with") then
				doAssignmentByRef SearchFor,db_addslashes(SearchFor)
				if bValue(isMysql) then
					doAssignmentByRef SearchFor,asp_str_replace("\\","\\\\",SearchFor)
				end if
				if bValue(IsCharType(var_type)) and not bValue(btexttype) then
					StrWhereAdv = (((CSmartStr(gstrField) & " ") & CSmartStr(var_like)) & " ") & CSmartStr(isEnableUpper(("'" & CSmartStr(SearchFor)) & "%'"))
					Exit Function
				else
					StrWhereAdv = ((((CSmartStr(gstrField) & " ") & CSmartStr(var_like)) & " '") & CSmartStr(SearchFor)) & "%'"
					Exit Function
				end if
			else
				if IsEqual(strSearchOption,"More than") then
					StrWhereAdv = (CSmartStr(gstrField) & ">") & CSmartStr(value1)
					Exit Function
				else
					if IsEqual(strSearchOption,"Less than") then
						StrWhereAdv = (CSmartStr(gstrField) & "<") & CSmartStr(value1)
						Exit Function
					else
						if IsEqual(strSearchOption,"Equal or more than") then
							StrWhereAdv = (CSmartStr(gstrField) & ">=") & CSmartStr(value1)
							Exit Function
						else
							if IsEqual(strSearchOption,"Equal or less than") then
								StrWhereAdv = (CSmartStr(gstrField) & "<=") & CSmartStr(value1)
								Exit Function
							else
								if IsEqual(strSearchOption,"Between") then
									ret = ((CSmartStr(gstrField) & ">=") & CSmartStr(value1)) & " and "
									if bValue(IsDateFieldType(var_type)) then
										doAssignmentByRef timeArr,db2time(cleanvalue2)
										if (IsEqual(ArrayElement(timeArr,3),0) and IsEqual(ArrayElement(timeArr,4),0)) and IsEqual(ArrayElement(timeArr,5),0) then
											doAssignmentByRef timeArr,adddays(timeArr,1)
											value2 = (((CSmartStr(ArrayElement(timeArr,0)) & "-") & CSmartStr(ArrayElement(timeArr,1))) & "-") & CSmartStr(ArrayElement(timeArr,2))
											doAssignmentByRef value2,add_db_quotes(strField,value2,strTableName)
											ret = CSmartStr(ret) & ((CSmartStr(gstrField) & "<") & CSmartStr(value2))
										else
											ret = CSmartStr(ret) & ((CSmartStr(gstrField) & "<=") & CSmartStr(value2))
										end if
									else
										ret = CSmartStr(ret) & ((CSmartStr(gstrField) & "<=") & CSmartStr(value2))
									end if
									doAssignmentByRef StrWhereAdv,ret
									Exit Function
								end if
							end if
						end if
					end if
				end if
			end if
		end if
	end if
	StrWhereAdv = ""
	Exit Function
End Function
Function gSQLRowCount(ByVal where,ByVal having)
	Dim sqlHead,sqlGroupBy,oHaving,sqlHaving
	doAssignmentByRef sqlHead,gQuery.HeadToSql()
	doAssignmentByRef sqlGroupBy,gQuery.GroupByToSql()
	doAssignmentByRef oHaving,gQuery.Having()
	doAssignmentByRef sqlHaving,oHaving.toSql_p1(gQuery)
	doAssignmentByRef gSQLRowCount,gSQLRowCount_int(sqlHead,gsqlFrom,gsqlWhereExpr,sqlGroupBy,sqlHaving,where,having)
	Exit Function
End Function
Function gSQLRowCount_int(ByVal sqlHead,ByVal sqlFrom,ByVal sqlWhere,ByVal sqlGroupBy,ByVal sqlHaving,ByVal where,ByVal having)
	Dim strWhere,countstr,countrs,countdata
	doAssignmentByRef strWhere,whereAdd(sqlWhere,where)
	if bValue(asp_strlen(strWhere)) then
		strWhere = (" where " & CSmartStr(strWhere)) & " "
	end if
	if bValue(asp_strlen(sqlGroupBy)) then
		countstr = ("select count(*) from (" & CSmartStr(gSQLWhere_having(sqlHead,sqlFrom,sqlWhere,sqlGroupBy,sqlHaving,where,having))) & ") a"
	else
		countstr = ("select count(*) " & CSmartStr(sqlFrom)) & CSmartStr(strWhere)
	end if
	doAssignmentByRef countrs,db_query(countstr,conn)
	doAssignmentByRef countdata,db_fetch_numarray(countrs)
	doAssignmentByRef gSQLRowCount_int,ArrayElement(countdata,0)
	Exit Function
End Function
Function GetRowCount(ByVal strSQL)
	Dim tstr,ind1,ind2,ind3,countstr,countrs,countdata
	doAssignmentByRef strSQL,asp_str_replace(CreateDictionary3(Empty,vbcrlf,Empty,vblf,Empty,"	")," ",strSQL)
	doAssignmentByRef tstr,asp_strtoupper(strSQL)
	doAssignmentByRef ind1,asp_strpos(tstr,"SELECT ",empty)
	doAssignmentByRef ind2,my_strrpos(tstr," FROM ")
	doAssignmentByRef ind3,my_strrpos(tstr," GROUP BY ")
	if IsFalse(ind3) then
		doAssignmentByRef ind3,asp_strpos(tstr," ORDER BY ",empty)
		if IsFalse(ind3) then
			doAssignmentByRef ind3,asp_strlen(strSQL)
		end if
	end if
	countstr = (CSmartStr(asp_substr(strSQL,0,CSmartDbl(ind1)+6)) & " count(*) ") & CSmartStr(asp_substr(strSQL,CSmartDbl(ind2)+1,CSmartDbl(ind3)-CSmartDbl(ind2)))
	doAssignmentByRef countrs,db_query(countstr,conn)
	doAssignmentByRef countdata,db_fetch_numarray(countrs)
	doAssignmentByRef GetRowCount,ArrayElement(countdata,0)
	Exit Function
End Function
Function AddTop(ByVal strSQL,ByVal n)
	Dim tstr,ind1
	doAssignmentByRef tstr,asp_strtoupper(strSQL)
	doAssignmentByRef ind1,asp_strpos(tstr,"SELECT",empty)
	AddTop = (((CSmartStr(asp_substr(strSQL,0,CSmartDbl(ind1)+6)) & " top ") & CSmartStr(n)) & " ") & CSmartStr(asp_substr(strSQL,CSmartDbl(ind1)+6,empty))
	Exit Function
End Function
Function AddTopDB2(ByVal strSQL,ByVal n)
	AddTopDB2 = ((CSmartStr(strSQL) & " fetch first ") & CSmartStr(n)) & " rows only"
	Exit Function
End Function
Function AddTopIfx(ByVal strSQL,ByVal n)
	AddTopIfx = (((CSmartStr(asp_substr(strSQL,0,7)) & "limit ") & CSmartStr(n)) & " ") & CSmartStr(asp_substr(strSQL,7,empty))
	Exit Function
End Function
Function AddRowNumber(ByVal strSQL,ByVal n)
	AddRowNumber = (("select * from (" & CSmartStr(strSQL)) & ") where rownum<") & CSmartStr(CSmartDbl(n)+1)
	Exit Function
End Function
Function NeedQuotesNumeric(ByVal var_type)
	if ((((((((((((IsEqual(var_type,203) or IsEqual(var_type,8)) or IsEqual(var_type,129)) or IsEqual(var_type,130)) or IsEqual(var_type,7)) or IsEqual(var_type,133)) or IsEqual(var_type,134)) or IsEqual(var_type,135)) or IsEqual(var_type,201)) or IsEqual(var_type,205)) or IsEqual(var_type,200)) or IsEqual(var_type,202)) or IsEqual(var_type,72)) or IsEqual(var_type,13) then
		NeedQuotesNumeric = true
		Exit Function
	else
		NeedQuotesNumeric = false
		Exit Function
	end if
End Function
Function IsNumberType(ByVal var_type)
	if ((((((((((((((IsEqual(var_type,20) or IsEqual(var_type,6)) or IsEqual(var_type,14)) or IsEqual(var_type,5)) or IsEqual(var_type,10)) or IsEqual(var_type,3)) or IsEqual(var_type,131)) or IsEqual(var_type,4)) or IsEqual(var_type,2)) or IsEqual(var_type,16)) or IsEqual(var_type,21)) or IsEqual(var_type,19)) or IsEqual(var_type,18)) or IsEqual(var_type,17)) or IsEqual(var_type,139)) or IsEqual(var_type,11) then
		IsNumberType = true
		Exit Function
	end if
	IsNumberType = false
	Exit Function
End Function
Function IsFloatType(ByVal var_type)
	if ((IsEqual(var_type,14) or IsEqual(var_type,5)) or IsEqual(var_type,131)) or IsEqual(var_type,6) then
		IsFloatType = true
		Exit Function
	end if
	IsFloatType = false
	Exit Function
End Function
Function NeedQuotes(ByVal var_type)
	NeedQuotes = not bValue(IsNumberType(var_type))
	Exit Function
End Function
Function IsBinaryType(ByVal var_type)
	if (IsEqual(var_type,128) or IsEqual(var_type,205)) or IsEqual(var_type,204) then
		IsBinaryType = true
		Exit Function
	end if
	IsBinaryType = false
	Exit Function
End Function
Function IsDateFieldType(ByVal var_type)
	if (IsEqual(var_type,7) or IsEqual(var_type,133)) or IsEqual(var_type,135) then
		IsDateFieldType = true
		Exit Function
	end if
	IsDateFieldType = false
	Exit Function
End Function
Function IsTimeType(ByVal var_type)
	if IsEqual(var_type,134) then
		IsTimeType = true
		Exit Function
	end if
	IsTimeType = false
	Exit Function
End Function
Function IsCharType(ByVal var_type)
	if ((((bValue(IsTextType(var_type)) or IsEqual(var_type,8)) or IsEqual(var_type,129)) or IsEqual(var_type,200)) or IsEqual(var_type,202)) or IsEqual(var_type,130) then
		IsCharType = true
		Exit Function
	end if
	IsCharType = false
	Exit Function
End Function
Function IsTextType(ByVal var_type)
	if IsEqual(var_type,201) or IsEqual(var_type,203) then
		IsTextType = true
		Exit Function
	end if
	IsTextType = false
	Exit Function
End Function
Function IsGuid(ByVal var_type)
	if IsEqual(var_type,72) then
		IsGuid = true
		Exit Function
	end if
	IsGuid = false
	Exit Function
End Function
Function IsAdmin()
	IsAdmin = false
	Exit Function
End Function
Function GetUserPermissions(ByVal table)
	GetUserPermissions = "ADESPIM"
	Exit Function
End Function
Function CheckFieldPermissions(ByVal field,ByVal table)
	doAssignmentByRef CheckFieldPermissions,GetFieldData(table,field,"FieldPermissions",false)
	Exit Function
End Function
Function CheckSecurity(ByVal strValue,ByVal strAction)
	Dim var_SESSION,strPerm
	CheckSecurity = true
	Exit Function
End Function
Function SecuritySQL(ByVal strAction,ByVal table)
	Dim ownerid,var_SESSION,ret,strPerm
	if not bValue(asp_strlen(table)) then
		doAssignment table,strTableName
	end if
	SecuritySQL = ""
	Exit Function
End Function
Function make_db_value(ByVal field,ByVal value,ByVal controltype,ByVal postfilename,ByVal table)
	Dim ret
	doAssignmentByRef ret,prepare_for_db(field,value,controltype,postfilename,table)
	if IsFalse(ret) then
		doAssignmentByRef make_db_value,ret
		Exit Function
	end if
	doAssignmentByRef make_db_value,add_db_quotes(field,ret,table)
	Exit Function
End Function
Function add_db_quotes(ByVal field,ByVal value,ByVal table)
	Dim var_type,strvalue
	doAssignmentByRef var_type,GetFieldType(field,table)
	if bValue(IsBinaryType(var_type)) then
		doAssignmentByRef add_db_quotes,db_addslashesbinary(value)
		Exit Function
	end if
	if (IsIdentical(value,"") or IsFalse(value)) and not bValue(IsCharType(var_type)) then
		add_db_quotes = "null"
		Exit Function
	end if
	if bValue(NeedQuotes(var_type)) then
		if not bValue(IsDateFieldType(var_type)) then
			value = ("'" & CSmartStr(db_addslashes(value))) & "'"
		else
			doAssignmentByRef value,db_datequotes(value)
		end if
	else
		strvalue = CSmartStr(value)
		doAssignmentByRef strvalue,asp_str_replace(",",".",strvalue)
		if bValue(asp_is_numeric(strvalue)) then
			doAssignment value,strvalue
		else
			value = 0
		end if
	end if
	doAssignmentByRef add_db_quotes,value
	Exit Function
End Function
Function prepare_for_db(ByVal field,ByVal value,ByVal controltype,ByVal postfilename,ByVal table)
	Dim filename,var_type,time,dformat,a,y,m,d,ret
	filename = ""
	doAssignmentByRef var_type,GetFieldType(field,table)
	if not bValue(controltype) or IsEqual(controltype,"multiselect") then
		if bValue(asp_is_array(value)) then
			doAssignmentByRef value,combinevalues(value)
		end if
		if (IsIdentical(value,"") or IsFalse(value)) and not bValue(IsCharType(var_type)) then
			prepare_for_db = ""
			Exit Function
		end if
		if bValue(IsGuid(var_type)) then
			if not bValue(IsGuidString(value)) then
				prepare_for_db = ""
				Exit Function
			end if
		end if
		doAssignmentByRef prepare_for_db,value
		Exit Function
	else
		if IsEqual(controltype,"time") then
			if not bValue(asp_strlen(value)) then
				prepare_for_db = ""
				Exit Function
			end if
			doAssignmentByRef time,localtime2db(value)
			if bValue(IsDateFieldType(GetFieldType(field,table))) then
				time = "2000-01-01 " & CSmartStr(time)
			end if
			doAssignmentByRef prepare_for_db,time
			Exit Function
		else
			if IsEqual(asp_substr(controltype,0,4),"date") then
				doAssignmentByRef dformat,asp_substr(controltype,4,empty)
				if IsEqual(dformat,EDIT_DATE_SIMPLE) or IsEqual(dformat,EDIT_DATE_SIMPLE_DP) then
					doAssignmentByRef time,localdatetime2db(value,"")
					if IsEqual(time,"null") then
						prepare_for_db = ""
						Exit Function
					end if
					doAssignmentByRef prepare_for_db,time
					Exit Function
				else
					if IsEqual(dformat,EDIT_DATE_DD) or IsEqual(dformat,EDIT_DATE_DD_DP) then
						doAssignmentByRef a,explode("-",value)
						if IsLess(asp_count(a),3) then
							prepare_for_db = ""
							Exit Function
						else
							doAssignment y,ArrayElement(a,0)
							doAssignment m,ArrayElement(a,1)
							doAssignment d,ArrayElement(a,2)
						end if
						if IsLess(y,100) then
							if IsLess(y,70) then
								y = CSmartDbl(y)+2000
							else
								y = CSmartDbl(y)+1900
							end if
						end if
						doAssignmentByRef prepare_for_db,mysprintf("%04d-%02d-%02d",CreateDictionary3(Empty,y,Empty,m,Empty,d))
						Exit Function
					else
						prepare_for_db = ""
						Exit Function
					end if
				end if
			else
				if IsEqual(asp_substr(controltype,0,8),"checkbox") then
					if IsEqual(value,"on") then
						ret = 1
					else
						if IsEqual(value,"none") then
							prepare_for_db = ""
							Exit Function
						else
							ret = 0
						end if
					end if
					doAssignmentByRef prepare_for_db,ret
					Exit Function
				else
					prepare_for_db = false
					Exit Function
				end if
			end if
		end if
	end if
End Function
Function DeleteUploadedFiles(ByVal where,ByVal table)
	Dim sql,rs,data,value,field,isAbs,filename
	doAssignmentByRef sql,gSQLWhere(where,"")
	doAssignmentByRef rs,db_query(sql,conn)
	if not bValue(doAssignmentByRef(data,db_fetch_array(rs))) then
		Exit Function
	end if
	GetCollectionBounds data,i_commonfunctions_loopIdx28,i_commonfunctions_loopMax28
	do while i_commonfunctions_loopIdx28<=i_commonfunctions_loopMax28
		field = GetCollectionKey(data,i_commonfunctions_loopIdx28)
		doAssignment value,ArrayElement(data,field)
		if bValue(asp_strlen(value)) and IsEqual(GetEditFormat(field,""),EDIT_FORMAT_FILE) then
			doAssignmentByRef isAbs,GetFieldData(table,field,"Absolute",false)
			filename = CSmartStr(GetUploadFolder(field,"")) & CSmartStr(value)
			if not bValue(isAbs) then
				doAssignmentByRef filename,getabspath(filename)
			end if
			runner_delete_file filename
			if bValue(GetCreateThumbnail(field,"")) then
				filename = (CSmartStr(GetUploadFolder(field,"")) & CSmartStr(GetThumbnailPrefix(field,""))) & CSmartStr(value)
				if not bValue(isAbs) then
					doAssignmentByRef filename,getabspath(filename)
				end if
				runner_delete_file filename
			end if
		end if
		i_commonfunctions_loopIdx28=i_commonfunctions_loopIdx28+1
	loop
End Function
Function combinevalues(ByVal arr)
	Dim ret,val
	ret = ""
	GetCollectionBounds arr,i_commonfunctions_loopIdx29,i_commonfunctions_loopMax29
	do while i_commonfunctions_loopIdx29<=i_commonfunctions_loopMax29
		i_commonfunctions_arrKey29 = GetCollectionKey(arr,i_commonfunctions_loopIdx29)
		doAssignment val,ArrayElement(arr,i_commonfunctions_arrKey29)
		if bValue(asp_strlen(ret)) then
			ret = CSmartStr(ret) & ","
		end if
		if IsFalse(asp_strpos(val,",",empty)) and IsFalse(asp_strpos(val,"""",empty)) then
			ret = CSmartStr(ret) & CSmartStr(val)
		else
			doAssignmentByRef val,asp_str_replace("""","""""",val)
			ret = CSmartStr(ret) & (("""" & CSmartStr(val)) & """")
		end if
		i_commonfunctions_loopIdx29=i_commonfunctions_loopIdx29+1
	loop
	doAssignmentByRef combinevalues,ret
	Exit Function
End Function
Function splitvalues(ByVal str)
	Dim arr,start,i,inquot,val
	Set arr = (CreateDictionary())
	start = 0
	i = 0
	inquot = false
	do while IsLessOrEqual(i,asp_strlen(str))
		if IsLess(i,asp_strlen(str)) and IsEqual(asp_substr(str,i,1),"""") then
			inquot = not bValue(inquot)
		else
			if IsEqual(i,asp_strlen(str)) or not bValue(inquot) and IsEqual(asp_substr(str,i,1),",") then
				doAssignmentByRef val,asp_substr(str,start,CSmartDbl(i)-CSmartDbl(start))
				start = CSmartDbl(i)+1
				if bValue(asp_strlen(val)) and IsEqual(asp_substr(val,0,1),"""") then
					doAssignmentByRef val,asp_substr(val,1,CSmartDbl(asp_strlen(val))-2)
					doAssignmentByRef val,asp_str_replace("""""","""",val)
				end if
				setArrElement arr,asp_count(arr),val
			end if
		end if
		i = CSmartDbl(i)+1
	loop
	doAssignmentByRef splitvalues,arr
	Exit Function
End Function
Function GetDateEdit(ByVal field,ByVal value,ByVal var_type,ByVal fieldNum,ByVal search,ByVal record_id,ByRef pageObj)
	Dim is508,strLabel,cfieldname,cfield,tvalue,time,dp,ovalue,fmt,sundayfirst,ovalue1,showtime,ret,retday,retmonth,retyear
	doAssignmentByRef is508,isEnableSection508()
	doAssignmentByRef strLabel,Label(field,"")
	doAssignmentByRef cfieldname,GoodFieldName(field)
	cfield = (("value_" & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(record_id)
	if bValue(fieldNum) then
		cfield = (((("value" & CSmartStr(fieldNum)) & "_") & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(record_id)
	end if
	doAssignment tvalue,value
	doAssignmentByRef time,db2time(tvalue)
	if not bValue(asp_count(time)) then
		Set time = (CreateDictionary6(Empty,0,Empty,0,Empty,0,Empty,0,Empty,0,Empty,0))
	end if
	dp = 0
	do
		If IsEqual(var_type,EDIT_DATE_SIMPLE_DP) then
			doAssignment ovalue,value
			if IsEqual(ArrayElement(locale_info,"LOCALE_IDATE"),1) then
				fmt = ((("dd" & CSmartStr(ArrayElement(locale_info,"LOCALE_SDATE"))) & "MM") & CSmartStr(ArrayElement(locale_info,"LOCALE_SDATE"))) & "yyyy"
				sundayfirst = "false"
			else
				if IsEqual(ArrayElement(locale_info,"LOCALE_IDATE"),0) then
					fmt = ((("MM" & CSmartStr(ArrayElement(locale_info,"LOCALE_SDATE"))) & "dd") & CSmartStr(ArrayElement(locale_info,"LOCALE_SDATE"))) & "yyyy"
					sundayfirst = "true"
				else
					fmt = ((("yyyy" & CSmartStr(ArrayElement(locale_info,"LOCALE_SDATE"))) & "MM") & CSmartStr(ArrayElement(locale_info,"LOCALE_SDATE"))) & "dd"
					sundayfirst = "false"
				end if
			end if
			if bValue(ArrayElement(time,5)) then
				fmt = CSmartStr(fmt) & " HH:mm:ss"
			else
				if bValue(ArrayElement(time,3)) or bValue(ArrayElement(time,4)) then
					fmt = CSmartStr(fmt) & " HH:mm"
				end if
			end if
			if bValue(ArrayElement(time,0)) then
				doAssignmentByRef ovalue,format_datetime_custom(time,fmt)
			end if
			ovalue1 = (((CSmartStr(ArrayElement(time,2)) & "-") & CSmartStr(ArrayElement(time,1))) & "-") & CSmartStr(ArrayElement(time,0))
			showtime = "false"
			if bValue(DateEditShowTime(field,"")) then
				showtime = "true"
				ovalue1 = CSmartStr(ovalue1) & (((((" " & CSmartStr(ArrayElement(time,3))) & ":") & CSmartStr(ArrayElement(time,4))) & ":") & CSmartStr(ArrayElement(time,5)))
			end if
			ret = ((((("<input id=""" & CSmartStr(cfield)) & """ type=""Text"" name=""") & CSmartStr(cfield)) & """ size=""20"" value=""") & CSmartStr(ovalue)) & """>"
			ret = CSmartStr(ret) & (((((("<input id=""ts" & CSmartStr(cfield)) & """ type=""Hidden"" name=""ts") & CSmartStr(cfield)) & """ value=""") & CSmartStr(ovalue1)) & """>&nbsp;&nbsp;")
			ret = CSmartStr(ret) & ((((("&nbsp;<a href=""#"" id=""imgCal_" & CSmartStr(cfield)) & """>") & "<img src=""images/cal.gif"" width=16 height=16 border=0 alt=""") & "Click Here to Pick up the date") & """></a>")
			ResponseWrite ret
			Exit Function
		End If
		If IsEqual(var_type,EDIT_DATE_SIMPLE_DP) or IsEqual(var_type,EDIT_DATE_DD_DP) then
			dp = 1
		End If
		If IsEqual(var_type,EDIT_DATE_SIMPLE_DP) or IsEqual(var_type,EDIT_DATE_DD_DP) or IsEqual(var_type,EDIT_DATE_DD) then
			retday = ((((("<select id=""day" & CSmartStr(cfield)) & """ ") & CSmartStr(IIF((IsEqual(search,MODE_INLINE_EDIT) or IsEqual(search,MODE_INLINE_ADD)) and IsEqual(is508,true),("alt=""" & CSmartStr(strLabel)) & """ ",""))) & "name=""day") & CSmartStr(cfield)) & """ ></select>"
			retmonth = ((((("<select id=""month" & CSmartStr(cfield)) & """ ") & CSmartStr(IIF((IsEqual(search,MODE_INLINE_EDIT) or IsEqual(search,MODE_INLINE_ADD)) and IsEqual(is508,true),("alt=""" & CSmartStr(strLabel)) & """ ",""))) & "name=""month") & CSmartStr(cfield)) & """ ></select>"
			retyear = ((((("<select id=""year" & CSmartStr(cfield)) & """ ") & CSmartStr(IIF((IsEqual(search,MODE_INLINE_EDIT) or IsEqual(search,MODE_INLINE_ADD)) and IsEqual(is508,true),("alt=""" & CSmartStr(strLabel)) & """ ",""))) & "name=""year") & CSmartStr(cfield)) & """ ></select>"
			sundayfirst = "false"
			if IsEqual(ArrayElement(locale_info,"LOCALE_ILONGDATE"),1) then
				ret = (((CSmartStr(retday) & "&nbsp;") & CSmartStr(retmonth)) & "&nbsp;") & CSmartStr(retyear)
			else
				if IsEqual(ArrayElement(locale_info,"LOCALE_ILONGDATE"),0) then
					ret = (((CSmartStr(retmonth) & "&nbsp;") & CSmartStr(retday)) & "&nbsp;") & CSmartStr(retyear)
					sundayfirst = "true"
				else
					ret = (((CSmartStr(retyear) & "&nbsp;") & CSmartStr(retmonth)) & "&nbsp;") & CSmartStr(retday)
				end if
			end if
			if (bValue(ArrayElement(time,0)) and bValue(ArrayElement(time,1))) and bValue(ArrayElement(time,2)) then
				ret = CSmartStr(ret) & (((((((((("<input id=""" & CSmartStr(cfield)) & """ type=hidden name=""") & CSmartStr(cfield)) & """ value=""") & CSmartStr(ArrayElement(time,0))) & "-") & CSmartStr(ArrayElement(time,1))) & "-") & CSmartStr(ArrayElement(time,2))) & """>")
			else
				ret = CSmartStr(ret) & (((("<input id=""" & CSmartStr(cfield)) & """ type=hidden name=""") & CSmartStr(cfield)) & """ value="""">")
			end if
			if bValue(dp) then
				ret = CSmartStr(ret) & (((((((((((((("&nbsp;<a href=""#"" id=""imgCal_" & CSmartStr(cfield)) & """>") & "<img src=""images/cal.gif"" width=16 height=16 border=0 alt=""Click Here to Pick up the date""></a>") & "<input id=""ts") & CSmartStr(cfield)) & """ type=hidden name=""ts") & CSmartStr(cfield)) & """ value=""") & CSmartStr(ArrayElement(time,2))) & "-") & CSmartStr(ArrayElement(time,1))) & "-") & CSmartStr(ArrayElement(time,0))) & """>")
			end if
			ResponseWrite ret
			Exit Function
		End If
		doAssignment ovalue,value
		if bValue(ArrayElement(time,0)) then
			if (bValue(ArrayElement(time,3)) or bValue(ArrayElement(time,4))) or bValue(ArrayElement(time,5)) then
				doAssignmentByRef ovalue,str_format_datetime(time)
			else
				doAssignmentByRef ovalue,format_shortdate(time)
			end if
		end if
		ResponseWrite ((((("<input id=""" & CSmartStr(cfield)) & """ type=text name=""") & CSmartStr(cfield)) & """ size=""20"" value=""") & CSmartStr(htmlspecialchars(ovalue))) & """>"
		Exit Function
	Loop While false
End Function
Function BuildSecondDropdownArray(ByVal arrName,ByVal strSQL)
	Dim i,rs,row
	ResponseWrite CSmartStr(arrName) & "=new Array();" & vbcrlf
	i = 0
	doAssignmentByRef rs,db_query(strSQL,conn)
	do while bValue(doAssignmentByRef(row,db_fetch_numarray(rs)))
		ResponseWrite ((((CSmartStr(arrName) & "[") & CSmartStr(CSmartDbl(i)*3)) & "]='") & CSmartStr(jsreplace(ArrayElement(row,0)))) & "';" & vbcrlf
		ResponseWrite ((((CSmartStr(arrName) & "[") & CSmartStr(CSmartDbl(i)*3+1)) & "]='") & CSmartStr(jsreplace(ArrayElement(row,1)))) & "';" & vbcrlf
		ResponseWrite ((((CSmartStr(arrName) & "[") & CSmartStr(CSmartDbl(i)*3+2)) & "]='") & CSmartStr(jsreplace(ArrayElement(row,2)))) & "';" & vbcrlf
		i = CSmartDbl(i)+1
	loop
End Function
Function BuildSelectControl(ByVal field,ByVal value,ByVal fieldNum,ByVal mode,ByVal id,ByVal additionalCtrlParams,ByRef pageObj)
	Dim table,strLabel,is508,alt,cfield,clookupfield,openlookup,ctype,addnewitem,advancedadd,strCategoryControl,categoryFieldId,bUseCategory,dependentLookups,lookupType,LCType,horizontalLookup,inputStyle,lookupTable,strLookupWhere,lookupSize,add_page,list_page,strPerm,multiple,postfix,avalue,className,arr,res,opt,spacer,i,celementvalue,lookup_value,lookupSQL,rs_lookup,data,rs,found,checked
	doAssignment table,strTableName
	doAssignmentByRef strLabel,Label(field,"")
	doAssignmentByRef is508,isEnableSection508()
	alt = ""
	if (IsEqual(mode,MODE_INLINE_EDIT) or IsEqual(mode,MODE_INLINE_ADD)) and bValue(is508) then
		alt = (" alt=""" & CSmartStr(htmlspecialchars(strLabel))) & """ "
	end if
	cfield = (("value_" & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)
	clookupfield = (("display_value_" & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)
	openlookup = (("open_lookup_" & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)
	ctype = (("type_" & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)
	if bValue(fieldNum) then
		cfield = (((("value" & CSmartStr(fieldNum)) & "_") & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)
		ctype = (((("type" & CSmartStr(fieldNum)) & "_") & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)
	end if
	addnewitem = false
	advancedadd = false
	doAssignmentByRef strCategoryControl,CategoryControl(field,table)
	doAssignmentByRef categoryFieldId,GoodFieldName(CategoryControl(field,table))
	doAssignmentByRef bUseCategory,UseCategory(field,table)
	doAssignmentByRef dependentLookups,GetFieldData(table,field,"DependentLookups",CreateDictionary())
	doAssignmentByRef lookupType,GetLookupType(field,table)
	doAssignmentByRef LCType,LookupControlType(field,table)
	doAssignmentByRef horizontalLookup,GetFieldData(table,field,"HorizontalLookup",false)
	doAssignment inputStyle,IIF(ArrayElement(additionalCtrlParams,"style"),("style=""" & CSmartStr(ArrayElement(additionalCtrlParams,"style"))) & """","")
	doAssignmentByRef lookupTable,GetLookupTable(field,table)
	doAssignmentByRef strLookupWhere,LookupWhere(field,table)
	doAssignmentByRef lookupSize,SelectSize(field,table)
	if IsEqual(LCType,LCT_CBLIST) then
		lookupSize = 2
	end if
	add_page = CSmartStr(GetTableURL(lookupTable)) & "_add.asp"
	list_page = CSmartStr(GetTableURL(lookupTable)) & "_list.asp"
	doAssignmentByRef strPerm,GetUserPermissions(lookupTable)
	if not IsFalse(asp_strpos(strPerm,"A",empty)) then
		doAssignmentByRef addnewitem,GetFieldData(table,field,"AllowToAdd",false)
		advancedadd = not bValue(GetFieldData(table,field,"SimpleAdd",false))
		if not bValue(advancedadd) then
			addnewitem = false
		end if
	end if
	if IsEqual(LCType,LCT_LIST) and IsFalse(asp_strpos(strPerm,"S",empty)) then
		LCType = LCT_DROPDOWN
	end if
	if IsEqual(LCType,LCT_LIST) then
		addnewitem = false
	end if
	if IsEqual(mode,MODE_SEARCH) then
		addnewitem = false
	end if
	multiple = ""
	postfix = ""
	if IsLess(1,lookupSize) then
		doAssignmentByRef avalue,splitvalues(value)
		multiple = " multiple"
		postfix = "[]"
	else
		Set avalue = (CreateDictionary1(Empty,CSmartStr(value)))
	end if
	className = "DropDownLookup"
	if IsEqual(LCType,LCT_AJAX) then
		className = "EditBoxLookup"
	else
		if IsEqual(LCType,LCT_LIST) then
			className = "ListPageLookup"
		else
			if IsEqual(LCType,LCT_CBLIST) then
				className = "CheckBoxLookup"
			end if
		end if
	end if
	if IsEqual(lookupType,LT_LISTOFVALUES) then
		doAssignmentByRef arr,GetFieldData(table,field,"LookupValues",CreateDictionary())
		if IsLess(1,lookupSize) then
			ResponseWrite ((("<input id=""" & CSmartStr(ctype)) & """ type=hidden name=""") & CSmartStr(ctype)) & """ value=""multiselect"">"
		end if
		if IsEqual(LCType,LCT_DROPDOWN) then
			alt = ""
			ResponseWrite (((((((((("<select id=""" & CSmartStr(cfield)) & """ size = """) & CSmartStr(lookupSize)) & """ ") & CSmartStr(alt)) & "name=""") & CSmartStr(cfield)) & CSmartStr(postfix)) & """ ") & CSmartStr(multiple)) & ">"
			if IsLess(lookupSize,2) then
				ResponseWrite ("<option value="""">" & CSmartStr("Please select")) & "</option>"
			else
				if IsEqual(mode,MODE_SEARCH) then
					ResponseWrite "<option value=""""> </option>"
				end if
			end if
			GetCollectionBounds arr,i_commonfunctions_loopIdx32,i_commonfunctions_loopMax32
			do while i_commonfunctions_loopIdx32<=i_commonfunctions_loopMax32
				i_commonfunctions_arrKey32 = GetCollectionKey(arr,i_commonfunctions_loopIdx32)
				doAssignment opt,ArrayElement(arr,i_commonfunctions_arrKey32)
				doAssignmentByRef res,asp_array_search(CSmartStr(opt),avalue,false)
				if not (IsNull(res) or IsFalse(res)) then
					ResponseWrite ((("<option value=""" & CSmartStr(htmlspecialchars(opt))) & """ selected>") & CSmartStr(htmlspecialchars(opt))) & "</option>"
				else
					ResponseWrite ((("<option value=""" & CSmartStr(htmlspecialchars(opt))) & """>") & CSmartStr(htmlspecialchars(opt))) & "</option>"
				end if
				i_commonfunctions_loopIdx32=i_commonfunctions_loopIdx32+1
			loop
			ResponseWrite "</select>"
		else
			if IsEqual(LCType,LCT_CBLIST) then
				ResponseWrite "<div align='left'>"
				spacer = "<br/>"
				if bValue(horizontalLookup) then
					spacer = "&nbsp;"
				end if
				i = 0
				GetCollectionBounds arr,i_commonfunctions_loopIdx33,i_commonfunctions_loopMax33
				do while i_commonfunctions_loopIdx33<=i_commonfunctions_loopMax33
					i_commonfunctions_arrKey33 = GetCollectionKey(arr,i_commonfunctions_loopIdx33)
					doAssignment opt,ArrayElement(arr,i_commonfunctions_arrKey33)
					ResponseWrite (((((((((("<input id=""" & CSmartStr(cfield)) & "_") & CSmartStr(i)) & """ type=""checkbox"" ") & CSmartStr(alt)) & " name=""") & CSmartStr(cfield)) & CSmartStr(postfix)) & """ value=""") & CSmartStr(htmlspecialchars(opt))) & """"
					doAssignmentByRef res,asp_array_search(CSmartStr(opt),avalue,false)
					if not (IsNull(res) or IsFalse(res)) then
						ResponseWrite " checked=""checked"" "
					end if
					ResponseWrite "/>"
					ResponseWrite (((((("&nbsp;<b id=""data_" & CSmartStr(cfield)) & "_") & CSmartStr(i)) & """>") & CSmartStr(htmlspecialchars(opt))) & "</b>") & CSmartStr(spacer)
					i = CSmartDbl(i)+1
					i_commonfunctions_loopIdx33=i_commonfunctions_loopIdx33+1
				loop
				ResponseWrite "</div>"
			end if
		end if
		Exit Function
	end if
	if IsEqual(LCType,LCT_AJAX) or IsEqual(LCType,LCT_LIST) then
		if bValue(UseCategory(field,"")) then
			celementvalue = ((((("var parVal = ''; var parCtrl = Runner.controls.ControlManager.getAt('" & CSmartStr(jsreplace(strTableName))) & "', ") & CSmartStr(id)) & ", '") & CSmartStr(jsreplace(field))) & "', 0).parentCtrl; if (parCtrl){ parVal = parCtrl.getStringValue();};"
			if IsEqual(LCType,LCT_AJAX) then
				ResponseWrite ((((((("<input type=""text"" categoryId=""" & CSmartStr(categoryFieldId)) & """ autocomplete=""off"" id=""") & CSmartStr(clookupfield)) & """ name=""") & CSmartStr(clookupfield)) & """ ") & CSmartStr(inputStyle)) & ">"
			else
				if IsEqual(LCType,LCT_LIST) then
					ResponseWrite ((((((("<input type=""text"" categoryId=""" & CSmartStr(categoryFieldId)) & """ autocomplete=""off"" id=""") & CSmartStr(clookupfield)) & """ name=""") & CSmartStr(clookupfield)) & """  readonly ") & CSmartStr(inputStyle)) & ">"
					ResponseWrite ((("&nbsp;<a href=# id=" & CSmartStr(openlookup)) & ">") & CSmartStr("Select")) & "</a>"
				end if
			end if
			ResponseWrite ((("<input type=""hidden"" id=""" & CSmartStr(cfield)) & """ name=""") & CSmartStr(cfield)) & """>"
			if bValue(addnewitem) then
				ResponseWrite ((("&nbsp;<a href=# id='addnew_" & CSmartStr(cfield)) & "'>") & CSmartStr("Add new")) & "</a>"
			end if
			Exit Function
		end if
		lookup_value = ""
		doAssignmentByRef lookupSQL,buildLookupSQL(field,table,"",value,false,true,false,true,false)
		doAssignmentByRef rs_lookup,db_query(lookupSQL,conn)
		if bValue(doAssignmentByRef(data,db_fetch_numarray(rs_lookup))) then
			doAssignment lookup_value,ArrayElement(data,1)
		else
			if bValue(asp_strlen(strLookupWhere)) then
				doAssignmentByRef lookupSQL,buildLookupSQL(field,table,"",value,false,true,false,true,false)
				doAssignmentByRef rs_lookup,db_query(lookupSQL,conn)
				if bValue(doAssignmentByRef(data,db_fetch_numarray(rs_lookup))) then
					doAssignment lookup_value,ArrayElement(data,1)
				end if
			end if
		end if
		if IsEqual(LCType,LCT_AJAX) then
			if not bValue(asp_strlen(lookup_value)) and bValue(GetFieldData(strTableName,field,"freeInput",false)) then
				doAssignment lookup_value,value
			end if
			ResponseWrite ((((((((("<input type=""text"" " & CSmartStr(inputStyle)) & " autocomplete=""off"" ") & CSmartStr(IIF((IsEqual(mode,MODE_INLINE_EDIT) or IsEqual(mode,MODE_INLINE_ADD)) and IsEqual(is508,true),("alt=""" & CSmartStr(strLabel)) & """ ",""))) & "id=""") & CSmartStr(clookupfield)) & """ name=""") & CSmartStr(clookupfield)) & """ value=""") & CSmartStr(htmlspecialchars(lookup_value))) & """>"
		else
			if IsEqual(LCType,LCT_LIST) then
				ResponseWrite ((((((((("<input type=""text"" autocomplete=""off"" " & CSmartStr(inputStyle)) & " id=""") & CSmartStr(clookupfield)) & """ ") & CSmartStr(IIF((IsEqual(mode,MODE_INLINE_EDIT) or IsEqual(mode,MODE_INLINE_ADD)) and IsEqual(is508,true),("alt=""" & CSmartStr(strLabel)) & """ ",""))) & "name=""") & CSmartStr(clookupfield)) & """ value=""") & CSmartStr(htmlspecialchars(lookup_value))) & """ 	readonly >"
				ResponseWrite ((("&nbsp;<a href=# id=" & CSmartStr(openlookup)) & ">") & CSmartStr("Select")) & "</a>"
			end if
		end if
		ResponseWrite ((((("<input type=""hidden"" id=""" & CSmartStr(cfield)) & """ name=""") & CSmartStr(cfield)) & """ value=""") & CSmartStr(htmlspecialchars(value))) & """>"
		if bValue(addnewitem) then
			ResponseWrite ((("&nbsp;<a href=# id='addnew_" & CSmartStr(cfield)) & "'>") & CSmartStr("Add new")) & "</a>"
		end if
		Exit Function
	end if
	doAssignmentByRef lookupSQL,buildLookupSQL(field,table,"","",false,false,false,true,false)
	doAssignmentByRef rs,db_query(lookupSQL,conn)
	if bValue(bUseCategory) then
		if IsLess(1,lookupSize) then
			ResponseWrite ((("<input id=""" & CSmartStr(ctype)) & """ type=hidden name=""") & CSmartStr(ctype)) & """ value=""multiselect"">"
		end if
		ResponseWrite (((((((("<select size = """ & CSmartStr(lookupSize)) & """ id=""") & CSmartStr(cfield)) & """ name=""") & CSmartStr(cfield)) & CSmartStr(postfix)) & """") & CSmartStr(multiple)) & ">"
		ResponseWrite ("<option value="""">" & CSmartStr("Please select")) & "</option>"
		ResponseWrite "</select>"
		if bValue(addnewitem) then
			ResponseWrite ((("&nbsp;<a href=# id='addnew_" & CSmartStr(cfield)) & "'>") & CSmartStr("Add new")) & "</a>"
		end if
		Exit Function
	end if
	if IsLess(1,lookupSize) then
		ResponseWrite ((("<input id=""" & CSmartStr(ctype)) & """ type=hidden name=""") & CSmartStr(ctype)) & """ value=""multiselect"">"
	end if
	if not IsEqual(LCType,LCT_CBLIST) then
		ResponseWrite (((((((((("<select size = """ & CSmartStr(lookupSize)) & """ id=""") & CSmartStr(cfield)) & """ ") & CSmartStr(IIF((IsEqual(mode,MODE_INLINE_EDIT) or IsEqual(mode,MODE_INLINE_ADD)) and IsEqual(is508,true),("alt=""" & CSmartStr(strLabel)) & """ ",""))) & "name=""") & CSmartStr(cfield)) & CSmartStr(postfix)) & """") & CSmartStr(multiple)) & ">"
		if IsLess(lookupSize,2) then
			ResponseWrite ("<option value="""">" & CSmartStr("Please select")) & "</option>"
		else
			if IsEqual(mode,MODE_SEARCH) then
				ResponseWrite "<option value=""""> </option>"
			end if
		end if
	else
		ResponseWrite "<div align='left'>"
		spacer = "<br/>"
		if bValue(horizontalLookup) then
			spacer = "&nbsp;"
		end if
	end if
	found = false
	i = 0
	do while bValue(doAssignmentByRef(data,db_fetch_numarray(rs)))
		doAssignmentByRef res,asp_array_search(CSmartStr(ArrayElement(data,0)),avalue,false)
		checked = ""
		if not (IsNull(res) or IsFalse(res)) then
			found = true
			if IsEqual(LCType,LCT_CBLIST) then
				checked = " checked=""checked"""
			else
				checked = " selected"
			end if
		end if
		if IsEqual(LCType,LCT_CBLIST) then
			ResponseWrite (((((((((((("<input id=""" & CSmartStr(cfield)) & "_") & CSmartStr(i)) & """ type=""checkbox"" ") & CSmartStr(alt)) & " name=""") & CSmartStr(cfield)) & CSmartStr(postfix)) & """ value=""") & CSmartStr(htmlspecialchars(ArrayElement(data,0)))) & """") & CSmartStr(checked)) & "/>"
			ResponseWrite (((((("&nbsp;<b id=""data_" & CSmartStr(cfield)) & "_") & CSmartStr(i)) & """>") & CSmartStr(htmlspecialchars(ArrayElement(data,1)))) & "</b>") & CSmartStr(spacer)
		else
			ResponseWrite ((((("<option value=""" & CSmartStr(htmlspecialchars(ArrayElement(data,0)))) & """") & CSmartStr(checked)) & ">") & CSmartStr(htmlspecialchars(ArrayElement(data,1)))) & "</option>"
		end if
		i = CSmartDbl(i)+1
	loop
	if ((not bValue(found) and bValue(asp_strlen(value))) and IsEqual(mode,MODE_EDIT)) and bValue(asp_strlen(strLookupWhere)) then
		doAssignmentByRef lookupSQL,buildLookupSQL(field,table,"",value,false,true,false,false,true)
		doAssignmentByRef rs,db_query(lookupSQL,conn)
		if bValue(doAssignmentByRef(data,db_fetch_numarray(rs))) then
			if IsEqual(LCType,LCT_CBLIST) then
				ResponseWrite (((((((((("<input id=""" & CSmartStr(cfield)) & "_") & CSmartStr(i)) & """ type=""checkbox"" ") & CSmartStr(alt)) & " name=""") & CSmartStr(cfield)) & CSmartStr(postfix)) & """ value=""") & CSmartStr(htmlspecialchars(ArrayElement(data,0)))) & """ checked=""checked""/>"
				ResponseWrite (((((("&nbsp;<b id=""data_" & CSmartStr(cfield)) & "_") & CSmartStr(i)) & """>") & CSmartStr(htmlspecialchars(ArrayElement(data,1)))) & "</b>") & CSmartStr(spacer)
			else
				ResponseWrite ((("<option value=""" & CSmartStr(htmlspecialchars(ArrayElement(data,0)))) & """ selected>") & CSmartStr(htmlspecialchars(ArrayElement(data,1)))) & "</option>"
			end if
		end if
	end if
	if not IsEqual(LCType,LCT_CBLIST) then
		ResponseWrite "</select>"
	else
		ResponseWrite "</div>"
	end if
	if bValue(addnewitem) then
		ResponseWrite ((("&nbsp;<a href=# id='addnew_" & CSmartStr(cfield)) & "'>") & CSmartStr("Add new")) & "</a>"
	end if
End Function
Function BuildRadioControl(ByVal field,ByVal value,ByVal fieldNum,ByVal id,ByVal mode)
	Dim is508,strLabel,cfieldname,cfield,ctype,spacer,arr,rs,i,data,checked,opt
	doAssignmentByRef is508,isEnableSection508()
	doAssignmentByRef strLabel,Label(field,"")
	cfieldname = (CSmartStr(GoodFieldName(field)) & "_") & CSmartStr(id)
	cfield = (("value_" & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)
	ctype = (("type_" & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)
	if bValue(fieldNum) then
		cfield = (((("value" & CSmartStr(fieldNum)) & "_") & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)
		ctype = (((("type" & CSmartStr(fieldNum)) & "_") & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)
	end if
	LookupSQL = ""
	spacer = "<br/>"
	if bValue(LookupSQL) then
		LogInfo LookupSQL
		doAssignmentByRef rs,db_query(LookupSQL,conn)
		ResponseWrite ((((("<input id=""" & CSmartStr(cfield)) & """ type=hidden name=""") & CSmartStr(cfield)) & """ value=""") & CSmartStr(htmlspecialchars(value))) & """>"
		i = 0
		do while bValue(doAssignmentByRef(data,db_fetch_numarray(rs)))
			checked = ""
			if IsEqual(ArrayElement(data,0),value) then
				checked = " checked"
			end if
			ResponseWrite ((((((((((((("<input type=""Radio"" id=""radio_" & CSmartStr(cfieldname)) & "_") & CSmartStr(i)) & """ ") & CSmartStr(IIF((IsEqual(mode,MODE_INLINE_EDIT) or IsEqual(mode,MODE_INLINE_ADD)) and IsEqual(is508,true),("alt=""" & CSmartStr(strLabel)) & """ ",""))) & "name=""radio_") & CSmartStr(cfieldname)) & """ ") & CSmartStr(checked)) & " value=""") & CSmartStr(htmlspecialchars(ArrayElement(data,0)))) & """>") & CSmartStr(htmlspecialchars(ArrayElement(data,1)))) & CSmartStr(spacer)
			i = CSmartDbl(i)+1
		loop
	else
		ResponseWrite ((((("<input id=""" & CSmartStr(cfield)) & """ type=hidden name=""") & CSmartStr(cfield)) & """ value=""") & CSmartStr(htmlspecialchars(value))) & """>"
		i = 0
		GetCollectionBounds arr,i_commonfunctions_loopIdx36,i_commonfunctions_loopMax36
		do while i_commonfunctions_loopIdx36<=i_commonfunctions_loopMax36
			i_commonfunctions_arrKey36 = GetCollectionKey(arr,i_commonfunctions_loopIdx36)
			doAssignment opt,ArrayElement(arr,i_commonfunctions_arrKey36)
			checked = ""
			if IsEqual(opt,value) then
				checked = " checked"
			end if
			ResponseWrite ((((((((((((("<input  type=""Radio"" id=""radio_" & CSmartStr(cfieldname)) & "_") & CSmartStr(i)) & """ ") & CSmartStr(IIF((IsEqual(mode,MODE_INLINE_EDIT) or IsEqual(mode,MODE_INLINE_ADD)) and IsEqual(is508,true),("alt=""" & CSmartStr(strLabel)) & """ ",""))) & "name=""radio_") & CSmartStr(cfieldname)) & """ ") & CSmartStr(checked)) & " value=""") & CSmartStr(htmlspecialchars(opt))) & """>") & CSmartStr(htmlspecialchars(opt))) & CSmartStr(spacer)
			i = CSmartDbl(i)+1
			i_commonfunctions_loopIdx36=i_commonfunctions_loopIdx36+1
		loop
	end if
	Exit Function
End Function
Function getLacaleAmPmForTimePicker(ByVal convention,ByVal useTimePicker)
	Dim am,pm,locale_convention,locale
	am = ""
	pm = ""
	if bValue(useTimePicker) then
		doAssignment locale_convention,IIF(ArrayElement(locale_info,"LOCALE_ITIME"),24,12)
		if IsEqual(convention,locale_convention) then
			doAssignment am,ArrayElement(locale_info,"LOCALE_S1159")
			doAssignment pm,ArrayElement(locale_info,"LOCALE_S2359")
			doAssignment locale,ArrayElement(locale_info,"LOCALE_STIMEFORMAT")
		else
			if IsEqual(convention,24) then
				am = ""
				pm = ""
				locale = "H:mm:ss"
			else
				am = "am"
				pm = "pm"
				locale = "h:mm:ss tt"
			end if
		end if
	else
		doAssignment locale,ArrayElement(locale_info,"LOCALE_STIMEFORMAT")
	end if
	doAssignmentByRef getLacaleAmPmForTimePicker,CreateDictionary3("am",am,"pm",pm,"locale",locale)
	Exit Function
End Function
Function getValForTimePicker(ByVal var_type,ByVal value,ByVal locale)
	Dim val,dbtime,arr
	val = ""
	Set dbtime = (CreateDictionary())
	if bValue(IsDateFieldType(var_type)) then
		doAssignmentByRef dbtime,db2time(value)
		if bValue(asp_count(dbtime)) then
			doAssignmentByRef val,format_datetime_custom(dbtime,locale)
		end if
	else
		doAssignmentByRef arr,parsenumbers(value)
		if bValue(asp_count(arr)) then
			do while IsLess(asp_count(arr),3)
				setArrElement arr,asp_count(arr),0
			loop
			Set dbtime = (CreateDictionary6(Empty,0,Empty,0,Empty,0,Empty,ArrayElement(arr,0),Empty,ArrayElement(arr,1),Empty,ArrayElement(arr,2)))
			doAssignmentByRef val,format_datetime_custom(dbtime,locale)
		end if
	end if
	doAssignmentByRef getValForTimePicker,CreateDictionary2("val",val,"dbTime",dbtime)
	Exit Function
End Function
Function BuildEditControl(ByVal field,ByVal value,ByVal format,ByVal edit,ByVal fieldNum,ByVal id,ByVal validate,ByVal additionalCtrlParams,ByRef pageObj)
	Dim inputStyle,cfieldname,cfield,ctype,is508,strLabel,var_type,arr,iquery,keylink,arrKeys,j,isHidden,arr_number,timeAttrs,convention,loc,tpVal,nWidth,nHeight,browser,var_REQUEST,checked,val,show,sel,v,i,disp,strfilename,itype,thumbnailed,thumbfield,filename,strtype,var_function,filename_size
	inputStyle = "style="""
	inputStyle = CSmartStr(inputStyle) & CSmartStr(IIF(ArrayElement(additionalCtrlParams,"style"),ArrayElement(additionalCtrlParams,"style"),""))
	inputStyle = CSmartStr(inputStyle) & """"
	cfieldname = (CSmartStr(GoodFieldName(field)) & "_") & CSmartStr(id)
	cfield = (("value_" & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)
	ctype = (("type_" & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)
	doAssignmentByRef is508,isEnableSection508()
	doAssignmentByRef strLabel,Label(field,"")
	if bValue(fieldNum) then
		cfield = (((("value" & CSmartStr(fieldNum)) & "_") & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)
		ctype = (((("type" & CSmartStr(fieldNum)) & "_") & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)
	end if
	doAssignmentByRef var_type,GetFieldType(field,"")
	arr = ""
	iquery = "field=" & CSmartStr(asp_rawurlencode(field))
	keylink = ""
	doAssignmentByRef arrKeys,GetTableKeys(strTableName)
	j = 0
	do while IsLess(j,asp_count(arrKeys))
		keylink = CSmartStr(keylink) & ((("&key" & CSmartStr(CSmartDbl(j)+1)) & "=") & CSmartStr(asp_rawurlencode(ArrayElement(data,ArrayElement(arrKeys,j)))))
		j = CSmartDbl(j)+1
	loop
	iquery = CSmartStr(iquery) & CSmartStr(keylink)
	isHidden = not IsEmpty(ArrayElement(additionalCtrlParams,"hidden")) and bValue(ArrayElement(additionalCtrlParams,"hidden"))
	ResponseWrite (((((((("<span id=""edit" & CSmartStr(id)) & "_") & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(fieldNum)) & """ style=""") & CSmartStr(IIF(isHidden,"display:none;",""))) & CSmartStr(IIF((IsEqual(format,EDIT_FORMAT_DATE) or IsEqual(format,EDIT_FORMAT_TEXT_FIELD)) and not IsEqual(edit,MODE_SEARCH),"white-space: nowrap;",""))) & """>"
	if IsEqual(format,EDIT_FORMAT_FILE) and IsEqual(edit,MODE_SEARCH) then
		format = ""
	end if
	if IsEqual(format,EDIT_FORMAT_TEXT_FIELD) then
		if bValue(IsDateFieldType(var_type)) then
			ResponseWrite (((((("<input id=""" & CSmartStr(ctype)) & """ type=""hidden"" name=""") & CSmartStr(ctype)) & """ value=""date") & CSmartStr(EDIT_DATE_SIMPLE)) & """>") & CSmartStr(GetDateEdit(field,value,0,fieldNum,edit,id,pageObj))
		else
			if IsEqual(edit,MODE_SEARCH) then
				ResponseWrite ((((((((((("<input id=""" & CSmartStr(cfield)) & """ ") & CSmartStr(inputStyle)) & " type=""text"" autocomplete=""off"" ") & CSmartStr(IIF((IsEqual(edit,MODE_INLINE_EDIT) or IsEqual(edit,MODE_INLINE_ADD)) and IsEqual(is508,true),("alt=""" & CSmartStr(strLabel)) & """ ",""))) & "name=""") & CSmartStr(cfield)) & """ ") & CSmartStr(GetEditParams(field,""))) & " value=""") & CSmartStr(htmlspecialchars(value))) & """>"
			else
				ResponseWrite ((((((((((("<input id=""" & CSmartStr(cfield)) & """ ") & CSmartStr(inputStyle)) & " type=""text"" ") & CSmartStr(IIF((IsEqual(edit,MODE_INLINE_EDIT) or IsEqual(edit,MODE_INLINE_ADD)) and IsEqual(is508,true),("alt=""" & CSmartStr(strLabel)) & """ ",""))) & "name=""") & CSmartStr(cfield)) & """ ") & CSmartStr(GetEditParams(field,""))) & " value=""") & CSmartStr(htmlspecialchars(value))) & """>"
			end if
		end if
	else
		if IsEqual(format,EDIT_FORMAT_TIME) then
			ResponseWrite ((((("<input id=""" & CSmartStr(ctype)) & """ ") & CSmartStr(inputStyle)) & " type=""hidden"" name=""") & CSmartStr(ctype)) & """ value=""time"">"
			doAssignmentByRef arr_number,parsenumbers(CSmartStr(value))
			if IsEqual(asp_count(arr_number),6) then
				doAssignmentByRef value,mysprintf("%d:%02d:%02d",CreateDictionary3(Empty,ArrayElement(arr_number,3),Empty,ArrayElement(arr_number,4),Empty,ArrayElement(arr_number,5)))
			end if
			doAssignmentByRef timeAttrs,GetFieldData(strTableName,field,"FormatTimeAttrs",CreateDictionary())
			if bValue(asp_count(timeAttrs)) then
				if bValue(ArrayElement(timeAttrs,"useTimePicker")) then
					doAssignment convention,ArrayElement(timeAttrs,"hours")
					doAssignmentByRef loc,getLacaleAmPmForTimePicker(convention,true)
					doAssignmentByRef tpVal,getValForTimePicker(var_type,value,ArrayElement(loc,"locale"))
					ResponseWrite ((((((((((("<input type=""text"" " & CSmartStr(inputStyle)) & " name=""") & CSmartStr(cfield)) & """ ") & CSmartStr(IIF((IsEqual(edit,MODE_INLINE_EDIT) or IsEqual(edit,MODE_INLINE_ADD)) and IsEqual(is508,true),("alt=""" & CSmartStr(strLabel)) & """ ",""))) & "id=""") & CSmartStr(cfield)) & """ ") & CSmartStr(GetEditParams(field,""))) & " value=""") & CSmartStr(htmlspecialchars(ArrayElement(tpVal,"val")))) & """>"
					ResponseWrite "&nbsp;"
					ResponseWrite ("<img src=""include/img/clock.gif"" alt=""Time"" border=""0"" style=""margin:4px 0 0 6px; visibility: hidden;"" id=""trigger-test-" & CSmartStr(cfield)) & """ />"
				else
					ResponseWrite ((((((((((("<input id=""" & CSmartStr(cfield)) & """ ") & CSmartStr(inputStyle)) & " type=""text"" ") & CSmartStr(IIF((IsEqual(edit,MODE_INLINE_EDIT) or IsEqual(edit,MODE_INLINE_ADD)) and IsEqual(is508,true),("alt=""" & CSmartStr(strLabel)) & """ ",""))) & "name=""") & CSmartStr(cfield)) & """ ") & CSmartStr(GetEditParams(field,""))) & " value=""") & CSmartStr(htmlspecialchars(value))) & """>"
				end if
			end if
		else
			if IsEqual(format,EDIT_FORMAT_TEXT_AREA) then
				doAssignmentByRef nWidth,GetNCols(field,"")
				doAssignmentByRef nHeight,GetNRows(field,"")
				if bValue(UseRTE(field,"")) then
					doAssignmentByRef value,RTESafe(value)
					browser = ""
					if IsEqual(GetRequestValue(Request,"browser"),"ie") then
						browser = "&browser=ie"
					end if
					ResponseWrite ((((((((("<iframe frameborder=""0"" vspace=""0"" hspace=""0"" marginwidth=""0"" marginheight=""0"" scrolling=""no"" id=""" & CSmartStr(cfield)) & """ ") & CSmartStr(IIF((IsEqual(edit,MODE_INLINE_EDIT) or IsEqual(edit,MODE_INLINE_ADD)) and IsEqual(is508,true),("alt=""" & CSmartStr(strLabel)) & """ ",""))) & "name=""") & CSmartStr(cfield)) & """ title=""Basic rich text editor"" style='width: ") & CSmartStr(CSmartDbl(nWidth)+1)) & "px;height: ") & CSmartStr(CSmartDbl(nHeight)+100)) & "px;'"
					ResponseWrite (((((((((" src=""rte.asp?table=" & CSmartStr(GetTableURL(strTableName))) & "&") & "id=") & CSmartStr(id)) & "&") & CSmartStr(iquery)) & CSmartStr(browser)) & "&") & CSmartStr(IIF(IsEqual(edit,MODE_ADD) or IsEqual(edit,MODE_INLINE_ADD),"action=add",""))) & """>"
					ResponseWrite "</iframe>"
				else
					ResponseWrite ((((((((((("<textarea id=""" & CSmartStr(cfield)) & """ ") & CSmartStr(IIF((IsEqual(edit,MODE_INLINE_EDIT) or IsEqual(edit,MODE_INLINE_ADD)) and IsEqual(is508,true),("alt=""" & CSmartStr(strLabel)) & """ ",""))) & "name=""") & CSmartStr(cfield)) & """ style=""width: ") & CSmartStr(nWidth)) & "px;height: ") & CSmartStr(nHeight)) & "px;"">") & CSmartStr(htmlspecialchars(value))) & "</textarea>"
				end if
			else
				if IsEqual(format,EDIT_FORMAT_PASSWORD) then
					ResponseWrite ((((((((((("<input " & CSmartStr(inputStyle)) & " id=""") & CSmartStr(cfield)) & """ type=""Password"" ") & CSmartStr(IIF((IsEqual(edit,MODE_INLINE_EDIT) or IsEqual(edit,MODE_INLINE_ADD)) and IsEqual(is508,true),("alt=""" & CSmartStr(strLabel)) & """ ",""))) & "name=""") & CSmartStr(cfield)) & """ ") & CSmartStr(GetEditParams(field,""))) & " value=""") & CSmartStr(htmlspecialchars(value))) & """>"
				else
					if IsEqual(format,EDIT_FORMAT_DATE) then
						ResponseWrite (((((("<input id=""" & CSmartStr(ctype)) & """ type=""hidden"" name=""") & CSmartStr(ctype)) & """ value=""date") & CSmartStr(DateEditType(field,""))) & """>") & CSmartStr(GetDateEdit(field,value,DateEditType(field,""),fieldNum,edit,id,pageObj))
					else
						if IsEqual(format,EDIT_FORMAT_RADIO) then
							BuildRadioControl field,value,fieldNum,id,edit
						else
							if IsEqual(format,EDIT_FORMAT_CHECKBOX) then
								if ((IsEqual(edit,MODE_ADD) or IsEqual(edit,MODE_INLINE_ADD)) or IsEqual(edit,MODE_EDIT)) or IsEqual(edit,MODE_INLINE_EDIT) then
									checked = ""
									if bValue(value) and not IsEqual(value,0) then
										checked = " checked"
									end if
									ResponseWrite ((("<input id=""" & CSmartStr(ctype)) & """ type=""hidden"" name=""") & CSmartStr(ctype)) & """ value=""checkbox"">"
									ResponseWrite ((((((("<input id=""" & CSmartStr(cfield)) & """ type=""Checkbox"" ") & CSmartStr(IIF((IsEqual(edit,MODE_INLINE_EDIT) or IsEqual(edit,MODE_INLINE_ADD)) and IsEqual(is508,true),("alt=""" & CSmartStr(strLabel)) & """ ",""))) & "name=""") & CSmartStr(cfield)) & """ ") & CSmartStr(checked)) & ">"
								else
									ResponseWrite ((("<input id=""" & CSmartStr(ctype)) & """ type=""hidden"" name=""") & CSmartStr(ctype)) & """ value=""checkbox"">"
									ResponseWrite ((((("<select id=""" & CSmartStr(cfield)) & """ ") & CSmartStr(IIF((IsEqual(edit,MODE_INLINE_EDIT) or IsEqual(edit,MODE_INLINE_ADD)) and IsEqual(is508,true),("alt=""" & CSmartStr(strLabel)) & """ ",""))) & "name=""") & CSmartStr(cfield)) & """>"
									Set val = (CreateDictionary3(Empty,"",Empty,"on",Empty,"off"))
									Set show = (CreateDictionary3(Empty,"",Empty,"True",Empty,"False"))
									GetCollectionBounds val,i_commonfunctions_loopIdx39,i_commonfunctions_loopMax39
									do while i_commonfunctions_loopIdx39<=i_commonfunctions_loopMax39
										i = GetCollectionKey(val,i_commonfunctions_loopIdx39)
										doAssignment v,ArrayElement(val,i)
										sel = ""
										if IsIdentical(value,v) then
											sel = " selected"
										end if
										ResponseWrite ((((("<option value=""" & CSmartStr(v)) & """") & CSmartStr(sel)) & ">") & CSmartStr(ArrayElement(show,i))) & "</option>"
										i_commonfunctions_loopIdx39=i_commonfunctions_loopIdx39+1
									loop
									ResponseWrite "</select>"
								end if
							else
								if IsEqual(format,EDIT_FORMAT_DATABASE_IMAGE) or IsEqual(format,EDIT_FORMAT_DATABASE_FILE) then
									disp = ""
									strfilename = ""
									if IsEqual(edit,MODE_EDIT) or IsEqual(edit,MODE_INLINE_EDIT) then
										doAssignmentByRef value,db_stripslashesbinary(value)
										doAssignmentByRef itype,SupposeImageType(value)
										thumbnailed = false
										thumbfield = ""
										if bValue(itype) then
											if bValue(thumbnailed) then
												disp = "<a "
												if bValue(IsUseiBox(field,strTableName)) then
													disp = CSmartStr(disp) & " rel='ibox'"
												else
													disp = CSmartStr(disp) & " target=_blank"
												end if
												disp = CSmartStr(disp) & ((((" href=""imager.asp?table=" & CSmartStr(GetTableURL(strTableName))) & "&") & CSmartStr(iquery)) & """>")
												disp = CSmartStr(disp) & (((((("<img id=""image_" & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)) & """ name=""") & CSmartStr(cfield)) & """ border=0")
												if bValue(isEnableSection508()) then
													disp = CSmartStr(disp) & " alt=""Image from DB"""
												end if
												disp = CSmartStr(disp) & (((((((" src=""imager.asp?table=" & CSmartStr(GetTableURL(strTableName))) & "&field=") & CSmartStr(asp_rawurlencode(thumbfield))) & "&alt=") & CSmartStr(asp_rawurlencode(field))) & CSmartStr(keylink)) & """>")
												disp = CSmartStr(disp) & "</a>"
											else
												disp = ((((("<img id=""image_" & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)) & """ name=""") & CSmartStr(cfield)) & """"
												if bValue(isEnableSection508()) then
													disp = CSmartStr(disp) & " alt=""Image from DB"""
												end if
												disp = CSmartStr(disp) & ((((" border=0 src=""imager.asp?table=" & CSmartStr(GetTableURL(strTableName))) & "&") & CSmartStr(iquery)) & """>")
											end if
										else
											if bValue(asp_strlen(value)) then
												disp = ((((("<img id=""image_" & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)) & """ name=""") & CSmartStr(cfield)) & """ border=0 "
												if bValue(isEnableSection508()) then
													disp = CSmartStr(disp) & " alt=""file"""
												end if
												disp = CSmartStr(disp) & " src=""images/file.gif"">"
											else
												disp = ((((("<img id=""image_" & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)) & """ name=""") & CSmartStr(cfield)) & """ border=""0"""
												if bValue(isEnableSection508()) then
													disp = CSmartStr(disp) & " alt="" """
												end if
												disp = CSmartStr(disp) & " src=""images/no_image.gif"">"
											end if
										end if
										if (IsEqual(format,EDIT_FORMAT_DATABASE_FILE) and not bValue(itype)) and bValue(asp_strlen(value)) then
											if not bValue(doAssignment(filename,ArrayElement(data,GetFilenameField(field,"")))) then
												filename = "file.bin"
											end if
											disp = ((((((("<a href=""getfile.asp?table=" & CSmartStr(GetTableURL(strTableName))) & "&filename=") & CSmartStr(htmlspecialchars(filename))) & "&") & CSmartStr(iquery)) & """.>") & CSmartStr(disp)) & "</a>"
										end if
										if IsEqual(format,EDIT_FORMAT_DATABASE_FILE) and bValue(GetFilenameField(field,"")) then
											if not bValue(doAssignment(filename,ArrayElement(data,GetFilenameField(field,"")))) then
												filename = ""
											end if
											if IsEqual(edit,MODE_INLINE_EDIT) then
												strfilename = ((((((((((("<br><label for=""filename_" & CSmartStr(cfieldname)) & """>") & CSmartStr("Filename")) & "</label>&nbsp;&nbsp;<input type=""text"" ") & CSmartStr(inputStyle)) & " id=""filename_") & CSmartStr(cfieldname)) & """ name=""filename_") & CSmartStr(cfieldname)) & """ size=""20"" maxlength=""50"" value=""") & CSmartStr(htmlspecialchars(filename))) & """>"
											else
												strfilename = ((((((((((("<br><label for=""filename_" & CSmartStr(cfieldname)) & """>") & CSmartStr("Filename")) & "</label>&nbsp;&nbsp;<input type=""text"" ") & CSmartStr(inputStyle)) & " id=""filename_") & CSmartStr(cfieldname)) & """ name=""filename_") & CSmartStr(cfieldname)) & """ size=""20"" maxlength=""50"" value=""") & CSmartStr(htmlspecialchars(filename))) & """>"
											end if
										end if
										strtype = (((("<br><input id=""" & CSmartStr(ctype)) & "_keep"" type=""Radio"" name=""") & CSmartStr(ctype)) & """ value=""file0"" checked>") & CSmartStr("Keep")
										if (bValue(asp_strlen(value)) or IsEqual(edit,MODE_INLINE_EDIT)) and not bValue(IsRequired(field,"")) then
											strtype = CSmartStr(strtype) & ((((("<input id=""" & CSmartStr(ctype)) & "_delete"" type=""Radio"" name=""") & CSmartStr(ctype)) & """ value=""file1"">") & CSmartStr("Delete"))
										end if
										strtype = CSmartStr(strtype) & ((((("<input id=""" & CSmartStr(ctype)) & "_update"" type=""Radio"" name=""") & CSmartStr(ctype)) & """ value=""file2"">") & CSmartStr("Update"))
									else
										strtype = ((("<input id=""" & CSmartStr(ctype)) & """ type=""hidden"" name=""") & CSmartStr(ctype)) & """ value=""file2"">"
										if IsEqual(format,EDIT_FORMAT_DATABASE_FILE) and bValue(GetFilenameField(field,"")) then
											strfilename = ((((((((("<br><label for=""filename_" & CSmartStr(cfieldname)) & """>") & CSmartStr("Filename")) & "</label>&nbsp;&nbsp;<input type=""text"" ") & CSmartStr(inputStyle)) & " id=""filename_") & CSmartStr(cfieldname)) & """ name=""filename_") & CSmartStr(cfieldname)) & """ size=""20"" maxlength=""50"">"
										end if
									end if
									if IsEqual(edit,MODE_INLINE_EDIT) and IsEqual(format,EDIT_FORMAT_DATABASE_FILE) then
										disp = ""
									end if
									ResponseWrite CSmartStr(disp) & CSmartStr(strtype)
									if IsEqual(edit,MODE_EDIT) or IsEqual(edit,MODE_INLINE_EDIT) then
										ResponseWrite "<br>"
									end if
									ResponseWrite (((((((("<input type=""File"" " & CSmartStr(inputStyle)) & " id=""") & CSmartStr(cfield)) & """ ") & CSmartStr(IIF((IsEqual(edit,MODE_INLINE_EDIT) or IsEqual(edit,MODE_INLINE_ADD)) and IsEqual(is508,true),("alt=""" & CSmartStr(strLabel)) & """ ",""))) & " name=""") & CSmartStr(cfield)) & """ >") & CSmartStr(strfilename)
									ResponseWrite ((("<input type=""Hidden"" id=""notempty_" & CSmartStr(cfieldname)) & """ value=""") & CSmartStr(IIF(asp_strlen(value),1,0))) & """>"
								else
									if IsEqual(format,EDIT_FORMAT_LOOKUP_WIZARD) then
										BuildSelectControl field,value,fieldNum,edit,id,additionalCtrlParams,pageObj
									else
										if IsEqual(format,EDIT_FORMAT_HIDDEN) then
											ResponseWrite ((((("<input id=""" & CSmartStr(cfield)) & """ type=""Hidden"" name=""") & CSmartStr(cfield)) & """ value=""") & CSmartStr(htmlspecialchars(value))) & """>"
										else
											if IsEqual(format,EDIT_FORMAT_READONLY) then
												ResponseWrite ((((("<input id=""" & CSmartStr(cfield)) & """ type=""Hidden"" name=""") & CSmartStr(cfield)) & """ value=""") & CSmartStr(htmlspecialchars(value))) & """>"
											else
												if IsEqual(format,EDIT_FORMAT_FILE) then
													disp = ""
													strfilename = ""
													var_function = ""
													if IsEqual(edit,MODE_EDIT) or IsEqual(edit,MODE_INLINE_EDIT) then
														if IsEqual(ViewFormat(field,""),FORMAT_FILE) or IsEqual(ViewFormat(field,""),FORMAT_FILE_IMAGE) then
															disp = CSmartStr(GetData(data,field,ViewFormat(field,""))) & "<br>"
														end if
														doAssignment filename,value
														filename_size = 30
														if bValue(UseTimestamp(field,"")) then
															filename_size = 50
														end if
														strfilename = ((((((((((((("<input type=hidden name=""filenameHidden_" & CSmartStr(cfieldname)) & """ value=""") & CSmartStr(htmlspecialchars(filename))) & """><br>") & CSmartStr("Filename")) & "&nbsp;&nbsp;<input type=""text"" style=""background-color:gainsboro"" disabled id=""filename_") & CSmartStr(cfieldname)) & """ name=""filename_") & CSmartStr(cfieldname)) & """ size=""") & CSmartStr(filename_size)) & """ maxlength=""100"" value=""") & CSmartStr(htmlspecialchars(filename))) & """>"
														if IsEqual(edit,MODE_INLINE_EDIT) then
															strtype = (((("<br><input id=""" & CSmartStr(ctype)) & "_keep"" type=""Radio"" name=""") & CSmartStr(ctype)) & """ value=""upload0"" checked>") & CSmartStr("Keep")
														else
															strtype = (((("<br><input id=""" & CSmartStr(ctype)) & "_keep"" type=""Radio"" name=""") & CSmartStr(ctype)) & """ value=""upload0"" checked>") & CSmartStr("Keep")
														end if
														if (bValue(asp_strlen(value)) or IsEqual(edit,MODE_INLINE_EDIT)) and not bValue(IsRequired(field,"")) then
															strtype = CSmartStr(strtype) & ((((("<input id=""" & CSmartStr(ctype)) & "_delete"" type=""Radio"" name=""") & CSmartStr(ctype)) & """ value=""upload1"">") & CSmartStr("Delete"))
														end if
														strtype = CSmartStr(strtype) & ((((("<input id=""" & CSmartStr(ctype)) & "_update"" type=""Radio"" name=""") & CSmartStr(ctype)) & """ value=""upload2"">") & CSmartStr("Update"))
													else
														filename_size = 30
														if bValue(UseTimestamp(field,"")) then
															filename_size = 50
														end if
														strtype = ((("<input id=""" & CSmartStr(ctype)) & """ type=""hidden"" name=""") & CSmartStr(ctype)) & """ value=""upload2"">"
														strfilename = ((((((("<br>" & CSmartStr("Filename")) & "&nbsp;&nbsp;<input type=""text"" id=""filename_") & CSmartStr(cfieldname)) & """ name=""filename_") & CSmartStr(cfieldname)) & """ size=""") & CSmartStr(filename_size)) & """ maxlength=""100"">"
													end if
													ResponseWrite (CSmartStr(disp) & CSmartStr(strtype)) & CSmartStr(var_function)
													if IsEqual(edit,MODE_EDIT) or IsEqual(edit,MODE_INLINE_EDIT) then
														ResponseWrite "<br>"
													end if
													ResponseWrite (((((("<input type=""File"" id=""" & CSmartStr(cfield)) & """ ") & CSmartStr(IIF((IsEqual(edit,MODE_INLINE_EDIT) or IsEqual(edit,MODE_INLINE_ADD)) and IsEqual(is508,true),("alt=""" & CSmartStr(strLabel)) & """ ",""))) & " name=""") & CSmartStr(cfield)) & """ >") & CSmartStr(strfilename)
													ResponseWrite ((("<input type=""Hidden"" id=""notempty_" & CSmartStr(cfieldname)) & """ value=""") & CSmartStr(IIF(asp_strlen(value),1,0))) & """>"
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
		end if
	end if
	if bValue(asp_count(ArrayElement(validate,"basicValidate"))) and not IsFalse(asp_array_search("IsRequired",ArrayElement(validate,"basicValidate"),false)) then
		ResponseWrite "&nbsp;<font color=""red"">*</font></span>"
	else
		ResponseWrite "</span>"
	end if
End Function
Function my_stripos(ByVal str,ByVal needle,ByVal offest)
	if IsEqual(asp_strlen(needle),0) or IsEqual(asp_strlen(str),0) then
		my_stripos = false
		Exit Function
	end if
	doAssignmentByRef my_stripos,asp_strpos(asp_strtolower(str),asp_strtolower(needle),offest)
	Exit Function
End Function
Function my_str_ireplace(ByVal search,ByVal replace,ByVal str)
	Dim pos
	doAssignmentByRef pos,my_stripos(str,search,0)
	if IsFalse(pos) then
		doAssignmentByRef my_str_ireplace,str
		Exit Function
	end if
	my_str_ireplace = (CSmartStr(asp_substr(str,0,pos)) & CSmartStr(replace)) & CSmartStr(asp_substr(str,CSmartDbl(pos)+CSmartDbl(asp_strlen(search)),empty))
	Exit Function
End Function
Function in_assoc_array(ByVal name,ByVal arr)
	Dim key
	GetCollectionBounds arr,i_commonfunctions_loopIdx40,i_commonfunctions_loopMax40
	do while i_commonfunctions_loopIdx40<=i_commonfunctions_loopMax40
		key = GetCollectionKey(arr,i_commonfunctions_loopIdx40)
		doAssignment value,ArrayElement(arr,key)
		if IsEqual(key,name) then
			in_assoc_array = true
			Exit Function
		end if
		i_commonfunctions_loopIdx40=i_commonfunctions_loopIdx40+1
	loop
	in_assoc_array = false
	Exit Function
End Function
Function buildLookupSQL(ByVal field,ByVal table,ByVal parentVal,ByVal childVal,ByVal doCategoryFilter,ByVal doValueFilter,ByVal addCategoryField,ByVal doWhereFilter,ByVal oneRecordMode)
	Dim nLookupType,bUnique,strLookupWhere,strOrderBy,bDesc,strCategoryFilter,LookupSQL,categoryWhere,childWhere,condition,strWhere
	if not bValue(asp_strlen(table)) then
		doAssignment table,strTableName
	end if
	doAssignmentByRef nLookupType,GetFieldData(table,field,"LookupType",LT_LISTOFVALUES)
	if not IsEqual(nLookupType,LT_LOOKUPTABLE) then
		buildLookupSQL = ""
		Exit Function
	end if
	doAssignmentByRef bUnique,GetFieldData(table,field,"LookupUnique",false)
	doAssignmentByRef strLookupWhere,LookupWhere(field,table)
	doAssignmentByRef strOrderBy,GetFieldData(table,field,"LookupOrderBy","")
	doAssignmentByRef bDesc,GetFieldData(table,field,"LookupDesc",false)
	doAssignmentByRef strCategoryFilter,GetFieldData(table,field,"CategoryFilter","")
	if bValue(doCategoryFilter) then
		doAssignmentByRef parentVal,make_db_value(CategoryControl(field,table),parentVal,"","","")
	end if
	if bValue(doValueFilter) then
		doAssignmentByRef childVal,make_db_value(field,childVal,"","","")
	end if
	LookupSQL = "SELECT "
	if bValue(oneRecordMode) then
		LookupSQL = CSmartStr(LookupSQL) & "top 1 "
	end if
	if bValue(bUnique) then
		LookupSQL = CSmartStr(LookupSQL) & "DISTINCT "
	end if
	LookupSQL = CSmartStr(LookupSQL) & CSmartStr(GetLWLinkField(field,table,true))
	LookupSQL = CSmartStr(LookupSQL) & ("," & CSmartStr(GetLWDisplayField(field,table,true)))
	if bValue(addCategoryField) and bValue(asp_strlen(strCategoryFilter)) then
		LookupSQL = CSmartStr(LookupSQL) & ("," & CSmartStr(AddFieldWrappers(strCategoryFilter)))
	end if
	LookupSQL = CSmartStr(LookupSQL) & (" FROM " & CSmartStr(AddTableWrappers(GetLookupTable(field,table))))
	categoryWhere = ""
	childWhere = ""
	if bValue(UseCategory(field,table)) and bValue(doCategoryFilter) then
		condition = "=" & CSmartStr(parentVal)
		if IsIdentical(childVal,"null") then
			condition = " is null"
		end if
		categoryWhere = CSmartStr(AddFieldWrappers(strCategoryFilter)) & CSmartStr(condition)
	end if
	if bValue(doValueFilter) then
		condition = "=" & CSmartStr(childVal)
		if IsIdentical(childVal,"null") then
			condition = " is null"
		end if
		childWhere = CSmartStr(AddFieldWrappers(GetLWLinkField(field,table,true))) & CSmartStr(condition)
	end if
	strWhere = ""
	if bValue(doWhereFilter) and bValue(asp_strlen(strLookupWhere)) then
		strWhere = ("(" & CSmartStr(strLookupWhere)) & ")"
	end if
	if bValue(asp_strlen(categoryWhere)) then
		if bValue(asp_strlen(strWhere)) then
			strWhere = CSmartStr(strWhere) & " AND "
		end if
		strWhere = CSmartStr(strWhere) & CSmartStr(categoryWhere)
	end if
	if bValue(asp_strlen(childWhere)) then
		if bValue(asp_strlen(strWhere)) then
			strWhere = CSmartStr(strWhere) & " AND "
		end if
		strWhere = CSmartStr(strWhere) & CSmartStr(childWhere)
	end if
	if bValue(asp_strlen(strWhere)) then
		LookupSQL = CSmartStr(LookupSQL) & (" WHERE " & CSmartStr(strWhere))
	end if
	if bValue(asp_strlen(strOrderBy)) then
		LookupSQL = CSmartStr(LookupSQL) & (((" ORDER BY " & CSmartStr(AddTableWrappers(GetLookupTable(field,table)))) & ".") & CSmartStr(AddFieldWrappers(strOrderBy)))
		if bValue(bDesc) then
			LookupSQL = CSmartStr(LookupSQL) & " DESC"
		end if
	end if
	doAssignmentByRef buildLookupSQL,LookupSQL
	Exit Function
End Function
Function loadSelectContent(ByVal childFieldName,ByVal parentVal,ByVal doCategoryFilter,ByVal childVal,ByVal initialLoad)
	Dim table,Lookup,var_response,output,rs,data
	doAssignment table,strTableName
	Lookup = ""
	Set var_response = (CreateDictionary())
	output = ""
	doAssignmentByRef LookupSQL,buildLookupSQL(childFieldName,table,parentVal,childVal,doCategoryFilter,bValue(FastType(childFieldName,table)) and bValue(initialLoad),false,true,false)
	doAssignmentByRef rs,db_query(LookupSQL,conn)
	if not bValue(FastType(childFieldName,table)) then
		do while bValue(doAssignmentByRef(data,db_fetch_numarray(rs)))
			setArrElement var_response,asp_count(var_response),ArrayElement(data,0)
			setArrElement var_response,asp_count(var_response),ArrayElement(data,1)
		loop
	else
		doAssignmentByRef data,db_fetch_numarray(rs)
		if bValue(data) and (bValue(asp_strlen(childVal)) or not bValue(db_fetch_numarray(rs))) then
			setArrElement var_response,asp_count(var_response),ArrayElement(data,0)
			setArrElement var_response,asp_count(var_response),ArrayElement(data,1)
		end if
	end if
	doAssignmentByRef loadSelectContent,var_response
	Exit Function
End Function
Function xmlencode(ByVal str)
	doAssignmentByRef str,asp_str_replace("&","&amp;",str)
	doAssignmentByRef str,asp_str_replace("<","&lt;",str)
	doAssignmentByRef str,asp_str_replace(">","&gt;",str)
	doAssignmentByRef str,asp_str_replace("'","&apos;",str)
	doAssignmentByRef xmlencode,escapeEntities(str)
	Exit Function
End Function
Function print_inline_array(ByRef arr,ByVal printkey)
	Dim val,key
	if not bValue(printkey) then
		GetCollectionBounds arr,i_commonfunctions_loopIdx42,i_commonfunctions_loopMax42
		do while i_commonfunctions_loopIdx42<=i_commonfunctions_loopMax42
			key = GetCollectionKey(arr,i_commonfunctions_loopIdx42)
			doAssignment val,ArrayElement(arr,key)
			ResponseWrite CSmartStr(asp_str_replace(CreateDictionary5(Empty,"&",Empty,"<",Empty,"\",Empty,vbcr,Empty,vblf),CreateDictionary5(Empty,"&amp;",Empty,"&lt;",Empty,"\\",Empty,"\r",Empty,"\n"),asp_str_replace(CreateDictionary3(Empty,"\",Empty,vbcr,Empty,vblf),CreateDictionary3(Empty,"\\",Empty,"\r",Empty,"\n"),val))) & "\n"
			i_commonfunctions_loopIdx42=i_commonfunctions_loopIdx42+1
		loop
	else
		GetCollectionBounds arr,i_commonfunctions_loopIdx43,i_commonfunctions_loopMax43
		do while i_commonfunctions_loopIdx43<=i_commonfunctions_loopMax43
			key = GetCollectionKey(arr,i_commonfunctions_loopIdx43)
			doAssignment val,ArrayElement(arr,key)
			ResponseWrite CSmartStr(asp_str_replace(CreateDictionary5(Empty,"&",Empty,"<",Empty,"\",Empty,vbcr,Empty,vblf),CreateDictionary5(Empty,"&amp;",Empty,"&lt;",Empty,"\\",Empty,"\r",Empty,"\n"),asp_str_replace(CreateDictionary3(Empty,"\",Empty,vbcr,Empty,vblf),CreateDictionary3(Empty,"\\",Empty,"\r",Empty,"\n"),key))) & "\n"
			i_commonfunctions_loopIdx43=i_commonfunctions_loopIdx43+1
		loop
	end if
End Function
Function GetChartXML(ByVal chartname)
	Dim strTableName
	doAssignmentByRef strTableName,GetTableByShort(chartname)
	doAssignmentByRef GetChartXML,GetTableData(strTableName,".chartXml","")
	Exit Function
End Function
Function GetSiteUrl()
	Dim url,var_SERVER
	url = "http://" & CSmartStr(GetRequestValue(Request.ServerVariables,"SERVER_NAME"))
	if not IsEqual(GetRequestValue(Request.ServerVariables,"SERVER_PORT"),80) then
		if IsEqual(GetRequestValue(Request.ServerVariables,"SERVER_PORT"),443) then
			url = "https://" & CSmartStr(GetRequestValue(Request.ServerVariables,"SERVER_NAME"))
		else
			url = CSmartStr(url) & (":" & CSmartStr(GetRequestValue(Request.ServerVariables,"SERVER_PORT")))
		end if
	end if
	doAssignmentByRef GetSiteUrl,url
	Exit Function
End Function
Function GetAuditObject(ByVal table)
	GetAuditObject = null
	Exit Function
	if bValue(GetTableData(table,".audit",false)) or not bValue(table) then
		asp_include getabspath("include/audit.asp"),true
	else
		GetAuditObject = null
		Exit Function
	end if
End Function
Function GetLockingObject(ByVal table)
	GetLockingObject = null
	Exit Function
	if not bValue(table) then
		doAssignment table,strTableName
	end if
	if bValue(GetTableData(table,".locking",false)) then
		asp_include getabspath("include/locking.asp"),true
		doAssignmentByRef GetLockingObject,CreateClass("oLocking",0,Empty,Empty,Empty,Empty,Empty,Empty,Empty)
		Exit Function
	else
		GetLockingObject = null
		Exit Function
	end if
End Function
Function isEnableSection508()
	isEnableSection508 = false
	Exit Function
End Function
Function isEnableUpper(ByVal val)
	if bValue(ArrayElement(ArrayElement(tables_data,strTableName),".NCSearch")) then
		doAssignmentByRef isEnableUpper,db_upper(val)
		Exit Function
	else
		doAssignmentByRef isEnableUpper,val
		Exit Function
	end if
End Function
Function getJsValidatorName(ByVal name)
	do
		If IsEqual(name,"Number") then
			getJsValidatorName = "IsNumeric"
			Exit Function
			exit do
		End If
		If IsEqual(name,"Number") or IsEqual(name,"Password") then
			getJsValidatorName = "IsPassword"
			Exit Function
			exit do
		End If
		If IsEqual(name,"Number") or IsEqual(name,"Password") or IsEqual(name,"Email") then
			getJsValidatorName = "IsEmail"
			Exit Function
			exit do
		End If
		If IsEqual(name,"Number") or IsEqual(name,"Password") or IsEqual(name,"Email") or IsEqual(name,"Currency") then
			getJsValidatorName = "IsMoney"
			Exit Function
			exit do
		End If
		If IsEqual(name,"Number") or IsEqual(name,"Password") or IsEqual(name,"Email") or IsEqual(name,"Currency") or IsEqual(name,"US ZIP Code") then
			getJsValidatorName = "IsZipCode"
			Exit Function
			exit do
		End If
		If IsEqual(name,"Number") or IsEqual(name,"Password") or IsEqual(name,"Email") or IsEqual(name,"Currency") or IsEqual(name,"US ZIP Code") or IsEqual(name,"US Phone Number") then
			getJsValidatorName = "IsPhoneNumber"
			Exit Function
			exit do
		End If
		If IsEqual(name,"Number") or IsEqual(name,"Password") or IsEqual(name,"Email") or IsEqual(name,"Currency") or IsEqual(name,"US ZIP Code") or IsEqual(name,"US Phone Number") or IsEqual(name,"US State") then
			getJsValidatorName = "IsState"
			Exit Function
			exit do
		End If
		If IsEqual(name,"Number") or IsEqual(name,"Password") or IsEqual(name,"Email") or IsEqual(name,"Currency") or IsEqual(name,"US ZIP Code") or IsEqual(name,"US Phone Number") or IsEqual(name,"US State") or IsEqual(name,"US SSN") then
			getJsValidatorName = "IsSSN"
			Exit Function
			exit do
		End If
		If IsEqual(name,"Number") or IsEqual(name,"Password") or IsEqual(name,"Email") or IsEqual(name,"Currency") or IsEqual(name,"US ZIP Code") or IsEqual(name,"US Phone Number") or IsEqual(name,"US State") or IsEqual(name,"US SSN") or IsEqual(name,"Credit Card") then
			getJsValidatorName = "IsCC"
			Exit Function
			exit do
		End If
		If IsEqual(name,"Number") or IsEqual(name,"Password") or IsEqual(name,"Email") or IsEqual(name,"Currency") or IsEqual(name,"US ZIP Code") or IsEqual(name,"US Phone Number") or IsEqual(name,"US State") or IsEqual(name,"US SSN") or IsEqual(name,"Credit Card") or IsEqual(name,"Time") then
			getJsValidatorName = "IsTime"
			Exit Function
			exit do
		End If
		If IsEqual(name,"Number") or IsEqual(name,"Password") or IsEqual(name,"Email") or IsEqual(name,"Currency") or IsEqual(name,"US ZIP Code") or IsEqual(name,"US Phone Number") or IsEqual(name,"US State") or IsEqual(name,"US SSN") or IsEqual(name,"Credit Card") or IsEqual(name,"Time") or IsEqual(name,"Regular expression") then
			getJsValidatorName = "RegExp"
			Exit Function
			exit do
		End If
		doAssignmentByRef getJsValidatorName,name
		Exit Function
		exit do
	Loop While false
End Function
Function GetInputElementId(ByVal field,ByVal id)
	Dim format,var_type,lookuptype
	doAssignmentByRef format,GetEditFormat(field,"")
	if IsEqual(format,EDIT_FORMAT_DATE) then
		doAssignmentByRef var_type,DateEditType(field,"")
		if IsEqual(var_type,EDIT_DATE_DD) or IsEqual(var_type,EDIT_DATE_DD_DP) then
			GetInputElementId = (("dayvalue_" & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)
			Exit Function
		else
			GetInputElementId = (("value_" & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)
			Exit Function
		end if
	else
		if IsEqual(format,EDIT_FORMAT_RADIO) then
			GetInputElementId = ((("radio_" & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)) & "_0"
			Exit Function
		else
			if IsEqual(format,EDIT_FORMAT_LOOKUP_WIZARD) then
				doAssignmentByRef lookuptype,LookupControlType(field,"")
				if IsEqual(lookuptype,LCT_AJAX) or IsEqual(lookuptype,LCT_LIST) then
					GetInputElementId = (("display_value_" & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)
					Exit Function
				else
					GetInputElementId = (("value_" & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)
					Exit Function
				end if
			else
				GetInputElementId = (("value_" & CSmartStr(GoodFieldName(field))) & "_") & CSmartStr(id)
				Exit Function
			end if
		end if
	end if
End Function
Function SetLangVars(ByVal links)
	Dim var_REQUEST,var_SESSION,var,is508
	xt.assign_p2 "lang_label",true
	if bValue(GetRequestValue(Request,"language")) then
		setArrElement Session,"language",GetRequestValue(Request,"language")
	end if
	var = CSmartStr(GoodFieldName(mlang_getcurrentlang())) & "_langattrs"
	xt.assign_p2 var,"selected"
	doAssignmentByRef is508,isEnableSection508()
	if bValue(is508) then
		xt.assign_section_p3 "lang_label","<label for=""lang"">","</label>"
	end if
	xt.assign_p2 "langselector_attrs",((("name=lang " & CSmartStr(IIF(IsEqual(is508,true),"id=""lang"" ",""))) & "onchange=""javascript: window.location='") & CSmartStr(links)) & ".asp?language='+this.options[this.selectedIndex].value"""
End Function
Function GetTableCaption(ByVal table)
	doAssignmentByRef GetTableCaption,ArrayElement(ArrayElement(tableCaptions,mlang_getcurrentlang()),table)
	Exit Function
End Function
Function GetFieldByLabel(ByVal table,ByVal label)
	Dim currLang,lables,val,key
	if not bValue(table) then
		doAssignment table,strTableName
	end if
	if not bValue(asp_array_key_exists(table,field_labels)) then
		GetFieldByLabel = ""
		Exit Function
	end if
	doAssignmentByRef currLang,mlang_getcurrentlang()
	if not bValue(asp_array_key_exists(currLang,ArrayElement(field_labels,table))) then
		GetFieldByLabel = ""
		Exit Function
	end if
	doAssignment lables,ArrayElement(ArrayElement(field_labels,table),mlang_getcurrentlang())
	GetCollectionBounds lables,i_commonfunctions_loopIdx45,i_commonfunctions_loopMax45
	do while i_commonfunctions_loopIdx45<=i_commonfunctions_loopMax45
		key = GetCollectionKey(lables,i_commonfunctions_loopIdx45)
		doAssignment val,ArrayElement(lables,key)
		if IsEqual(val,label) then
			doAssignmentByRef GetFieldByLabel,key
			Exit Function
		end if
		i_commonfunctions_loopIdx45=i_commonfunctions_loopIdx45+1
	loop
	GetFieldByLabel = ""
	Exit Function
End Function
Function GetFieldLabel(ByVal table,ByVal field)
	if not bValue(asp_array_key_exists(table,field_labels)) then
		GetFieldLabel = ""
		Exit Function
	end if
	doAssignmentByRef GetFieldLabel,ArrayElement(ArrayElement(ArrayElement(field_labels,table),mlang_getcurrentlang()),field)
	Exit Function
End Function
Function GetFieldToolTip(ByVal table,ByVal field)
	if not bValue(asp_array_key_exists(table,fieldToolTips)) then
		GetFieldToolTip = ""
		Exit Function
	end if
	doAssignmentByRef GetFieldToolTip,ArrayElement(ArrayElement(ArrayElement(fieldToolTips,table),mlang_getcurrentlang()),field)
	Exit Function
End Function
Function GetCustomLabel(ByVal custom)
	doAssignmentByRef GetCustomLabel,ArrayElement(ArrayElement(custom_labels,mlang_getcurrentlang()),custom)
	Exit Function
End Function
Function mlang_getcurrentlang()
	Dim var_REQUEST,var_SESSION
	if bValue(GetRequestValue(Request,"language")) then
		setArrElement Session,"language",GetRequestValue(Request,"language")
	end if
	if bValue(Session("language")) then
		doAssignmentByRef mlang_getcurrentlang,Session("language")
		Exit Function
	end if
	doAssignmentByRef mlang_getcurrentlang,mlang_defaultlang
	Exit Function
End Function
Function mlang_getlanglist()
	doAssignmentByRef mlang_getlanglist,asp_array_keys(mlang_messages,empty)
	Exit Function
End Function
Function displayDetailsOn(ByVal table,ByVal page)
	Dim key,i
	if not (not IsEmpty(ArrayElement(detailsTablesData,table))) and not bValue(asp_is_array(ArrayElement(detailsTablesData,table))) then
		displayDetailsOn = false
		Exit Function
	end if
	if IsEqual(page,PAGE_EDIT) then
		key = "previewOnEdit"
	else
		if IsEqual(page,PAGE_ADD) then
			key = "previewOnAdd"
		else
			if IsEqual(page,PAGE_VIEW) then
				key = "previewOnView"
			else
				key = "previewOnList"
			end if
		end if
	end if
	i = 0
	do while IsLess(i,asp_count(ArrayElement(detailsTablesData,table)))
		if bValue(ArrayElement(ArrayElement(ArrayElement(detailsTablesData,table),i),key)) then
			displayDetailsOn = true
			Exit Function
		end if
		i = CSmartDbl(i)+1
	loop
	displayDetailsOn = false
	Exit Function
End Function
Function showDetailTable(ByVal params)
	Dim oldTableName
	doAssignment oldTableName,strTableName
	doAssignment strTableName,ArrayElement(params,"table")
	if bValue(ArrayElement(params,"dpObject").isDispGrid()) then
		ArrayElement(params,"dpObject").showPage 
	end if
	doAssignment strTableName,oldTableName
End Function
Function DoUpdateRecordSQL(ByVal table,ByRef evalues,ByRef blobfields,ByVal strWhereClause,ByVal pageid,ByRef pageObject)
	Dim strSQL,blobs,ekey,strValues,value,status,IsSaved
	if not bValue(asp_count(evalues)) then
		DoUpdateRecordSQL = true
		Exit Function
	end if
	strSQL = ("update " & CSmartStr(AddTableWrappers(table))) & " set "
	doAssignmentByRef blobs,PrepareBlobs(evalues,blobfields)
	GetCollectionBounds evalues,i_commonfunctions_loopIdx47,i_commonfunctions_loopMax47
	do while i_commonfunctions_loopIdx47<=i_commonfunctions_loopMax47
		ekey = GetCollectionKey(evalues,i_commonfunctions_loopIdx47)
		doAssignment value,ArrayElement(evalues,ekey)
		if bValue(asp_in_array(ekey,blobfields,false)) then
			doAssignment strValues,value
		else
			doAssignmentByRef strValues,add_db_quotes(ekey,value,"")
		end if
		strSQL = CSmartStr(strSQL) & (((CSmartStr(GetFullFieldName(ekey,"")) & "=") & CSmartStr(strValues)) & ", ")
		i_commonfunctions_loopIdx47=i_commonfunctions_loopIdx47+1
	loop
	if IsEqual(asp_substr(strSQL,-2,empty),", ") then
		doAssignmentByRef strSQL,asp_substr(strSQL,0,CSmartDbl(asp_strlen(strSQL))-2)
	end if
	if IsIdentical(strWhereClause,"") then
		strWhereClause = " (1=1) "
	end if
	strSQL = CSmartStr(strSQL) & (" where " & CSmartStr(strWhereClause))
	if bValue(SecuritySQL("Edit","")) then
		strSQL = CSmartStr(strSQL) & ((" and (" & CSmartStr(SecuritySQL("Edit",""))) & ")")
	end if
	if not bValue(ExecuteUpdate(strSQL,blobs,false)) then
		DoUpdateRecordSQL = false
		Exit Function
	end if
	pageObject.ProcessFiles 
	if bValue(inlineedit) then
		status = "UPDATED"
		message = ("" & CSmartStr("Record updated")) & ""
		IsSaved = true
	else
		message = ("<<< " & CSmartStr("Record updated")) & " >>>"
	end if
	if not IsEqual(usermessage,"") then
		doAssignment message,usermessage
	end if
	DoUpdateRecordSQL = true
	Exit Function
End Function
Function DoInsertRecordSQL(ByVal table,ByRef avalues,ByRef blobfields,ByVal pageid,ByRef pageObject)
	Dim strSQL,strFields,strValues,blobs,akey,value,status,IsSaved,auditObj,keyfields,k,lastrs,lastdata,seqname,mystrSQL,myrs,mydata
	strSQL = ("insert into " & CSmartStr(AddTableWrappers(table))) & " "
	strFields = "("
	strValues = "("
	doAssignmentByRef blobs,PrepareBlobs(avalues,blobfields)
	GetCollectionBounds avalues,i_commonfunctions_loopIdx48,i_commonfunctions_loopMax48
	do while i_commonfunctions_loopIdx48<=i_commonfunctions_loopMax48
		akey = GetCollectionKey(avalues,i_commonfunctions_loopIdx48)
		doAssignment value,ArrayElement(avalues,akey)
		strFields = CSmartStr(strFields) & (CSmartStr(GetFullFieldName(akey,"")) & ", ")
		if bValue(asp_in_array(akey,blobfields,false)) then
			strValues = CSmartStr(strValues) & (CSmartStr(value) & ", ")
		else
			strValues = CSmartStr(strValues) & (CSmartStr(add_db_quotes(akey,value,"")) & ", ")
		end if
		i_commonfunctions_loopIdx48=i_commonfunctions_loopIdx48+1
	loop
	if IsEqual(asp_substr(strFields,-2,empty),", ") then
		doAssignmentByRef strFields,asp_substr(strFields,0,CSmartDbl(asp_strlen(strFields))-2)
	end if
	if IsEqual(asp_substr(strValues,-2,empty),", ") then
		doAssignmentByRef strValues,asp_substr(strValues,0,CSmartDbl(asp_strlen(strValues))-2)
	end if
	strSQL = CSmartStr(strSQL) & (((CSmartStr(strFields) & ") values ") & CSmartStr(strValues)) & ")")
	if not bValue(ExecuteUpdate(strSQL,blobs,true)) then
		DoInsertRecordSQL = false
		Exit Function
	end if
	if bValue(error_happened) then
		DoInsertRecordSQL = false
		Exit Function
	end if
	pageObject.ProcessFiles 
	if IsEqual(inlineadd,ADD_INLINE) then
		status = "ADDED"
		message = ("" & CSmartStr("Record was added")) & ""
		IsSaved = true
	else
		message = ("<<< " & CSmartStr("Record was added")) & " >>>"
	end if
	if not IsEqual(usermessage,"") then
		doAssignment message,usermessage
	end if
	doAssignmentByRef auditObj,GetAuditObject(table)
	if (((((IsEqual(inlineadd,ADD_SIMPLE) or IsEqual(inlineadd,ADD_INLINE)) or IsEqual(inlineadd,ADD_ONTHEFLY)) or IsEqual(inlineadd,ADD_POPUP)) or IsEqual(inlineadd,ADD_MASTER)) or bValue(tableEventExists("AfterAdd",strTableName))) or bValue(auditObj) then
		failed_inline_add = false
		doAssignmentByRef keyfields,GetTableKeys("")
		GetCollectionBounds keyfields,i_commonfunctions_loopIdx49,i_commonfunctions_loopMax49
		do while i_commonfunctions_loopIdx49<=i_commonfunctions_loopMax49
			i_commonfunctions_arrKey49 = GetCollectionKey(keyfields,i_commonfunctions_loopIdx49)
			doAssignment k,ArrayElement(keyfields,i_commonfunctions_arrKey49)
			if bValue(asp_array_key_exists(k,avalues)) then
				setArrElement keys,k,ArrayElement(avalues,k)
			else
				if bValue(IsAutoincField(k,"")) then
					doAssignmentByRef lastrs,db_query("select @@IDENTITY",conn)
					if bValue(doAssignmentByRef(lastdata,db_fetch_numarray(lastrs))) then
						setArrElement keys,k,ArrayElement(lastdata,0)
					end if
				else
					failed_inline_add = true
				end if
			end if
			i_commonfunctions_loopIdx49=i_commonfunctions_loopIdx49+1
		loop
	end if
	DoInsertRecordSQL = true
	Exit Function
End Function
Function getEventObject(ByVal table)
	Dim ret
	ret = null
	if not bValue(asp_array_key_exists(table,tableEvents)) then
		doAssignmentByRef getEventObject,ret
		Exit Function
	end if
	doAssignmentByRef getEventObject,ArrayElement(tableEvents,table)
	Exit Function
End Function
Function tableEventExists(ByVal var_event,ByVal table)
	if not bValue(asp_array_key_exists(table,tableEvents)) then
		tableEventExists = false
		Exit Function
	end if
	doAssignmentByRef tableEventExists,ArrayElement(tableEvents,table).exists_p1(var_event)
	Exit Function
End Function
Function add_nocache_headers()
	asp_header "Cache-Control: no-cache, no-store, max-age=0, must-revalidate"
	asp_header "Pragma: no-cache"
	asp_header "Expires: Fri, 01 Jan 1990 00:00:00 GMT"
End Function
Function IsGuidString(ByVal str)
	Dim i,c
	if not IsEqual(asp_strlen(str),38) then
		IsGuidString = false
		Exit Function
	end if
	i = 0
	do while IsLess(i,38)
		doAssignmentByRef c,asp_substr(str,i,1)
		if IsEqual(i,0) then
			if not IsEqual(c,"{") then
				IsGuidString = false
				Exit Function
			end if
		else
			if IsEqual(i,37) then
				if not IsEqual(c,"}") then
					IsGuidString = false
					Exit Function
				end if
			else
				if ((IsEqual(i,9) or IsEqual(i,14)) or IsEqual(i,19)) or IsEqual(i,24) then
					if not IsEqual(c,"-") then
						IsGuidString = false
						Exit Function
					end if
				else
					if ((IsLess(c,"0") or IsLess("9",c)) and (IsLess(c,"a") or IsLess("f",c))) and (IsLess(c,"A") or IsLess("F",c)) then
						IsGuidString = false
						Exit Function
					end if
				end if
			end if
		end if
		i = CSmartDbl(i)+1
	loop
	IsGuidString = true
	Exit Function
End Function
Function IsStoredProcedure(ByVal strSQL)
	Dim c
	if IsLess(6,asp_strlen(strSQL)) then
		doAssignmentByRef c,asp_strtolower(asp_substr(strSQL,6,1))
		if ((IsEqual(asp_strtolower(asp_substr(strSQL,0,6)),"select") and (IsLess(c,"0") or IsLess("9",c))) and (IsLess(c,"a") or IsLess("z",c))) and not IsEqual(c,"_") then
			IsStoredProcedure = false
			Exit Function
		else
			IsStoredProcedure = true
			Exit Function
		end if
	else
		IsStoredProcedure = true
		Exit Function
	end if
End Function
Function CreateCKeditor(ByVal cfield,ByVal value)
	ResponseWrite ((((("<textarea id=""" & CSmartStr(cfield)) & """ name=""") & CSmartStr(cfield)) & """ rows=""8"" cols=""60"">") & CSmartStr(htmlspecialchars(value))) & "</textarea>"
End Function
%>
