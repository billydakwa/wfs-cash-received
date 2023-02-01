<%
Set tdatapayment = (CreateDictionary())
setArrElement tdatapayment,".NumberOfChars",80
setArrElement tdatapayment,".ShortName","payment"
setArrElement tdatapayment,".OwnerID",""
setArrElement tdatapayment,".OriginalTable","payment"
Set fieldLabelspayment = (CreateDictionary())
if IsEqual(mlang_getcurrentlang(),"English") then
	setArrElement fieldLabelspayment,"English",CreateDictionary()
	setArrElement fieldToolTipspayment,"English",CreateDictionary()
	setArrElementN fieldLabelspayment,CreateArray2("English","station"),"Station"
	setArrElementN fieldToolTipspayment,CreateArray2("English","station"),""
	setArrElementN fieldLabelspayment,CreateArray2("English","payment_date"),"Payment Date"
	setArrElementN fieldToolTipspayment,CreateArray2("English","payment_date"),""
	setArrElementN fieldLabelspayment,CreateArray2("English","fop"),"Fop"
	setArrElementN fieldToolTipspayment,CreateArray2("English","fop"),""
	setArrElementN fieldLabelspayment,CreateArray2("English","name"),"Name"
	setArrElementN fieldToolTipspayment,CreateArray2("English","name"),""
	setArrElementN fieldLabelspayment,CreateArray2("English","prefix"),"Prefix"
	setArrElementN fieldToolTipspayment,CreateArray2("English","prefix"),""
	setArrElementN fieldLabelspayment,CreateArray2("English","serial"),"Serial"
	setArrElementN fieldToolTipspayment,CreateArray2("English","serial"),""
	setArrElementN fieldLabelspayment,CreateArray2("English","amount_received"),"Amount Received"
	setArrElementN fieldToolTipspayment,CreateArray2("English","amount_received"),""
	if bValue(asp_count(ArrayElement(fieldToolTipspayment,"English"))) then
		setArrElement tdatapayment,".isUseToolTips",true
	end if
end if
setArrElement tdatapayment,".NCSearch",true
setArrElement tdatapayment,".shortTableName","payment"
setArrElement tdatapayment,".nSecOptions",0
setArrElement tdatapayment,".recsPerRowList",1
setArrElement tdatapayment,".tableGroupBy","0"
setArrElement tdatapayment,".mainTableOwnerID",""
setArrElement tdatapayment,".moveNext",1
setArrElement tdatapayment,".showAddInPopup",false
setArrElement tdatapayment,".showEditInPopup",false
setArrElement tdatapayment,".showViewInPopup",false
setArrElement tdatapayment,".fieldsForRegister",CreateDictionary()
setArrElement tdatapayment,".listAjax",false
setArrElement tdatapayment,".audit",false
setArrElement tdatapayment,".locking",false
setArrElement tdatapayment,".listIcons",true
setArrElement tdatapayment,".exportTo",true
setArrElement tdatapayment,".showSimpleSearchOptions",false
setArrElement tdatapayment,".showSearchPanel",true
setArrElement tdatapayment,".isUseAjaxSuggest",true
setArrElement tdatapayment,".rowHighlite",true
setArrElement tdatapayment,".addPageEvents",false
setArrElement tdatapayment,".isUseCalendarForSearch",true
setArrElement tdatapayment,".isUseTimeForSearch",false
setArrElement tdatapayment,".isUseiBox",false
setArrElement tdatapayment,".isUseInlineJs",bValue(ArrayElement(tdatapayment,".isUseInlineAdd")) or bValue(ArrayElement(tdatapayment,".isUseInlineEdit"))
setArrElement tdatapayment,".allSearchFields",CreateDictionary()
setArrElementN tdatapayment,CreateArray2(".globSearchFields",empty),"payment_date"
if not bValue(asp_in_array("payment_date",ArrayElement(tdatapayment,".allSearchFields"),false)) then
	setArrElementN tdatapayment,CreateArray2(".allSearchFields",empty),"payment_date"
end if
setArrElementN tdatapayment,CreateArray2(".globSearchFields",empty),"fop"
if not bValue(asp_in_array("fop",ArrayElement(tdatapayment,".allSearchFields"),false)) then
	setArrElementN tdatapayment,CreateArray2(".allSearchFields",empty),"fop"
end if
setArrElementN tdatapayment,CreateArray2(".googleLikeFields",empty),"name"
setArrElementN tdatapayment,CreateArray2(".googleLikeFields",empty),"serial"
setArrElementN tdatapayment,CreateArray2(".googleLikeFields",empty),"amount_received"
setArrElementN tdatapayment,CreateArray2(".googleLikeFields",empty),"prefix"
setArrElementN tdatapayment,CreateArray2(".googleLikeFields",empty),"station"
setArrElementN tdatapayment,CreateArray2(".googleLikeFields",empty),"payment_date"
setArrElementN tdatapayment,CreateArray2(".googleLikeFields",empty),"fop"
setArrElementN tdatapayment,CreateArray2(".advSearchFields",empty),"name"
if not bValue(asp_in_array("name",ArrayElement(tdatapayment,".allSearchFields"),false)) then
	setArrElementN tdatapayment,CreateArray2(".allSearchFields",empty),"name"
end if
setArrElementN tdatapayment,CreateArray2(".advSearchFields",empty),"serial"
if not bValue(asp_in_array("serial",ArrayElement(tdatapayment,".allSearchFields"),false)) then
	setArrElementN tdatapayment,CreateArray2(".allSearchFields",empty),"serial"
end if
setArrElementN tdatapayment,CreateArray2(".advSearchFields",empty),"amount_received"
if not bValue(asp_in_array("amount_received",ArrayElement(tdatapayment,".allSearchFields"),false)) then
	setArrElementN tdatapayment,CreateArray2(".allSearchFields",empty),"amount_received"
end if
setArrElementN tdatapayment,CreateArray2(".advSearchFields",empty),"prefix"
if not bValue(asp_in_array("prefix",ArrayElement(tdatapayment,".allSearchFields"),false)) then
	setArrElementN tdatapayment,CreateArray2(".allSearchFields",empty),"prefix"
end if
setArrElementN tdatapayment,CreateArray2(".advSearchFields",empty),"station"
if not bValue(asp_in_array("station",ArrayElement(tdatapayment,".allSearchFields"),false)) then
	setArrElementN tdatapayment,CreateArray2(".allSearchFields",empty),"station"
end if
setArrElementN tdatapayment,CreateArray2(".advSearchFields",empty),"payment_date"
if not bValue(asp_in_array("payment_date",ArrayElement(tdatapayment,".allSearchFields"),false)) then
	setArrElementN tdatapayment,CreateArray2(".allSearchFields",empty),"payment_date"
end if
setArrElementN tdatapayment,CreateArray2(".advSearchFields",empty),"fop"
if not bValue(asp_in_array("fop",ArrayElement(tdatapayment,".allSearchFields"),false)) then
	setArrElementN tdatapayment,CreateArray2(".allSearchFields",empty),"fop"
end if
setArrElement tdatapayment,".isTableType","list"
setArrElement tdatapayment,".isDisplayLoading",true
setArrElement tdatapayment,".isResizeColumns",false
setArrElement tdatapayment,".pageSize",20
gstrOrderBy = "order by payment.payment_date"
if bValue(asp_strlen(gstrOrderBy)) and not IsEqual(asp_strtolower(asp_substr(gstrOrderBy,0,8)),"order by") then
	gstrOrderBy = "order by " & CSmartStr(gstrOrderBy)
end if
setArrElement tdatapayment,".strOrderBy",gstrOrderBy
setArrElement tdatapayment,".orderindexes",CreateDictionary()
setArrElementN tdatapayment,CreateArray2(".orderindexes",empty),CreateDictionary3(Empty,3,Empty,IIF(1,"ASC","DESC"),Empty,"payment.payment_date")
setArrElement tdatapayment,".sqlHead","SELECT payment.station,  part.name,  payment.payment_date,  payment.fop,  awb.[prefix],  awb.serial,  payment.amount_received"
setArrElement tdatapayment,".sqlFrom","FROM payment  INNER JOIN part ON payment.part = part.code  INNER JOIN awb ON payment.awb_seq = awb.seq"
setArrElement tdatapayment,".sqlWhereExpr","(payment.payment_date between '2022-12-31' and GETDATE())"
setArrElement tdatapayment,".sqlTail",""
Set arrRPP = (CreateDictionary())
setArrElement arrRPP,asp_count(arrRPP),10
setArrElement arrRPP,asp_count(arrRPP),20
setArrElement arrRPP,asp_count(arrRPP),30
setArrElement arrRPP,asp_count(arrRPP),50
setArrElement arrRPP,asp_count(arrRPP),100
setArrElement arrRPP,asp_count(arrRPP),500
setArrElement arrRPP,asp_count(arrRPP),-1
setArrElement tdatapayment,".arrRecsPerPage",arrRPP
Set arrGPP = (CreateDictionary())
setArrElement arrGPP,asp_count(arrGPP),1
setArrElement arrGPP,asp_count(arrGPP),3
setArrElement arrGPP,asp_count(arrGPP),5
setArrElement arrGPP,asp_count(arrGPP),10
setArrElement arrGPP,asp_count(arrGPP),50
setArrElement arrGPP,asp_count(arrGPP),100
setArrElement arrGPP,asp_count(arrGPP),-1
setArrElement tdatapayment,".arrGroupsPerPage",arrGPP
Set tableKeys = (CreateDictionary())
setArrElement tdatapayment,".Keys",tableKeys
Set fdata = (CreateDictionary())
setArrElement fdata,"strName","station"
setArrElement fdata,"ownerTable","payment"
setArrElement fdata,"Label","Station"
setArrElement fdata,"FieldType",129
setArrElement fdata,"UseiBox",false
setArrElement fdata,"EditFormat","Text field"
setArrElement fdata,"ViewFormat",""
setArrElement fdata,"NeedEncode",true
setArrElement fdata,"GoodName","station"
setArrElement fdata,"FullName","payment.station"
setArrElement fdata,"Index",1
setArrElement fdata,"EditParams",""
setArrElement fdata,"EditParams",CSmartStr(ArrayElement(fdata,"EditParams")) & " maxlength=3"
setArrElement fdata,"bListPage",true
setArrElement fdata,"bExportPage",true
setArrElement fdata,"validateAs",CreateDictionary()
setArrElement fdata,"FieldPermissions",true
setArrElement tdatapayment,"station",fdata
Set fdata = (CreateDictionary())
setArrElement fdata,"strName","name"
setArrElement fdata,"ownerTable","part"
setArrElement fdata,"Label","Name"
setArrElement fdata,"FieldType",200
setArrElement fdata,"UseiBox",false
setArrElement fdata,"EditFormat","Text field"
setArrElement fdata,"ViewFormat",""
setArrElement fdata,"NeedEncode",true
setArrElement fdata,"GoodName","name"
setArrElement fdata,"FullName","part.name"
setArrElement fdata,"Index",2
setArrElement fdata,"EditParams",""
setArrElement fdata,"bListPage",true
setArrElement fdata,"bExportPage",true
setArrElement fdata,"validateAs",CreateDictionary()
setArrElement fdata,"FieldPermissions",true
setArrElement tdatapayment,"name",fdata
Set fdata = (CreateDictionary())
setArrElement fdata,"strName","payment_date"
setArrElement fdata,"ownerTable","payment"
setArrElement fdata,"Label","Payment Date"
setArrElement fdata,"FieldType",135
setArrElement fdata,"UseiBox",false
setArrElement fdata,"EditFormat","Date"
setArrElement fdata,"ViewFormat","Short Date"
setArrElement fdata,"NeedEncode",true
setArrElement fdata,"GoodName","payment_date"
setArrElement fdata,"FullName","payment.payment_date"
setArrElement fdata,"Index",3
setArrElement fdata,"DateEditType",13
setArrElement fdata,"InitialYearFactor",100
setArrElement fdata,"LastYearFactor",10
setArrElement fdata,"bListPage",true
setArrElement fdata,"bAdvancedSearch",true
setArrElement fdata,"bExportPage",true
setArrElement fdata,"validateAs",CreateDictionary()
setArrElement fdata,"FieldPermissions",true
setArrElement tdatapayment,"payment_date",fdata
Set fdata = (CreateDictionary())
setArrElement fdata,"strName","fop"
setArrElement fdata,"ownerTable","payment"
setArrElement fdata,"Label","Fop"
setArrElement fdata,"FieldType",129
setArrElement fdata,"UseiBox",false
setArrElement fdata,"EditFormat","Text field"
setArrElement fdata,"ViewFormat",""
setArrElement fdata,"NeedEncode",true
setArrElement fdata,"GoodName","fop"
setArrElement fdata,"FullName","payment.fop"
setArrElement fdata,"Index",4
setArrElement fdata,"EditParams",""
setArrElement fdata,"EditParams",CSmartStr(ArrayElement(fdata,"EditParams")) & " maxlength=3"
setArrElement fdata,"bListPage",true
setArrElement fdata,"bAdvancedSearch",true
setArrElement fdata,"bExportPage",true
setArrElement fdata,"validateAs",CreateDictionary()
setArrElement fdata,"FieldPermissions",true
setArrElement tdatapayment,"fop",fdata
Set fdata = (CreateDictionary())
setArrElement fdata,"strName","prefix"
setArrElement fdata,"ownerTable","awb"
setArrElement fdata,"Label","Prefix"
setArrElement fdata,"FieldType",129
setArrElement fdata,"UseiBox",false
setArrElement fdata,"EditFormat","Text field"
setArrElement fdata,"ViewFormat",""
setArrElement fdata,"NeedEncode",true
setArrElement fdata,"GoodName","prefix"
setArrElement fdata,"FullName","awb.[prefix]"
setArrElement fdata,"Index",5
setArrElement fdata,"EditParams",""
setArrElement fdata,"bListPage",true
setArrElement fdata,"bExportPage",true
setArrElement fdata,"validateAs",CreateDictionary()
setArrElement fdata,"FieldPermissions",true
setArrElement tdatapayment,"prefix",fdata
Set fdata = (CreateDictionary())
setArrElement fdata,"strName","serial"
setArrElement fdata,"ownerTable","awb"
setArrElement fdata,"Label","Serial"
setArrElement fdata,"FieldType",200
setArrElement fdata,"UseiBox",false
setArrElement fdata,"EditFormat","Text field"
setArrElement fdata,"ViewFormat",""
setArrElement fdata,"NeedEncode",true
setArrElement fdata,"GoodName","serial"
setArrElement fdata,"FullName","awb.serial"
setArrElement fdata,"Index",6
setArrElement fdata,"EditParams",""
setArrElement fdata,"bListPage",true
setArrElement fdata,"bExportPage",true
setArrElement fdata,"validateAs",CreateDictionary()
setArrElement fdata,"FieldPermissions",true
setArrElement tdatapayment,"serial",fdata
Set fdata = (CreateDictionary())
setArrElement fdata,"strName","amount_received"
setArrElement fdata,"ownerTable","payment"
setArrElement fdata,"Label","Amount Received"
setArrElement fdata,"FieldType",131
setArrElement fdata,"UseiBox",false
setArrElement fdata,"EditFormat","Text field"
setArrElement fdata,"ViewFormat","Number"
setArrElement fdata,"DecimalDigits",2
setArrElement fdata,"NeedEncode",true
setArrElement fdata,"GoodName","amount_received"
setArrElement fdata,"FullName","payment.amount_received"
setArrElement fdata,"Index",7
setArrElement fdata,"EditParams",""
setArrElement fdata,"bListPage",true
setArrElement fdata,"bExportPage",true
setArrElement fdata,"validateAs",CreateDictionary()
setArrElementN fdata,CreateArray3("validateAs","basicValidate",empty),getJsValidatorName("Number")
setArrElement fdata,"FieldPermissions",true
setArrElement tdatapayment,"amount_received",fdata
setArrElementByRef tables_data,"payment",tdatapayment
setArrElementByRef field_labels,"payment",fieldLabelspayment
setArrElementByRef fieldToolTips,"payment",fieldToolTipspayment
setArrElement detailsTablesData,"payment",CreateDictionary()
setArrElement masterTablesData,"payment",CreateDictionary()
asp_include getabspath("classes/sql.asp"),true
Set proto0 = (CreateDictionary())
setArrElement proto0,"m_strHead","SELECT"
setArrElement proto0,"m_strFieldList","payment.station,  part.name,  payment.payment_date,  payment.fop,  awb.[prefix],  awb.serial,  payment.amount_received"
setArrElement proto0,"m_strFrom","FROM payment  INNER JOIN part ON payment.part = part.code  INNER JOIN awb ON payment.awb_seq = awb.seq"
setArrElement proto0,"m_strWhere","(payment.payment_date between '2022-12-31' and GETDATE())"
setArrElement proto0,"m_strOrderBy","order by payment.payment_date"
setArrElement proto0,"m_strTail",""
Set proto1 = (CreateDictionary())
setArrElement proto1,"m_sql","payment.payment_date between '2022-12-31' and GETDATE()"
setArrElement proto1,"m_uniontype","SQLL_UNKNOWN"
Set obj = (CreateClass("SQLField",1,CreateDictionary2("m_strName","payment_date","m_strTable","payment"),Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto1,"m_column",obj
setArrElement proto1,"m_contained",CreateDictionary()
setArrElement proto1,"m_strCase","between '2022-12-31' and GETDATE()"
setArrElement proto1,"m_havingmode","0"
setArrElement proto1,"m_inBrackets","0"
setArrElement proto1,"m_useAlias","0"
Set obj = (CreateClass("SQLLogicalExpr",1,proto1,Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto0,"m_where",obj
Set proto3 = (CreateDictionary())
setArrElement proto3,"m_sql",""
setArrElement proto3,"m_uniontype","SQLL_UNKNOWN"
Set obj = (CreateClass("SQLNonParsed",1,CreateDictionary1("m_sql",""),Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto3,"m_column",obj
setArrElement proto3,"m_contained",CreateDictionary()
setArrElement proto3,"m_strCase",""
setArrElement proto3,"m_havingmode","0"
setArrElement proto3,"m_inBrackets","0"
setArrElement proto3,"m_useAlias","0"
Set obj = (CreateClass("SQLLogicalExpr",1,proto3,Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto0,"m_having",obj
setArrElement proto0,"m_fieldlist",CreateDictionary()
Set proto5 = (CreateDictionary())
Set obj = (CreateClass("SQLField",1,CreateDictionary2("m_strName","station","m_strTable","payment"),Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto5,"m_expr",obj
setArrElement proto5,"m_alias",""
Set obj = (CreateClass("SQLFieldListItem",1,proto5,Empty,Empty,Empty,Empty,Empty,Empty))

setArrElementN proto0,CreateArray2("m_fieldlist",empty),obj
Set proto7 = (CreateDictionary())
Set obj = (CreateClass("SQLField",1,CreateDictionary2("m_strName","name","m_strTable","part"),Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto7,"m_expr",obj
setArrElement proto7,"m_alias",""
Set obj = (CreateClass("SQLFieldListItem",1,proto7,Empty,Empty,Empty,Empty,Empty,Empty))

setArrElementN proto0,CreateArray2("m_fieldlist",empty),obj
Set proto9 = (CreateDictionary())
Set obj = (CreateClass("SQLField",1,CreateDictionary2("m_strName","payment_date","m_strTable","payment"),Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto9,"m_expr",obj
setArrElement proto9,"m_alias",""
Set obj = (CreateClass("SQLFieldListItem",1,proto9,Empty,Empty,Empty,Empty,Empty,Empty))

setArrElementN proto0,CreateArray2("m_fieldlist",empty),obj
Set proto11 = (CreateDictionary())
Set obj = (CreateClass("SQLField",1,CreateDictionary2("m_strName","fop","m_strTable","payment"),Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto11,"m_expr",obj
setArrElement proto11,"m_alias",""
Set obj = (CreateClass("SQLFieldListItem",1,proto11,Empty,Empty,Empty,Empty,Empty,Empty))

setArrElementN proto0,CreateArray2("m_fieldlist",empty),obj
Set proto13 = (CreateDictionary())
Set obj = (CreateClass("SQLField",1,CreateDictionary2("m_strName","prefix","m_strTable","awb"),Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto13,"m_expr",obj
setArrElement proto13,"m_alias",""
Set obj = (CreateClass("SQLFieldListItem",1,proto13,Empty,Empty,Empty,Empty,Empty,Empty))

setArrElementN proto0,CreateArray2("m_fieldlist",empty),obj
Set proto15 = (CreateDictionary())
Set obj = (CreateClass("SQLField",1,CreateDictionary2("m_strName","serial","m_strTable","awb"),Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto15,"m_expr",obj
setArrElement proto15,"m_alias",""
Set obj = (CreateClass("SQLFieldListItem",1,proto15,Empty,Empty,Empty,Empty,Empty,Empty))

setArrElementN proto0,CreateArray2("m_fieldlist",empty),obj
Set proto17 = (CreateDictionary())
Set obj = (CreateClass("SQLField",1,CreateDictionary2("m_strName","amount_received","m_strTable","payment"),Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto17,"m_expr",obj
setArrElement proto17,"m_alias",""
Set obj = (CreateClass("SQLFieldListItem",1,proto17,Empty,Empty,Empty,Empty,Empty,Empty))

setArrElementN proto0,CreateArray2("m_fieldlist",empty),obj
setArrElement proto0,"m_fromlist",CreateDictionary()
Set proto19 = (CreateDictionary())
setArrElement proto19,"m_link","SQLL_MAIN"
Set proto20 = (CreateDictionary())
setArrElement proto20,"m_strName","payment"
setArrElement proto20,"m_columns",CreateDictionary()
setArrElementN proto20,CreateArray2("m_columns",empty),"seq"
setArrElementN proto20,CreateArray2("m_columns",empty),"station"
setArrElementN proto20,CreateArray2("m_columns",empty),"currency"
setArrElementN proto20,CreateArray2("m_columns",empty),"part"
setArrElementN proto20,CreateArray2("m_columns",empty),"payment_date"
setArrElementN proto20,CreateArray2("m_columns",empty),"amount"
setArrElementN proto20,CreateArray2("m_columns",empty),"fop"
setArrElementN proto20,CreateArray2("m_columns",empty),"fop_code"
setArrElementN proto20,CreateArray2("m_columns",empty),"fop_number"
setArrElementN proto20,CreateArray2("m_columns",empty),"confirm_number"
setArrElementN proto20,CreateArray2("m_columns",empty),"oper"
setArrElementN proto20,CreateArray2("m_columns",empty),"remark"
setArrElementN proto20,CreateArray2("m_columns",empty),"fop_date"
setArrElementN proto20,CreateArray2("m_columns",empty),"fop_ref"
setArrElementN proto20,CreateArray2("m_columns",empty),"awb_seq"
setArrElementN proto20,CreateArray2("m_columns",empty),"awb_inv_number"
setArrElementN proto20,CreateArray2("m_columns",empty),"time_stamp"
setArrElementN proto20,CreateArray2("m_columns",empty),"exp_imp"
setArrElementN proto20,CreateArray2("m_columns",empty),"account_flag"
setArrElementN proto20,CreateArray2("m_columns",empty),"run_seq"
setArrElementN proto20,CreateArray2("m_columns",empty),"run_inv_number"
setArrElementN proto20,CreateArray2("m_columns",empty),"inv_pay_seq"
setArrElementN proto20,CreateArray2("m_columns",empty),"freight_charge"
setArrElementN proto20,CreateArray2("m_columns",empty),"other_charge"
setArrElementN proto20,CreateArray2("m_columns",empty),"cass_file_header"
setArrElementN proto20,CreateArray2("m_columns",empty),"payment_type"
setArrElementN proto20,CreateArray2("m_columns",empty),"product_inv_seq"
setArrElementN proto20,CreateArray2("m_columns",empty),"cash_box"
setArrElementN proto20,CreateArray2("m_columns",empty),"cass_reason_code"
setArrElementN proto20,CreateArray2("m_columns",empty),"cass_contract_no"
setArrElementN proto20,CreateArray2("m_columns",empty),"vat_amount"
setArrElementN proto20,CreateArray2("m_columns",empty),"csr_no"
setArrElementN proto20,CreateArray2("m_columns",empty),"awb_dlv_seq"
setArrElementN proto20,CreateArray2("m_columns",empty),"carrier"
setArrElementN proto20,CreateArray2("m_columns",empty),"state"
setArrElementN proto20,CreateArray2("m_columns",empty),"driver_license"
setArrElementN proto20,CreateArray2("m_columns",empty),"cash_box_run"
setArrElementN proto20,CreateArray2("m_columns",empty),"original_currency"
setArrElementN proto20,CreateArray2("m_columns",empty),"original_amount"
setArrElementN proto20,CreateArray2("m_columns",empty),"original_vat_amt"
setArrElementN proto20,CreateArray2("m_columns",empty),"exch_rate"
setArrElementN proto20,CreateArray2("m_columns",empty),"son_flag"
setArrElementN proto20,CreateArray2("m_columns",empty),"awb_exp_seq"
setArrElementN proto20,CreateArray2("m_columns",empty),"cca_number"
setArrElementN proto20,CreateArray2("m_columns",empty),"receipt_number"
setArrElementN proto20,CreateArray2("m_columns",empty),"withold_tax"
setArrElementN proto20,CreateArray2("m_columns",empty),"amount_received"
setArrElementN proto20,CreateArray2("m_columns",empty),"incomplete"
setArrElementN proto20,CreateArray2("m_columns",empty),"credit_account_seq"
setArrElementN proto20,CreateArray2("m_columns",empty),"bank_code_mand"
setArrElementN proto20,CreateArray2("m_columns",empty),"grid_row"
setArrElementN proto20,CreateArray2("m_columns",empty),"balance_remain"
setArrElementN proto20,CreateArray2("m_columns",empty),"balance_remain_pay"
setArrElementN proto20,CreateArray2("m_columns",empty),"token_id"
setArrElementN proto20,CreateArray2("m_columns",empty),"txn_id"
setArrElementN proto20,CreateArray2("m_columns",empty),"refund_flag"
setArrElementN proto20,CreateArray2("m_columns",empty),"hash"
Set obj = (CreateClass("SQLTable",1,proto20,Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto19,"m_table",obj
setArrElement proto19,"m_alias",""
Set proto21 = (CreateDictionary())
setArrElement proto21,"m_sql",""
setArrElement proto21,"m_uniontype","SQLL_UNKNOWN"
Set obj = (CreateClass("SQLNonParsed",1,CreateDictionary1("m_sql",""),Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto21,"m_column",obj
setArrElement proto21,"m_contained",CreateDictionary()
setArrElement proto21,"m_strCase",""
setArrElement proto21,"m_havingmode","0"
setArrElement proto21,"m_inBrackets","0"
setArrElement proto21,"m_useAlias","0"
Set obj = (CreateClass("SQLLogicalExpr",1,proto21,Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto19,"m_joinon",obj
Set obj = (CreateClass("SQLFromListItem",1,proto19,Empty,Empty,Empty,Empty,Empty,Empty))

setArrElementN proto0,CreateArray2("m_fromlist",empty),obj
Set proto23 = (CreateDictionary())
setArrElement proto23,"m_link","SQLL_INNERJOIN"
Set proto24 = (CreateDictionary())
setArrElement proto24,"m_strName","part"
setArrElement proto24,"m_columns",CreateDictionary()
setArrElementN proto24,CreateArray2("m_columns",empty),"code"
setArrElementN proto24,CreateArray2("m_columns",empty),"type_list"
setArrElementN proto24,CreateArray2("m_columns",empty),"part_id"
setArrElementN proto24,CreateArray2("m_columns",empty),"name"
setArrElementN proto24,CreateArray2("m_columns",empty),"short_name"
setArrElementN proto24,CreateArray2("m_columns",empty),"branch"
setArrElementN proto24,CreateArray2("m_columns",empty),"sales_region"
setArrElementN proto24,CreateArray2("m_columns",empty),"tel1"
setArrElementN proto24,CreateArray2("m_columns",empty),"fax"
setArrElementN proto24,CreateArray2("m_columns",empty),"address1"
setArrElementN proto24,CreateArray2("m_columns",empty),"address2"
setArrElementN proto24,CreateArray2("m_columns",empty),"state"
setArrElementN proto24,CreateArray2("m_columns",empty),"zip"
setArrElementN proto24,CreateArray2("m_columns",empty),"city"
setArrElementN proto24,CreateArray2("m_columns",empty),"country"
setArrElementN proto24,CreateArray2("m_columns",empty),"account"
setArrElementN proto24,CreateArray2("m_columns",empty),"vat_number"
setArrElementN proto24,CreateArray2("m_columns",empty),"customs_number"
setArrElementN proto24,CreateArray2("m_columns",empty),"iata_comm_perc"
setArrElementN proto24,CreateArray2("m_columns",empty),"is_cass"
setArrElementN proto24,CreateArray2("m_columns",empty),"is_import"
setArrElementN proto24,CreateArray2("m_columns",empty),"cass_import"
setArrElementN proto24,CreateArray2("m_columns",empty),"deal_optimize"
setArrElementN proto24,CreateArray2("m_columns",empty),"currency"
setArrElementN proto24,CreateArray2("m_columns",empty),"my_agent"
setArrElementN proto24,CreateArray2("m_columns",empty),"awb_seq"
setArrElementN proto24,CreateArray2("m_columns",empty),"remarks"
setArrElementN proto24,CreateArray2("m_columns",empty),"status"
setArrElementN proto24,CreateArray2("m_columns",empty),"address_cont"
setArrElementN proto24,CreateArray2("m_columns",empty),"invoice_due_days"
setArrElementN proto24,CreateArray2("m_columns",empty),"no_dunning"
setArrElementN proto24,CreateArray2("m_columns",empty),"exclude_invoice"
setArrElementN proto24,CreateArray2("m_columns",empty),"sales_office"
setArrElementN proto24,CreateArray2("m_columns",empty),"sales_area"
setArrElementN proto24,CreateArray2("m_columns",empty),"sales_oper"
setArrElementN proto24,CreateArray2("m_columns",empty),"name2"
setArrElementN proto24,CreateArray2("m_columns",empty),"department"
setArrElementN proto24,CreateArray2("m_columns",empty),"security"
setArrElementN proto24,CreateArray2("m_columns",empty),"billing_agent"
setArrElementN proto24,CreateArray2("m_columns",empty),"no_cc_fee"
setArrElementN proto24,CreateArray2("m_columns",empty),"run_flag"
setArrElementN proto24,CreateArray2("m_columns",empty),"last_use_date"
setArrElementN proto24,CreateArray2("m_columns",empty),"language"
setArrElementN proto24,CreateArray2("m_columns",empty),"no_vat"
setArrElementN proto24,CreateArray2("m_columns",empty),"flags"
setArrElementN proto24,CreateArray2("m_columns",empty),"all_branch_code"
setArrElementN proto24,CreateArray2("m_columns",empty),"tax_number"
setArrElementN proto24,CreateArray2("m_columns",empty),"oper"
setArrElementN proto24,CreateArray2("m_columns",empty),"time_stamp"
setArrElementN proto24,CreateArray2("m_columns",empty),"cha_discount"
setArrElementN proto24,CreateArray2("m_columns",empty),"consignor"
setArrElementN proto24,CreateArray2("m_columns",empty),"valuation_perc"
setArrElementN proto24,CreateArray2("m_columns",empty),"client_id"
setArrElementN proto24,CreateArray2("m_columns",empty),"firms_code"
setArrElementN proto24,CreateArray2("m_columns",empty),"color"
setArrElementN proto24,CreateArray2("m_columns",empty),"product_code"
setArrElementN proto24,CreateArray2("m_columns",empty),"invoice_limit"
setArrElementN proto24,CreateArray2("m_columns",empty),"stock_holder"
setArrElementN proto24,CreateArray2("m_columns",empty),"no_arr_print"
setArrElementN proto24,CreateArray2("m_columns",empty),"iac_code"
setArrElementN proto24,CreateArray2("m_columns",empty),"iac_until_date"
setArrElementN proto24,CreateArray2("m_columns",empty),"handle_agent"
setArrElementN proto24,CreateArray2("m_columns",empty),"account2"
setArrElementN proto24,CreateArray2("m_columns",empty),"po_mand"
setArrElementN proto24,CreateArray2("m_columns",empty),"pima_address"
setArrElementN proto24,CreateArray2("m_columns",empty),"ccs_address"
setArrElementN proto24,CreateArray2("m_columns",empty),"dsip"
setArrElementN proto24,CreateArray2("m_columns",empty),"tsa"
setArrElementN proto24,CreateArray2("m_columns",empty),"account_oper"
setArrElementN proto24,CreateArray2("m_columns",empty),"billing_cycle"
setArrElementN proto24,CreateArray2("m_columns",empty),"is_customs_ware"
setArrElementN proto24,CreateArray2("m_columns",empty),"invoice_layout"
setArrElementN proto24,CreateArray2("m_columns",empty),"is_pay_slip"
setArrElementN proto24,CreateArray2("m_columns",empty),"account3"
setArrElementN proto24,CreateArray2("m_columns",empty),"is_agent"
setArrElementN proto24,CreateArray2("m_columns",empty),"is_shipper"
setArrElementN proto24,CreateArray2("m_columns",empty),"is_consignee"
setArrElementN proto24,CreateArray2("m_columns",empty),"is_temp"
setArrElementN proto24,CreateArray2("m_columns",empty),"my_agent_until_date"
setArrElementN proto24,CreateArray2("m_columns",empty),"agent_group"
setArrElementN proto24,CreateArray2("m_columns",empty),"invoice_per_item"
setArrElementN proto24,CreateArray2("m_columns",empty),"region_location_code"
setArrElementN proto24,CreateArray2("m_columns",empty),"cost_center"
setArrElementN proto24,CreateArray2("m_columns",empty),"billing_currency"
setArrElementN proto24,CreateArray2("m_columns",empty),"part_restrict_flags"
setArrElementN proto24,CreateArray2("m_columns",empty),"net_comm_perc"
setArrElementN proto24,CreateArray2("m_columns",empty),"cass_branch"
setArrElementN proto24,CreateArray2("m_columns",empty),"iata_agent_code"
setArrElementN proto24,CreateArray2("m_columns",empty),"agent_type"
setArrElementN proto24,CreateArray2("m_columns",empty),"external_id"
setArrElementN proto24,CreateArray2("m_columns",empty),"parent_agent"
setArrElementN proto24,CreateArray2("m_columns",empty),"efreight_group"
setArrElementN proto24,CreateArray2("m_columns",empty),"shipper_invoice"
setArrElementN proto24,CreateArray2("m_columns",empty),"mail_part_id"
setArrElementN proto24,CreateArray2("m_columns",empty),"mail_oe_group"
setArrElementN proto24,CreateArray2("m_columns",empty),"mail_print_inv"
setArrElementN proto24,CreateArray2("m_columns",empty),"mail_email_pdf"
setArrElementN proto24,CreateArray2("m_columns",empty),"mail_email_xls"
setArrElementN proto24,CreateArray2("m_columns",empty),"sap_number"
setArrElementN proto24,CreateArray2("m_columns",empty),"gsa_code"
setArrElementN proto24,CreateArray2("m_columns",empty),"tax_group"
setArrElementN proto24,CreateArray2("m_columns",empty),"creation_date"
setArrElementN proto24,CreateArray2("m_columns",empty),"on_invoice"
setArrElementN proto24,CreateArray2("m_columns",empty),"formatted_message"
setArrElementN proto24,CreateArray2("m_columns",empty),"rakc_id"
setArrElementN proto24,CreateArray2("m_columns",empty),"son_flag"
Set obj = (CreateClass("SQLTable",1,proto24,Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto23,"m_table",obj
setArrElement proto23,"m_alias",""
Set proto25 = (CreateDictionary())
setArrElement proto25,"m_sql","payment.part = part.code"
setArrElement proto25,"m_uniontype","SQLL_UNKNOWN"
Set obj = (CreateClass("SQLField",1,CreateDictionary2("m_strName","part","m_strTable","payment"),Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto25,"m_column",obj
setArrElement proto25,"m_contained",CreateDictionary()
setArrElement proto25,"m_strCase","= part.code"
setArrElement proto25,"m_havingmode","0"
setArrElement proto25,"m_inBrackets","0"
setArrElement proto25,"m_useAlias","0"
Set obj = (CreateClass("SQLLogicalExpr",1,proto25,Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto23,"m_joinon",obj
Set obj = (CreateClass("SQLFromListItem",1,proto23,Empty,Empty,Empty,Empty,Empty,Empty))

setArrElementN proto0,CreateArray2("m_fromlist",empty),obj
Set proto27 = (CreateDictionary())
setArrElement proto27,"m_link","SQLL_INNERJOIN"
Set proto28 = (CreateDictionary())
setArrElement proto28,"m_strName","awb"
setArrElement proto28,"m_columns",CreateDictionary()
setArrElementN proto28,CreateArray2("m_columns",empty),"seq"
setArrElementN proto28,CreateArray2("m_columns",empty),"prefix"
setArrElementN proto28,CreateArray2("m_columns",empty),"serial"
setArrElementN proto28,CreateArray2("m_columns",empty),"awb_exp_seq"
setArrElementN proto28,CreateArray2("m_columns",empty),"issue_date"
setArrElementN proto28,CreateArray2("m_columns",empty),"issue_place"
setArrElementN proto28,CreateArray2("m_columns",empty),"issue_carrier"
setArrElementN proto28,CreateArray2("m_columns",empty),"agent"
setArrElementN proto28,CreateArray2("m_columns",empty),"service"
setArrElementN proto28,CreateArray2("m_columns",empty),"origin"
setArrElementN proto28,CreateArray2("m_columns",empty),"dest"
setArrElementN proto28,CreateArray2("m_columns",empty),"value_carriage"
setArrElementN proto28,CreateArray2("m_columns",empty),"value_customs"
setArrElementN proto28,CreateArray2("m_columns",empty),"insurance"
setArrElementN proto28,CreateArray2("m_columns",empty),"tax"
setArrElementN proto28,CreateArray2("m_columns",empty),"volume_code"
setArrElementN proto28,CreateArray2("m_columns",empty),"volume"
setArrElementN proto28,CreateArray2("m_columns",empty),"density"
setArrElementN proto28,CreateArray2("m_columns",empty),"kg_lb"
setArrElementN proto28,CreateArray2("m_columns",empty),"pieces"
setArrElementN proto28,CreateArray2("m_columns",empty),"weight"
setArrElementN proto28,CreateArray2("m_columns",empty),"dim_msr_code"
setArrElementN proto28,CreateArray2("m_columns",empty),"master_seq"
setArrElementN proto28,CreateArray2("m_columns",empty),"profit"
setArrElementN proto28,CreateArray2("m_columns",empty),"customs_status"
setArrElementN proto28,CreateArray2("m_columns",empty),"arrival_date"
setArrElementN proto28,CreateArray2("m_columns",empty),"nature_goods"
setArrElementN proto28,CreateArray2("m_columns",empty),"shipper_reference"
setArrElementN proto28,CreateArray2("m_columns",empty),"book_reference"
setArrElementN proto28,CreateArray2("m_columns",empty),"notifier"
setArrElementN proto28,CreateArray2("m_columns",empty),"import_agent"
setArrElementN proto28,CreateArray2("m_columns",empty),"import_service"
setArrElementN proto28,CreateArray2("m_columns",empty),"awb_imp_seq"
setArrElementN proto28,CreateArray2("m_columns",empty),"time_stamp"
setArrElementN proto28,CreateArray2("m_columns",empty),"priority"
setArrElementN proto28,CreateArray2("m_columns",empty),"price_class"
setArrElementN proto28,CreateArray2("m_columns",empty),"book_status"
setArrElementN proto28,CreateArray2("m_columns",empty),"service_level"
setArrElementN proto28,CreateArray2("m_columns",empty),"sales_region"
setArrElementN proto28,CreateArray2("m_columns",empty),"oper"
setArrElementN proto28,CreateArray2("m_columns",empty),"remarks"
setArrElementN proto28,CreateArray2("m_columns",empty),"son_flag"
setArrElementN proto28,CreateArray2("m_columns",empty),"house_flag"
setArrElementN proto28,CreateArray2("m_columns",empty),"cha_code"
setArrElementN proto28,CreateArray2("m_columns",empty),"osi_remarks"
setArrElementN proto28,CreateArray2("m_columns",empty),"message_status"
setArrElementN proto28,CreateArray2("m_columns",empty),"pua_code"
setArrElementN proto28,CreateArray2("m_columns",empty),"book_price_seq"
setArrElementN proto28,CreateArray2("m_columns",empty),"customs_reference"
setArrElementN proto28,CreateArray2("m_columns",empty),"shc_list"
setArrElementN proto28,CreateArray2("m_columns",empty),"net_prorate"
setArrElementN proto28,CreateArray2("m_columns",empty),"prorate_date"
setArrElementN proto28,CreateArray2("m_columns",empty),"client_id"
setArrElementN proto28,CreateArray2("m_columns",empty),"security_flag"
setArrElementN proto28,CreateArray2("m_columns",empty),"incl_stats"
setArrElementN proto28,CreateArray2("m_columns",empty),"slc_pieces"
setArrElementN proto28,CreateArray2("m_columns",empty),"un_number"
setArrElementN proto28,CreateArray2("m_columns",empty),"book_weight"
setArrElementN proto28,CreateArray2("m_columns",empty),"book_volume"
setArrElementN proto28,CreateArray2("m_columns",empty),"export_station"
setArrElementN proto28,CreateArray2("m_columns",empty),"book_rate"
setArrElementN proto28,CreateArray2("m_columns",empty),"all_prorate"
setArrElementN proto28,CreateArray2("m_columns",empty),"departure_date"
setArrElementN proto28,CreateArray2("m_columns",empty),"x_ray"
setArrElementN proto28,CreateArray2("m_columns",empty),"rcv_weight"
setArrElementN proto28,CreateArray2("m_columns",empty),"book_station"
setArrElementN proto28,CreateArray2("m_columns",empty),"book_office"
setArrElementN proto28,CreateArray2("m_columns",empty),"dgr_check"
setArrElementN proto28,CreateArray2("m_columns",empty),"csr_flt_line"
setArrElementN proto28,CreateArray2("m_columns",empty),"screen_oper"
setArrElementN proto28,CreateArray2("m_columns",empty),"customs_code"
setArrElementN proto28,CreateArray2("m_columns",empty),"screen_pieces"
setArrElementN proto28,CreateArray2("m_columns",empty),"stage_flag"
setArrElementN proto28,CreateArray2("m_columns",empty),"gsa_comm_perc"
setArrElementN proto28,CreateArray2("m_columns",empty),"son_flag2"
setArrElementN proto28,CreateArray2("m_columns",empty),"direct_weight"
setArrElementN proto28,CreateArray2("m_columns",empty),"security_profile_seq"
setArrElementN proto28,CreateArray2("m_columns",empty),"part_secure_seq"
setArrElementN proto28,CreateArray2("m_columns",empty),"rule_code"
setArrElementN proto28,CreateArray2("m_columns",empty),"balance_flag"
setArrElementN proto28,CreateArray2("m_columns",empty),"po_ref"
setArrElementN proto28,CreateArray2("m_columns",empty),"gcg_status_code"
setArrElementN proto28,CreateArray2("m_columns",empty),"issue_date_log"
setArrElementN proto28,CreateArray2("m_columns",empty),"issue_timestamp"
setArrElementN proto28,CreateArray2("m_columns",empty),"serial_prefix"
setArrElementN proto28,CreateArray2("m_columns",empty),"compatibility_check"
setArrElementN proto28,CreateArray2("m_columns",empty),"issue_place_text"
setArrElementN proto28,CreateArray2("m_columns",empty),"signature"
setArrElementN proto28,CreateArray2("m_columns",empty),"dim_density"
setArrElementN proto28,CreateArray2("m_columns",empty),"rating_weight"
setArrElementN proto28,CreateArray2("m_columns",empty),"biling_type"
setArrElementN proto28,CreateArray2("m_columns",empty),"biling_account"
setArrElementN proto28,CreateArray2("m_columns",empty),"cit_reason"
setArrElementN proto28,CreateArray2("m_columns",empty),"iata_note"
Set obj = (CreateClass("SQLTable",1,proto28,Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto27,"m_table",obj
setArrElement proto27,"m_alias",""
Set proto29 = (CreateDictionary())
setArrElement proto29,"m_sql","payment.awb_seq = awb.seq"
setArrElement proto29,"m_uniontype","SQLL_UNKNOWN"
Set obj = (CreateClass("SQLField",1,CreateDictionary2("m_strName","awb_seq","m_strTable","payment"),Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto29,"m_column",obj
setArrElement proto29,"m_contained",CreateDictionary()
setArrElement proto29,"m_strCase","= awb.seq"
setArrElement proto29,"m_havingmode","0"
setArrElement proto29,"m_inBrackets","0"
setArrElement proto29,"m_useAlias","0"
Set obj = (CreateClass("SQLLogicalExpr",1,proto29,Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto27,"m_joinon",obj
Set obj = (CreateClass("SQLFromListItem",1,proto27,Empty,Empty,Empty,Empty,Empty,Empty))

setArrElementN proto0,CreateArray2("m_fromlist",empty),obj
setArrElement proto0,"m_groupby",CreateDictionary()
setArrElement proto0,"m_orderby",CreateDictionary()
Set proto31 = (CreateDictionary())
Set obj = (CreateClass("SQLField",1,CreateDictionary2("m_strName","payment_date","m_strTable","payment"),Empty,Empty,Empty,Empty,Empty,Empty))

setArrElement proto31,"m_column",obj
setArrElement proto31,"m_bAsc",1
setArrElement proto31,"m_nColumn",0
Set obj = (CreateClass("SQLOrderByItem",1,proto31,Empty,Empty,Empty,Empty,Empty,Empty))

setArrElementN proto0,CreateArray2("m_orderby",empty),obj
Set obj = (CreateClass("SQLQuery",1,proto0,Empty,Empty,Empty,Empty,Empty,Empty))

doAssignment queryData_payment,obj

setArrElement tdatapayment,".sqlquery",queryData_payment
setArrElement tableEvents,"payment",CreateClass("eventsBase",0,Empty,Empty,Empty,Empty,Empty,Empty,Empty)
%>
