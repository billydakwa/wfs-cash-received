<%
%>
<!--#include file="include/dbcommon.asp"-->
<%
add_nocache_headers 
asp_include "include/payment_variables.asp",false
asp_include "classes/searchcontrol.asp",false
asp_include "classes/advancedsearchcontrol.asp",false
asp_include "classes/panelsearchcontrol.asp",false
asp_include "classes/searchclause.asp",false
doAssignment sessionPrefix,strTableName
includes = ""
predefFieldNum = 0
asp_include "include/xtempl.asp",false
asp_include "classes/runnerpage.asp",false
Set xt = (CreateClass("Xtempl",0,Empty,Empty,Empty,Empty,Empty,Empty,Empty))
if bValue(postvalue("id")) then
	doAssignmentByRef id,postvalue("id")
else
	id = 1
end if
mode = SEARCH_SIMPLE
templatefile = "payment_search.htm"
if IsEqual(postvalue("mode"),"inlineLoadCtrl") then
	mode = SEARCH_LOAD_CONTROL
	templatefile = "payment_inline_search.htm"
end if
calendar = false
calendar = true
timepicker = false
Set params = (CreateDictionary())
setArrElement params,"id",id
setArrElement params,"mode",mode
setArrElement params,"calendar",calendar
setArrElement params,"timepicker",timepicker
setArrElementByRef params,"xt",xt
setArrElement params,"shortTableName","payment"
setArrElement params,"origTName",strOriginalTableName
setArrElement params,"sessionPrefix",sessionPrefix
setArrElement params,"tName",strTableName
setArrElement params,"includes_js",includes_js
setArrElement params,"includes_jsreq",includes_jsreq
setArrElement params,"includes_css",includes_css
setArrElement params,"locale_info",locale_info
setArrElement params,"pageType",PAGE_SEARCH
Set pageObject = (CreateClass("RunnerPage",1,params,Empty,Empty,Empty,Empty,Empty,Empty))
doAssignment searchControllerId,IIF(postvalue("searchControllerId"),postvalue("searchControllerId"),pageObject.id)
if bValue(eventObj.exists_p1("BeforeProcessSearch")) then
	eventObj.BeforeProcessSearch_p1 conn
end if
if IsEqual(mode,SEARCH_SIMPLE) then
	Set searchControlBuilder = (CreateClass("AdvancedSearchControl",4,searchControllerId,strTableName,pageObject.searchClauseObj,pageObject,Empty,Empty,Empty))
	doAssignmentByRef onLoadJsCode,GetTableData(pageObject.tName,".jsOnloadSearch","")
	pageObject.addOnLoadJsEvent_p1 onLoadJsCode
	pageObject.addButtonHandlers 
	includes = CSmartStr(includes) & "<script language=""JavaScript"" src=""include/jquery.js""></script>" & vbcrlf
	includes = CSmartStr(includes) & "<script language=""JavaScript"" src=""include/customlabels.js""></script>" & vbcrlf
	includes = CSmartStr(includes) & "<script language=""JavaScript"" src=""include/jsfunctions.js""></script>" & vbcrlf
	if IsIdentical(pageObject.debugJSMode,true) then
		includes = CSmartStr(includes) & "<script language=""JavaScript"" src=""include/runnerJS/RunnerBase.js""></script>" & vbcrlf
	else
		includes = CSmartStr(includes) & "<script language=""JavaScript"" src=""include/runnerJS/RunnerBase.js""></script>" & vbcrlf
	end if
	includes = CSmartStr(includes) & "<div id=""search_suggest"" class=""search_suggest""></div>"
	doAssignmentByRef searchRadio,searchControlBuilder.getSearchRadio()
	xt.assign_section_p3 "all_checkbox_label",ArrayElement(ArrayElement(searchRadio,"all_checkbox_label"),0),ArrayElement(ArrayElement(searchRadio,"all_checkbox_label"),1)
	xt.assign_section_p3 "any_checkbox_label",ArrayElement(ArrayElement(searchRadio,"any_checkbox_label"),0),ArrayElement(ArrayElement(searchRadio,"any_checkbox_label"),1)
	xt.assignbyref_p2 "all_checkbox",ArrayElement(searchRadio,"all_checkbox")
	xt.assignbyref_p2 "any_checkbox",ArrayElement(searchRadio,"any_checkbox")
	if bValue(GetLookupTable("name",strTableName)) then
		setArrElementN pageObject.settingsMap,CreateArray3("globalSettings","shortTNames",GetLookupTable("name",strTableName)),GetTableURL(GetLookupTable("name",strTableName))
	end if
	pageObject.fillFieldToolTips_p1 "name"
	doAssignmentByRef srchFields,pageObject.searchClauseObj.getSearchCtrlParams_p1("name")
	Set firstFieldParams = (CreateDictionary())
	if bValue(asp_count(srchFields)) then
		doAssignment firstFieldParams,ArrayElement(srchFields,0)
	else
		setArrElement firstFieldParams,"fName","name"
		setArrElement firstFieldParams,"eType",""
		setArrElement firstFieldParams,"value1",""
		setArrElement firstFieldParams,"opt",""
		setArrElement firstFieldParams,"value2",""
		setArrElement firstFieldParams,"not",false
	end if
	doAssignmentByRef ctrlBlockArr,searchControlBuilder.buildSearchCtrlBlockArr_p8(id,ArrayElement(firstFieldParams,"fName"),0,ArrayElement(firstFieldParams,"opt"),ArrayElement(firstFieldParams,"not"),false,ArrayElement(firstFieldParams,"value1"),ArrayElement(firstFieldParams,"value2"))
	if bValue(isEnableSection508()) then
		xt.assign_section_p3 "name_label",("<label for=""" & CSmartStr(GetInputElementId("name",id))) & """>","</label>"
	else
		xt.assign_p2 "name_label",true
	end if
	xt.assign_p2 "name_fieldblock",true
	xt.assignbyref_p2 "name_editcontrol",ArrayElement(ctrlBlockArr,"searchcontrol")
	xt.assign_p2 "name_notbox",ArrayElement(ctrlBlockArr,"notbox")
	xt.assignbyref_p2 "name_editcontrol1",ArrayElement(ctrlBlockArr,"searchcontrol1")
	xt.assign_p2 "searchtype_name",ArrayElement(ctrlBlockArr,"searchtype")
	doAssignmentByRef isFieldNeedSecCtrl,searchControlBuilder.isNeedSecondCtrl_p1("name")
	ctrlInd = 0
	if bValue(isFieldNeedSecCtrl) then
		setArrElementN pageObject.controlsMap,CreateArray3("search","searchBlocks",empty),CreateDictionary3("fName","name","recId",id,"ctrlsMap",CreateDictionary2(0,ctrlInd,1,CSmartDbl(ctrlInd)+1))
		ctrlInd = CSmartDbl(ctrlInd)+2
	else
		setArrElementN pageObject.controlsMap,CreateArray3("search","searchBlocks",empty),CreateDictionary3("fName","name","recId",id,"ctrlsMap",CreateDictionary1(0,ctrlInd))
		ctrlInd = CSmartDbl(ctrlInd)+1
	end if
	if bValue(GetLookupTable("serial",strTableName)) then
		setArrElementN pageObject.settingsMap,CreateArray3("globalSettings","shortTNames",GetLookupTable("serial",strTableName)),GetTableURL(GetLookupTable("serial",strTableName))
	end if
	pageObject.fillFieldToolTips_p1 "serial"
	doAssignmentByRef srchFields,pageObject.searchClauseObj.getSearchCtrlParams_p1("serial")
	Set firstFieldParams = (CreateDictionary())
	if bValue(asp_count(srchFields)) then
		doAssignment firstFieldParams,ArrayElement(srchFields,0)
	else
		setArrElement firstFieldParams,"fName","serial"
		setArrElement firstFieldParams,"eType",""
		setArrElement firstFieldParams,"value1",""
		setArrElement firstFieldParams,"opt",""
		setArrElement firstFieldParams,"value2",""
		setArrElement firstFieldParams,"not",false
	end if
	doAssignmentByRef ctrlBlockArr,searchControlBuilder.buildSearchCtrlBlockArr_p8(id,ArrayElement(firstFieldParams,"fName"),0,ArrayElement(firstFieldParams,"opt"),ArrayElement(firstFieldParams,"not"),false,ArrayElement(firstFieldParams,"value1"),ArrayElement(firstFieldParams,"value2"))
	if bValue(isEnableSection508()) then
		xt.assign_section_p3 "serial_label",("<label for=""" & CSmartStr(GetInputElementId("serial",id))) & """>","</label>"
	else
		xt.assign_p2 "serial_label",true
	end if
	xt.assign_p2 "serial_fieldblock",true
	xt.assignbyref_p2 "serial_editcontrol",ArrayElement(ctrlBlockArr,"searchcontrol")
	xt.assign_p2 "serial_notbox",ArrayElement(ctrlBlockArr,"notbox")
	xt.assignbyref_p2 "serial_editcontrol1",ArrayElement(ctrlBlockArr,"searchcontrol1")
	xt.assign_p2 "searchtype_serial",ArrayElement(ctrlBlockArr,"searchtype")
	doAssignmentByRef isFieldNeedSecCtrl,searchControlBuilder.isNeedSecondCtrl_p1("serial")
	ctrlInd = 0
	if bValue(isFieldNeedSecCtrl) then
		setArrElementN pageObject.controlsMap,CreateArray3("search","searchBlocks",empty),CreateDictionary3("fName","serial","recId",id,"ctrlsMap",CreateDictionary2(0,ctrlInd,1,CSmartDbl(ctrlInd)+1))
		ctrlInd = CSmartDbl(ctrlInd)+2
	else
		setArrElementN pageObject.controlsMap,CreateArray3("search","searchBlocks",empty),CreateDictionary3("fName","serial","recId",id,"ctrlsMap",CreateDictionary1(0,ctrlInd))
		ctrlInd = CSmartDbl(ctrlInd)+1
	end if
	if bValue(GetLookupTable("amount_received",strTableName)) then
		setArrElementN pageObject.settingsMap,CreateArray3("globalSettings","shortTNames",GetLookupTable("amount_received",strTableName)),GetTableURL(GetLookupTable("amount_received",strTableName))
	end if
	pageObject.fillFieldToolTips_p1 "amount_received"
	doAssignmentByRef srchFields,pageObject.searchClauseObj.getSearchCtrlParams_p1("amount_received")
	Set firstFieldParams = (CreateDictionary())
	if bValue(asp_count(srchFields)) then
		doAssignment firstFieldParams,ArrayElement(srchFields,0)
	else
		setArrElement firstFieldParams,"fName","amount_received"
		setArrElement firstFieldParams,"eType",""
		setArrElement firstFieldParams,"value1",""
		setArrElement firstFieldParams,"opt",""
		setArrElement firstFieldParams,"value2",""
		setArrElement firstFieldParams,"not",false
	end if
	doAssignmentByRef ctrlBlockArr,searchControlBuilder.buildSearchCtrlBlockArr_p8(id,ArrayElement(firstFieldParams,"fName"),0,ArrayElement(firstFieldParams,"opt"),ArrayElement(firstFieldParams,"not"),false,ArrayElement(firstFieldParams,"value1"),ArrayElement(firstFieldParams,"value2"))
	if bValue(isEnableSection508()) then
		xt.assign_section_p3 "amount_received_label",("<label for=""" & CSmartStr(GetInputElementId("amount_received",id))) & """>","</label>"
	else
		xt.assign_p2 "amount_received_label",true
	end if
	xt.assign_p2 "amount_received_fieldblock",true
	xt.assignbyref_p2 "amount_received_editcontrol",ArrayElement(ctrlBlockArr,"searchcontrol")
	xt.assign_p2 "amount_received_notbox",ArrayElement(ctrlBlockArr,"notbox")
	xt.assignbyref_p2 "amount_received_editcontrol1",ArrayElement(ctrlBlockArr,"searchcontrol1")
	xt.assign_p2 "searchtype_amount_received",ArrayElement(ctrlBlockArr,"searchtype")
	doAssignmentByRef isFieldNeedSecCtrl,searchControlBuilder.isNeedSecondCtrl_p1("amount_received")
	ctrlInd = 0
	if bValue(isFieldNeedSecCtrl) then
		setArrElementN pageObject.controlsMap,CreateArray3("search","searchBlocks",empty),CreateDictionary3("fName","amount_received","recId",id,"ctrlsMap",CreateDictionary2(0,ctrlInd,1,CSmartDbl(ctrlInd)+1))
		ctrlInd = CSmartDbl(ctrlInd)+2
	else
		setArrElementN pageObject.controlsMap,CreateArray3("search","searchBlocks",empty),CreateDictionary3("fName","amount_received","recId",id,"ctrlsMap",CreateDictionary1(0,ctrlInd))
		ctrlInd = CSmartDbl(ctrlInd)+1
	end if
	if bValue(GetLookupTable("prefix",strTableName)) then
		setArrElementN pageObject.settingsMap,CreateArray3("globalSettings","shortTNames",GetLookupTable("prefix",strTableName)),GetTableURL(GetLookupTable("prefix",strTableName))
	end if
	pageObject.fillFieldToolTips_p1 "prefix"
	doAssignmentByRef srchFields,pageObject.searchClauseObj.getSearchCtrlParams_p1("prefix")
	Set firstFieldParams = (CreateDictionary())
	if bValue(asp_count(srchFields)) then
		doAssignment firstFieldParams,ArrayElement(srchFields,0)
	else
		setArrElement firstFieldParams,"fName","prefix"
		setArrElement firstFieldParams,"eType",""
		setArrElement firstFieldParams,"value1",""
		setArrElement firstFieldParams,"opt",""
		setArrElement firstFieldParams,"value2",""
		setArrElement firstFieldParams,"not",false
	end if
	doAssignmentByRef ctrlBlockArr,searchControlBuilder.buildSearchCtrlBlockArr_p8(id,ArrayElement(firstFieldParams,"fName"),0,ArrayElement(firstFieldParams,"opt"),ArrayElement(firstFieldParams,"not"),false,ArrayElement(firstFieldParams,"value1"),ArrayElement(firstFieldParams,"value2"))
	if bValue(isEnableSection508()) then
		xt.assign_section_p3 "prefix_label",("<label for=""" & CSmartStr(GetInputElementId("prefix",id))) & """>","</label>"
	else
		xt.assign_p2 "prefix_label",true
	end if
	xt.assign_p2 "prefix_fieldblock",true
	xt.assignbyref_p2 "prefix_editcontrol",ArrayElement(ctrlBlockArr,"searchcontrol")
	xt.assign_p2 "prefix_notbox",ArrayElement(ctrlBlockArr,"notbox")
	xt.assignbyref_p2 "prefix_editcontrol1",ArrayElement(ctrlBlockArr,"searchcontrol1")
	xt.assign_p2 "searchtype_prefix",ArrayElement(ctrlBlockArr,"searchtype")
	doAssignmentByRef isFieldNeedSecCtrl,searchControlBuilder.isNeedSecondCtrl_p1("prefix")
	ctrlInd = 0
	if bValue(isFieldNeedSecCtrl) then
		setArrElementN pageObject.controlsMap,CreateArray3("search","searchBlocks",empty),CreateDictionary3("fName","prefix","recId",id,"ctrlsMap",CreateDictionary2(0,ctrlInd,1,CSmartDbl(ctrlInd)+1))
		ctrlInd = CSmartDbl(ctrlInd)+2
	else
		setArrElementN pageObject.controlsMap,CreateArray3("search","searchBlocks",empty),CreateDictionary3("fName","prefix","recId",id,"ctrlsMap",CreateDictionary1(0,ctrlInd))
		ctrlInd = CSmartDbl(ctrlInd)+1
	end if
	if bValue(GetLookupTable("station",strTableName)) then
		setArrElementN pageObject.settingsMap,CreateArray3("globalSettings","shortTNames",GetLookupTable("station",strTableName)),GetTableURL(GetLookupTable("station",strTableName))
	end if
	pageObject.fillFieldToolTips_p1 "station"
	doAssignmentByRef srchFields,pageObject.searchClauseObj.getSearchCtrlParams_p1("station")
	Set firstFieldParams = (CreateDictionary())
	if bValue(asp_count(srchFields)) then
		doAssignment firstFieldParams,ArrayElement(srchFields,0)
	else
		setArrElement firstFieldParams,"fName","station"
		setArrElement firstFieldParams,"eType",""
		setArrElement firstFieldParams,"value1",""
		setArrElement firstFieldParams,"opt",""
		setArrElement firstFieldParams,"value2",""
		setArrElement firstFieldParams,"not",false
	end if
	doAssignmentByRef ctrlBlockArr,searchControlBuilder.buildSearchCtrlBlockArr_p8(id,ArrayElement(firstFieldParams,"fName"),0,ArrayElement(firstFieldParams,"opt"),ArrayElement(firstFieldParams,"not"),false,ArrayElement(firstFieldParams,"value1"),ArrayElement(firstFieldParams,"value2"))
	if bValue(isEnableSection508()) then
		xt.assign_section_p3 "station_label",("<label for=""" & CSmartStr(GetInputElementId("station",id))) & """>","</label>"
	else
		xt.assign_p2 "station_label",true
	end if
	xt.assign_p2 "station_fieldblock",true
	xt.assignbyref_p2 "station_editcontrol",ArrayElement(ctrlBlockArr,"searchcontrol")
	xt.assign_p2 "station_notbox",ArrayElement(ctrlBlockArr,"notbox")
	xt.assignbyref_p2 "station_editcontrol1",ArrayElement(ctrlBlockArr,"searchcontrol1")
	xt.assign_p2 "searchtype_station",ArrayElement(ctrlBlockArr,"searchtype")
	doAssignmentByRef isFieldNeedSecCtrl,searchControlBuilder.isNeedSecondCtrl_p1("station")
	ctrlInd = 0
	if bValue(isFieldNeedSecCtrl) then
		setArrElementN pageObject.controlsMap,CreateArray3("search","searchBlocks",empty),CreateDictionary3("fName","station","recId",id,"ctrlsMap",CreateDictionary2(0,ctrlInd,1,CSmartDbl(ctrlInd)+1))
		ctrlInd = CSmartDbl(ctrlInd)+2
	else
		setArrElementN pageObject.controlsMap,CreateArray3("search","searchBlocks",empty),CreateDictionary3("fName","station","recId",id,"ctrlsMap",CreateDictionary1(0,ctrlInd))
		ctrlInd = CSmartDbl(ctrlInd)+1
	end if
	if bValue(GetLookupTable("payment_date",strTableName)) then
		setArrElementN pageObject.settingsMap,CreateArray3("globalSettings","shortTNames",GetLookupTable("payment_date",strTableName)),GetTableURL(GetLookupTable("payment_date",strTableName))
	end if
	pageObject.fillFieldToolTips_p1 "payment_date"
	doAssignmentByRef srchFields,pageObject.searchClauseObj.getSearchCtrlParams_p1("payment_date")
	Set firstFieldParams = (CreateDictionary())
	if bValue(asp_count(srchFields)) then
		doAssignment firstFieldParams,ArrayElement(srchFields,0)
	else
		setArrElement firstFieldParams,"fName","payment_date"
		setArrElement firstFieldParams,"eType",""
		setArrElement firstFieldParams,"value1",""
		setArrElement firstFieldParams,"opt",""
		setArrElement firstFieldParams,"value2",""
		setArrElement firstFieldParams,"not",false
	end if
	doAssignmentByRef ctrlBlockArr,searchControlBuilder.buildSearchCtrlBlockArr_p8(id,ArrayElement(firstFieldParams,"fName"),0,ArrayElement(firstFieldParams,"opt"),ArrayElement(firstFieldParams,"not"),false,ArrayElement(firstFieldParams,"value1"),ArrayElement(firstFieldParams,"value2"))
	if bValue(isEnableSection508()) then
		xt.assign_section_p3 "payment_date_label",("<label for=""" & CSmartStr(GetInputElementId("payment_date",id))) & """>","</label>"
	else
		xt.assign_p2 "payment_date_label",true
	end if
	xt.assign_p2 "payment_date_fieldblock",true
	xt.assignbyref_p2 "payment_date_editcontrol",ArrayElement(ctrlBlockArr,"searchcontrol")
	xt.assign_p2 "payment_date_notbox",ArrayElement(ctrlBlockArr,"notbox")
	xt.assignbyref_p2 "payment_date_editcontrol1",ArrayElement(ctrlBlockArr,"searchcontrol1")
	xt.assign_p2 "searchtype_payment_date",ArrayElement(ctrlBlockArr,"searchtype")
	doAssignmentByRef isFieldNeedSecCtrl,searchControlBuilder.isNeedSecondCtrl_p1("payment_date")
	ctrlInd = 0
	if bValue(isFieldNeedSecCtrl) then
		setArrElementN pageObject.controlsMap,CreateArray3("search","searchBlocks",empty),CreateDictionary3("fName","payment_date","recId",id,"ctrlsMap",CreateDictionary2(0,ctrlInd,1,CSmartDbl(ctrlInd)+1))
		ctrlInd = CSmartDbl(ctrlInd)+2
	else
		setArrElementN pageObject.controlsMap,CreateArray3("search","searchBlocks",empty),CreateDictionary3("fName","payment_date","recId",id,"ctrlsMap",CreateDictionary1(0,ctrlInd))
		ctrlInd = CSmartDbl(ctrlInd)+1
	end if
	if bValue(GetLookupTable("fop",strTableName)) then
		setArrElementN pageObject.settingsMap,CreateArray3("globalSettings","shortTNames",GetLookupTable("fop",strTableName)),GetTableURL(GetLookupTable("fop",strTableName))
	end if
	pageObject.fillFieldToolTips_p1 "fop"
	doAssignmentByRef srchFields,pageObject.searchClauseObj.getSearchCtrlParams_p1("fop")
	Set firstFieldParams = (CreateDictionary())
	if bValue(asp_count(srchFields)) then
		doAssignment firstFieldParams,ArrayElement(srchFields,0)
	else
		setArrElement firstFieldParams,"fName","fop"
		setArrElement firstFieldParams,"eType",""
		setArrElement firstFieldParams,"value1",""
		setArrElement firstFieldParams,"opt",""
		setArrElement firstFieldParams,"value2",""
		setArrElement firstFieldParams,"not",false
	end if
	doAssignmentByRef ctrlBlockArr,searchControlBuilder.buildSearchCtrlBlockArr_p8(id,ArrayElement(firstFieldParams,"fName"),0,ArrayElement(firstFieldParams,"opt"),ArrayElement(firstFieldParams,"not"),false,ArrayElement(firstFieldParams,"value1"),ArrayElement(firstFieldParams,"value2"))
	if bValue(isEnableSection508()) then
		xt.assign_section_p3 "fop_label",("<label for=""" & CSmartStr(GetInputElementId("fop",id))) & """>","</label>"
	else
		xt.assign_p2 "fop_label",true
	end if
	xt.assign_p2 "fop_fieldblock",true
	xt.assignbyref_p2 "fop_editcontrol",ArrayElement(ctrlBlockArr,"searchcontrol")
	xt.assign_p2 "fop_notbox",ArrayElement(ctrlBlockArr,"notbox")
	xt.assignbyref_p2 "fop_editcontrol1",ArrayElement(ctrlBlockArr,"searchcontrol1")
	xt.assign_p2 "searchtype_fop",ArrayElement(ctrlBlockArr,"searchtype")
	doAssignmentByRef isFieldNeedSecCtrl,searchControlBuilder.isNeedSecondCtrl_p1("fop")
	ctrlInd = 0
	if bValue(isFieldNeedSecCtrl) then
		setArrElementN pageObject.controlsMap,CreateArray3("search","searchBlocks",empty),CreateDictionary3("fName","fop","recId",id,"ctrlsMap",CreateDictionary2(0,ctrlInd,1,CSmartDbl(ctrlInd)+1))
		ctrlInd = CSmartDbl(ctrlInd)+2
	else
		setArrElementN pageObject.controlsMap,CreateArray3("search","searchBlocks",empty),CreateDictionary3("fName","fop","recId",id,"ctrlsMap",CreateDictionary1(0,ctrlInd))
		ctrlInd = CSmartDbl(ctrlInd)+1
	end if
	setArrElement pageObject.body,"begin",CSmartStr(ArrayElement(pageObject.body,"begin")) & CSmartStr(includes)
	pageObject.addCommonJs 
	xt.assignbyref_p2 "body",pageObject.body
	xt.assign_p2 "contents_block",true
	xt.assign_p2 "conditions_block",true
	xt.assign_p2 "search_button",true
	xt.assign_p2 "reset_button",true
	xt.assign_p2 "back_button",true
	xt.assign_p2 "searchbutton_attrs",("id=""searchButton" & CSmartStr(id)) & """"
	xt.assign_p2 "resetbutton_attrs",("id=""resetButton" & CSmartStr(id)) & """"
	xt.assign_p2 "backbutton_attrs",("id=""backButton" & CSmartStr(id)) & """"
	if bValue(eventObj.exists_p1("BeforeShowSearch")) then
		eventObj.BeforeShowSearch_p2 xt,templatefile
	end if
	pageObject.fillSetCntrlMaps 
	setArrElement pageObject.body,"end",CSmartStr(ArrayElement(pageObject.body,"end")) & "<script>"
	setArrElement pageObject.body,"end",CSmartStr(ArrayElement(pageObject.body,"end")) & (("window.controlsMap = '" & CSmartStr(jsreplace(my_json_encode(pageObject.controlsHTMLMap)))) & "';")
	setArrElement pageObject.body,"end",CSmartStr(ArrayElement(pageObject.body,"end")) & (("window.settings = '" & CSmartStr(jsreplace(my_json_encode(pageObject.jsSettings)))) & "';")
	setArrElement pageObject.body,"end",CSmartStr(ArrayElement(pageObject.body,"end")) & "</script>"
	setArrElement pageObject.body,"end",CSmartStr(ArrayElement(pageObject.body,"end")) & (("<script>" & CSmartStr(pageObject.PrepareJs())) & "</script>")
	xt.assignbyref_p2 "body",pageObject.body
	xt.display_p1 templatefile
	Response.End
else
	if IsEqual(mode,SEARCH_LOAD_CONTROL) then
		Set searchControlBuilder = (CreateClass("PanelSearchControl",4,searchControllerId,strTableName,pageObject.searchClauseObj,pageObject,Empty,Empty,Empty))
		doAssignmentByRef ctrlField,postvalue("ctrlField")
		doAssignmentByRef ctrlBlockArr,searchControlBuilder.buildSearchCtrlBlockArr_p8(id,ctrlField,0,"",false,true,"","")
		Set resArr = (CreateDictionary())
		setArrElement resArr,"control1",trim(xt.call_func_p1(ArrayElement(ctrlBlockArr,"searchcontrol")))
		setArrElement resArr,"control2",trim(xt.call_func_p1(ArrayElement(ctrlBlockArr,"searchcontrol1")))
		setArrElement resArr,"comboHtml",trim(ArrayElement(ctrlBlockArr,"searchtype"))
		setArrElement resArr,"delButt",trim(ArrayElement(ctrlBlockArr,"delCtrlButt"))
		setArrElement resArr,"delButtId",trim(searchControlBuilder.getDelButtonId_p2(ctrlField,id))
		setArrElement resArr,"divInd",trim(id)
		setArrElement resArr,"fLabel",GetFieldLabel(GoodFieldName(strTableName),GoodFieldName(ctrlField))
		setArrElement resArr,"ctrlMap",ArrayElement(pageObject.controlsMap,"controls")
		if IsEqual(postvalue("isNeedSettings"),"true") then
			pageObject.fillGlobalSettings 
			pageObject.fillTableSettings 
			setArrElementN resArr,CreateArray2("settings",ctrlField),pageObject.jsSettings
		end if
		ResponseWrite my_json_encode(resArr)
		Response.End
	end if
end if
%>
