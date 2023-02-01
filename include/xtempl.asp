<%
asp_include getabspath("include/menuitem.asp"),true
asp_include getabspath("include/testing.asp"),false

'------ Class XTempl ------
Class XTempl
	Public xt_vars
	Public xt_stack
	Public xt_events
	Public template
	Public charsets
	Public testingFlag
	Public eventsObject
	Public Function getvar_p1(ByVal name)
		DoAssignmentByRef getvar_p1,method_XTempl_getvar(me,name)
	End Function
	Public Function recTesting_p1(ByRef arr)
		DoAssignmentByRef recTesting_p1,method_XTempl_recTesting(me,arr)
	End Function
	Public Function Testing()
		DoAssignmentByRef Testing,method_XTempl_Testing(me)
	End Function
	Public Function report_error_p1(ByVal message)
		DoAssignmentByRef report_error_p1,method_XTempl_report_error(me,message)
	End Function
	Public Function init_XTempl()
		DoAssignmentByRef init_XTempl,method_XTempl_XTempl(me)
	End Function
	Public Function getReadingOrder()
		DoAssignmentByRef getReadingOrder,method_XTempl_getReadingOrder(me)
	End Function
	Public Function assign(ByVal name,ByVal val)
		DoAssignmentByRef assign,method_XTempl_assign(me,name,val)
	End Function
	Public Function assign_p2(ByVal name,ByVal val)
		DoAssignmentByRef assign_p2,method_XTempl_assign(me,name,val)
	End Function
	Public Function assignbyref(ByVal name,ByRef var)
		DoAssignmentByRef assignbyref,method_XTempl_assignbyref(me,name,var)
	End Function
	Public Function assignbyref_p2(ByVal name,ByRef var)
		DoAssignmentByRef assignbyref_p2,method_XTempl_assignbyref(me,name,var)
	End Function
	Public Function enable_section_p1(ByVal name)
		DoAssignmentByRef enable_section_p1,method_XTempl_enable_section(me,name)
	End Function
	Public Function assign_section(ByVal name,ByVal begin,ByVal var_end)
		DoAssignmentByRef assign_section,method_XTempl_assign_section(me,name,begin,var_end)
	End Function
	Public Function assign_section_p3(ByVal name,ByVal begin,ByVal var_end)
		DoAssignmentByRef assign_section_p3,method_XTempl_assign_section(me,name,begin,var_end)
	End Function
	Public Function assign_loopsection(ByVal name,ByRef data)
		DoAssignmentByRef assign_loopsection,method_XTempl_assign_loopsection(me,name,data)
	End Function
	Public Function assign_loopsection_p2(ByVal name,ByRef data)
		DoAssignmentByRef assign_loopsection_p2,method_XTempl_assign_loopsection(me,name,data)
	End Function
	Public Function assign_function(ByVal name,ByVal func,ByVal params)
		DoAssignmentByRef assign_function,method_XTempl_assign_function(me,name,func,params)
	End Function
	Public Function assign_function_p3(ByVal name,ByVal func,ByVal params)
		DoAssignmentByRef assign_function_p3,method_XTempl_assign_function(me,name,func,params)
	End Function
	Public Function assign_method(ByVal name,ByRef object,ByVal method,ByVal params)
		DoAssignmentByRef assign_method,method_XTempl_assign_method(me,name,object,method,params)
	End Function
	Public Function assign_method_p4(ByVal name,ByRef object,ByVal method,ByVal params)
		DoAssignmentByRef assign_method_p4,method_XTempl_assign_method(me,name,object,method,params)
	End Function
	Public Function assign_event(ByVal name,ByRef object,ByVal method,ByVal params)
		DoAssignmentByRef assign_event,method_XTempl_assign_event(me,name,object,method,params)
	End Function
	Public Function assign_event_p4(ByVal name,ByRef object,ByVal method,ByVal params)
		DoAssignmentByRef assign_event_p4,method_XTempl_assign_event(me,name,object,method,params)
	End Function
	Public Function xt_doevent_p1(ByVal params)
		DoAssignmentByRef xt_doevent_p1,method_XTempl_xt_doevent(me,params)
	End Function
	Public Function fetchVar_p1(ByVal varName)
		DoAssignmentByRef fetchVar_p1,method_XTempl_fetchVar(me,varName)
	End Function
	Public Function fetch_loaded_p1(ByVal filtertag)
		DoAssignmentByRef fetch_loaded_p1,method_XTempl_fetch_loaded(me,filtertag)
	End Function
	Public Function fetch_loaded()
		DoAssignmentByRef fetch_loaded,method_XTempl_fetch_loaded(me,"")
	End Function
	Public Function fetch_loaded_before_p1(ByVal filtertag)
		DoAssignmentByRef fetch_loaded_before_p1,method_XTempl_fetch_loaded_before(me,filtertag)
	End Function
	Public Function fetch_loaded_after_p1(ByVal filtertag)
		DoAssignmentByRef fetch_loaded_after_p1,method_XTempl_fetch_loaded_after(me,filtertag)
	End Function
	Public Function call_func_p1(ByVal var)
		DoAssignmentByRef call_func_p1,method_XTempl_call_func(me,var)
	End Function
	Public Function load_template_p1(ByVal template)
		DoAssignmentByRef load_template_p1,method_XTempl_load_template(me,template)
	End Function
	Public Function display_loaded_p1(ByVal filtertag)
		DoAssignmentByRef display_loaded_p1,method_XTempl_display_loaded(me,filtertag)
	End Function
	Public Function display_loaded()
		DoAssignmentByRef display_loaded,method_XTempl_display_loaded(me,"")
	End Function
	Public Function display_p1(ByVal template)
		DoAssignmentByRef display_p1,method_XTempl_display(me,template)
	End Function
	Public Function processVar_p2(ByRef var,ByRef varparams)
		DoAssignmentByRef processVar_p2,method_XTempl_processVar(me,var,varparams)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"xt_vars", xt_vars
		setArrElement out,"xt_stack", xt_stack
		setArrElement out,"xt_events", xt_events
		setArrElement out,"template", template
		setArrElement out,"charsets", charsets
		setArrElement out,"testingFlag", testingFlag
		setArrElement out,"eventsObject", eventsObject
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment xt_vars, dict("xt_vars")
		DoAssignment xt_stack, dict("xt_stack")
		DoAssignment xt_events, dict("xt_events")
		DoAssignment template, dict("template")
		DoAssignment charsets, dict("charsets")
		DoAssignment testingFlag, dict("testingFlag")
		DoAssignment eventsObject, dict("eventsObject")
	End Sub
' end serialize
End Class
' XTempl implementation 
Function method_XTempl_getvar(byref this_object,ByVal name)
	Dim this
	doAssignmentByRef method_XTempl_getvar,xt_getvar(this_object,name)
	Exit Function
End Function
Function method_XTempl_recTesting(byref this_object,ByRef arr)
	Dim v,this,k
	GetCollectionBounds arr,i_xtempl_loopIdx2,i_xtempl_loopMax2
	do while i_xtempl_loopIdx2<=i_xtempl_loopMax2
		k = GetCollectionKey(arr,i_xtempl_loopIdx2)
		doAssignment v,ArrayElement(arr,k)
		if bValue(asp_is_array(v)) then
			this_object.recTesting_p1 ArrayElement(arr,k)
		else
			if bValue(asp_array_key_exists(k,testingLinks)) then
				setArrElement arr,k,CSmartStr(ArrayElement(arr,k)) & ((" func=""" & CSmartStr(ArrayElement(testingLinks,k))) & """")
			end if
		end if
		i_xtempl_loopIdx2=i_xtempl_loopIdx2+1
	loop
End Function
Function method_XTempl_Testing(byref this_object)
	Dim this
	if not bValue(this_object.testingFlag) then
		Exit Function
	end if
	this_object.recTesting_p1 this_object.xt_vars
End Function
Function method_XTempl_report_error(byref this_object,ByVal message)
	ResponseWrite message
	Response.End
End Function
Function method_XTempl_XTempl(byref this_object)
	doClassAssignment this_object,"xt_vars",CreateDictionary()
	doClassAssignment this_object,"xt_events",CreateDictionary()
	doClassAssignment this_object,"charsets",CreateDictionary()
	this_object.testingFlag = false
	Dim this,mlang_charsets,order,html_attrs
	doClassAssignment this_object,"xt_vars",CreateDictionary()
	doClassAssignment this_object,"xt_stack",CreateDictionary()
	setArrElementByRef this_object.xt_stack,asp_count(this_object.xt_stack),this_object.xt_vars
	xtempl_include_header this_object,"header","include/header.asp"
	xtempl_include_header this_object,"footer","include/footer.asp"
	this_object.assign_method_p4 "event",this_object,"xt_doevent",CreateDictionary()
	this_object.assign_function_p3 "label","xt_label",CreateDictionary()
	this_object.assign_function_p3 "custom","xt_custom",CreateDictionary()
	this_object.assign_function_p3 "caption","xt_caption",CreateDictionary()
	this_object.assign_function_p3 "mainmenu","xt_displaymenu",CreateDictionary()
	this_object.assign_function_p3 "quickjump_options","xt_displaymenu",CreateDictionary1("quickjump",true)
	this_object.assign_function_p3 "TabGroup","xt_displaytabs",CreateDictionary()
	this_object.assign_function_p3 "Section","xt_displaytabs",CreateDictionary()
	Set mlang_charsets = (CreateDictionary())
	mlang_charsets.Add "English","Windows-1252"

	doClassAssignmentByRef this_object,"charsets",mlang_charsets
	doAssignmentByRef order,this_object.getReadingOrder()
	html_attrs = ""
	if IsEqual(order,"RTL") then
		this_object.assign_p2 "RTL_block",true
		html_attrs = CSmartStr(html_attrs) & "dir=""RTL"" "
	else
		this_object.assign_p2 "LTR_block",true
	end if
	if IsEqual(mlang_getcurrentlang(),"English") then
		html_attrs = CSmartStr(html_attrs) & "lang=""en"""
	end if
	this_object.assign_p2 "html_attrs",html_attrs
End Function
Function method_XTempl_getReadingOrder(byref this_object)
	Dim var_REQUEST,charset,var_SESSION,cp
	if bValue(GetRequestValue(Request,"language")) then
		doAssignment charset,ArrayElement(this_object.charsets,GetRequestValue(Request,"language"))
	else
		if bValue(Session("language")) then
			doAssignment charset,ArrayElement(this_object.charsets,Session("language"))
		else
			doAssignment charset,ArrayElement(this_object.charsets,"English")
		end if
	end if
	doAssignmentByRef cp,asp_strtolower(charset)
	if IsEqual(cp,"windows-1256") or IsEqual(cp,"windows-1255") then
		method_XTempl_getReadingOrder = "RTL"
		Exit Function
	else
		method_XTempl_getReadingOrder = "LTR"
		Exit Function
	end if
End Function
Function method_XTempl_assign(byref this_object,ByVal name,ByVal val)
	setArrElement this_object.xt_vars,name,val
End Function
Function method_XTempl_assignbyref(byref this_object,ByVal name,ByRef var)
	setArrElementByRef this_object.xt_vars,name,var
End Function
Function method_XTempl_enable_section(byref this_object,ByVal name)
	if not bValue(asp_array_key_exists(name,this_object.xt_vars)) then
		setArrElement this_object.xt_vars,name,true
	else
		if IsEqual(ArrayElement(this_object.xt_vars,name),false) then
			setArrElement this_object.xt_vars,name,true
		end if
	end if
End Function
Function method_XTempl_assign_section(byref this_object,ByVal name,ByVal begin,ByVal var_end)
	Dim arr
	Set arr = (CreateDictionary())
	setArrElement arr,"begin",begin
	setArrElement arr,"end",var_end
	setArrElementByRef this_object.xt_vars,name,arr
End Function
Function method_XTempl_assign_loopsection(byref this_object,ByVal name,ByRef data)
	Dim arr
	Set arr = (CreateDictionary())
	setArrElementByRef arr,"data",data
	setArrElementByRef this_object.xt_vars,name,arr
End Function
Function method_XTempl_assign_function(byref this_object,ByVal name,ByVal func,ByVal params)
	setArrElement this_object.xt_vars,name,CreateDictionary2("func",func,"params",params)
End Function
Function method_XTempl_assign_method(byref this_object,ByVal name,ByRef object,ByVal method,ByVal params)
	setArrElement this_object.xt_vars,name,CreateDictionary2("method",method,"params",params)
	setArrElementByRefN this_object.xt_vars,CreateArray2(name,"object"),object
End Function
Function method_XTempl_assign_event(byref this_object,ByVal name,ByRef object,ByVal method,ByVal params)
	setArrElement this_object.xt_events,name,CreateDictionary2("method",method,"params",params)
	setArrElementByRefN this_object.xt_events,CreateArray2(name,"object"),object
End Function
Function method_XTempl_xt_doevent(byref this_object,ByVal params)
	Dim eventArr,method,eventObj,eventName
	if bValue(asp_array_key_exists(ArrayElement(params,"custom1"),this_object.xt_events)) then
		doAssignment eventArr,ArrayElement(this_object.xt_events,ArrayElement(params,"custom1"))
		if bValue(asp_array_key_exists("method",eventArr)) then
			Set params = (CreateDictionary())
			if bValue(asp_array_key_exists("params",eventArr)) then
				doAssignment params,ArrayElement(eventArr,"params")
			end if
			doAssignment method,ArrayElement(eventArr,"method")
			callVariableMethod ArrayElement(eventArr,"object"),method,CreateArray1(params)
			Exit Function
		end if
	end if
	if bValue(this_object.eventsObject) then
		doAssignmentByRef eventObj,this_object.eventsObject
	else
		if bValue(asp_strlen(strTableName)) then
			doAssignmentByRef eventObj,getEventObject(strTableName)
		else
			doAssignmentByRef eventObj,globalEvents
		end if
	end if
	if not bValue(eventObj) then
		Exit Function
	end if
	doAssignment eventName,ArrayElement(params,"custom1")
	if not bValue(eventObj.exists_p1(eventName)) then
		Exit Function
	end if
	callVariableMethod eventObj,eventName,CreateArray1(params)
End Function
Function method_XTempl_fetchVar(byref this_object,ByVal varName)
	Dim varParams,this,out
	ob_start 
	Set varParams = (CreateDictionary())
	this_object.processVar_p2 this_object.getVar_p1(varName),varParams
	doAssignmentByRef out,ob_get_contents()
	ob_end_clean 
	doAssignmentByRef method_XTempl_fetchVar,out
	Exit Function
End Function
Function method_XTempl_fetch_loaded(byref this_object,ByVal filtertag)
	Dim this,out
	ob_start 
	this_object.display_loaded_p1 filtertag
	doAssignmentByRef out,ob_get_contents()
	ob_end_clean 
	doAssignmentByRef method_XTempl_fetch_loaded,out
	Exit Function
End Function
Function method_XTempl_fetch_loaded_before(byref this_object,ByVal filtertag)
	Dim pos1,str,this,out
	doAssignmentByRef pos1,asp_strpos(this_object.template,("{BEGIN " & CSmartStr(filtertag)) & "}",empty)
	if IsFalse(pos1) then
		Exit Function
	end if
	doAssignmentByRef str,asp_substr(this_object.template,0,pos1)
	ob_start 
	this_object.Testing 
	xt_process_template this_object,str
	doAssignmentByRef out,ob_get_contents()
	ob_end_clean 
	doAssignmentByRef method_XTempl_fetch_loaded_before,out
	Exit Function
End Function
Function method_XTempl_fetch_loaded_after(byref this_object,ByVal filtertag)
	Dim pos2,str,this,out
	doAssignmentByRef pos2,asp_strpos(this_object.template,("{END " & CSmartStr(filtertag)) & "}",empty)
	if IsFalse(pos2) then
		Exit Function
	end if
	doAssignmentByRef str,asp_substr(this_object.template,CSmartDbl(pos2)+CSmartDbl(asp_strlen(("{END " & CSmartStr(filtertag)) & "}")),empty)
	ob_start 
	this_object.Testing 
	xt_process_template this_object,str
	doAssignmentByRef out,ob_get_contents()
	ob_end_clean 
	doAssignmentByRef method_XTempl_fetch_loaded_after,out
	Exit Function
End Function
Function method_XTempl_call_func(byref this_object,ByVal var)
	doAssignmentByRef method_XTempl_call_func,xtempl_get_func_output(var)
	Exit Function
End Function
Function method_XTempl_load_template(byref this_object,ByVal template)
	doClassAssignment this_object,"template",myfile_get_contents(getabspath("templates/" & CSmartStr(template)),"rb")
End Function
Function method_XTempl_display_loaded(byref this_object,ByVal filtertag)
	Dim str,pos1,pos2,this
	doAssignment str,this_object.template
	if bValue(filtertag) then
		doAssignmentByRef pos1,asp_strpos(this_object.template,("{BEGIN " & CSmartStr(filtertag)) & "}",empty)
		doAssignmentByRef pos2,asp_strpos(this_object.template,("{END " & CSmartStr(filtertag)) & "}",empty)
		if IsFalse(pos1) or IsFalse(pos2) then
			Exit Function
		end if
		pos2 = CSmartDbl(pos2)+CSmartDbl(asp_strlen(("{END " & CSmartStr(filtertag)) & "}"))
		doAssignmentByRef str,asp_substr(this_object.template,pos1,CSmartDbl(pos2)-CSmartDbl(pos1))
	end if
	this_object.Testing 
	xt_process_template this_object,str
End Function
Function method_XTempl_display(byref this_object,ByVal template)
	Dim this
	this_object.load_template_p1 template
	this_object.Testing 
	xt_process_template this_object,this_object.template
End Function
Function method_XTempl_processVar(byref this_object,ByRef var,ByRef varparams)
	Dim params,key,val,func,method,this
	if not bValue(asp_is_array(var)) then
		ResponseWrite var
	else
		if bValue(asp_array_key_exists("func",var)) then
			Set params = (CreateDictionary())
			if bValue(asp_array_key_exists("params",var)) then
				doAssignment params,ArrayElement(var,"params")
			end if
			GetCollectionBounds varparams,i_xtempl_loopIdx3,i_xtempl_loopMax3
			do while i_xtempl_loopIdx3<=i_xtempl_loopMax3
				key = GetCollectionKey(varparams,i_xtempl_loopIdx3)
				doAssignment val,ArrayElement(varparams,key)
				setArrElement params,"custom" & CSmartStr(key),val
				i_xtempl_loopIdx3=i_xtempl_loopIdx3+1
			loop
			doAssignment func,ArrayElement(var,"func")
			xtempl_call_func func,params
		else
			if bValue(asp_array_key_exists("method",var)) then
				Set params = (CreateDictionary())
				if bValue(asp_array_key_exists("params",var)) then
					doAssignment params,ArrayElement(var,"params")
				end if
				GetCollectionBounds varparams,i_xtempl_loopIdx4,i_xtempl_loopMax4
				do while i_xtempl_loopIdx4<=i_xtempl_loopMax4
					key = GetCollectionKey(varparams,i_xtempl_loopIdx4)
					doAssignment val,ArrayElement(varparams,key)
					setArrElement params,"custom" & CSmartStr(key),val
					i_xtempl_loopIdx4=i_xtempl_loopIdx4+1
				loop
				doAssignment method,ArrayElement(var,"method")
				callVariableMethod ArrayElement(var,"object"),method,CreateArray1(params)
			else
				this_object.report_error_p1 "Incorrect variable value - " & CSmartStr(var_name)
				Exit Function
			end if
		end if
	end if
End Function
Set menuparams = (CreateDictionary())
Function xt_displaymenu(ByVal params)
	doAssignment menuparams,params
	asp_include getabspath("include/displaymenu.asp"),false
End Function
Set tabparams = (CreateDictionary())
Function xt_displaytabs(ByVal params)
	Dim savedTemplate
	doAssignment tabparams,params
	doAssignment savedTemplate,xt.template
	asp_include getabspath("include/displaytabs.asp"),false
	doClassAssignment xt,"template",savedTemplate
End Function
Function xt_buildeditcontrol(ByRef params)
	Dim field,mode,fieldNum,id,format,append,validate,additionalCtrlParams,pageObj
	doAssignment field,ArrayElement(params,"field")
	if IsEqual(ArrayElement(params,"mode"),"edit") then
		mode = MODE_EDIT
	else
		if IsEqual(ArrayElement(params,"mode"),"add") then
			mode = MODE_ADD
		else
			if IsEqual(ArrayElement(params,"mode"),"inline_edit") then
				mode = MODE_INLINE_EDIT
			else
				if IsEqual(ArrayElement(params,"mode"),"inline_add") then
					mode = MODE_INLINE_ADD
				else
					mode = MODE_SEARCH
				end if
			end if
		end if
	end if
	fieldNum = 0
	if bValue(ArrayElement(params,"fieldNum")) then
		doAssignment fieldNum,ArrayElement(params,"fieldNum")
	end if
	id = ""
	if not IsIdentical(ArrayElement(params,"id"),"") then
		doAssignment id,ArrayElement(params,"id")
	end if
	doAssignmentByRef format,GetEditFormat(field,"")
	if not IsEqual(ArrayElement(params,"format"),"") then
		doAssignment format,ArrayElement(params,"format")
	end if
	append = ""
	if (((IsEqual(mode,MODE_EDIT) or IsEqual(mode,MODE_ADD)) or IsEqual(mode,MODE_INLINE_EDIT)) or IsEqual(mode,MODE_INLINE_ADD)) and IsEqual(format,EDIT_FORMAT_READONLY) then
		ResponseWrite ArrayElement(readonlyfields,field)
	end if
	if IsEqual(mode,MODE_SEARCH) then
		doAssignment format,ArrayElement(params,"format")
	end if
	Set validate = (CreateDictionary())
	if bValue(asp_count(ArrayElement(params,"validate"))) then
		doAssignment validate,ArrayElement(params,"validate")
	end if
	Set additionalCtrlParams = (CreateDictionary())
	if bValue(asp_count(ArrayElement(params,"additionalCtrlParams"))) then
		doAssignment additionalCtrlParams,ArrayElement(params,"additionalCtrlParams")
	end if
	doAssignment pageObj,IIF(not IsEmpty(ArrayElement(params,"pageObj")),ArrayElement(params,"pageObj"),null)
	BuildEditControl field,ArrayElement(params,"value"),format,mode,fieldNum,id,validate,additionalCtrlParams,pageObj
End Function
Function xt_showchart(ByVal params)
	Dim width,height,refresh,var_SERVER,http
	width = 700
	height = 530
	if bValue(asp_array_key_exists("custom1",params)) then
		doAssignment width,ArrayElement(params,"custom1")
	end if
	if bValue(asp_array_key_exists("custom2",params)) then
		doAssignment height,ArrayElement(params,"custom2")
	end if
	refresh = CSmartDbl(GetTableData(ArrayElement(params,"table"),".ChartRefreshTime",10))*60000
	if IsEqual(GetRequestValue(Request.ServerVariables,"SERVER_PORT"),443) then
		http = "https"
	else
		http = "http"
	end if
	ResponseWrite "<div id='"
	ResponseWrite ArrayElement(params,"chartname")
	ResponseWrite "'>" & vbcrlf & _
		"<noscript>" & vbcrlf & _
		"	<object id="""
	ResponseWrite ArrayElement(params,"chartname")
	ResponseWrite """ " & vbcrlf & _
		"			name="""
	ResponseWrite ArrayElement(params,"chartname")
	ResponseWrite """ " & vbcrlf & _
		"			classid=""clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"" " & vbcrlf & _
		"			width=""100%"" " & vbcrlf & _
		"			height=""100%"" " & vbcrlf & _
		"			codebase="""
	ResponseWrite http
	ResponseWrite "://fpdownload.macromedia.com/get/flashplayer/current/swflash.cab"">" & vbcrlf & _
		"		<param name=""movie"" value=""libs/swf/Preloader.swf"" />" & vbcrlf & _
		"		<param name=""bgcolor"" value=""#FFFFFF"" />" & vbcrlf & _
		"		<param name=""wmode"" value=""opaque"" />" & vbcrlf & _
		"		<param name=""allowScriptAccess"" value=""always"" />" & vbcrlf & _
		"		<param name=""flashvars"" value=""swfFile="
	ResponseWrite "dchartdata.asp%3Fchartname%3D" & CSmartStr(ArrayElement(params,"chartname"))
	ResponseWrite """ />" & vbcrlf & _
		"		" & vbcrlf & _
		"		<embed type=""application/x-shockwave-flash"" " & vbcrlf & _
		"			   pluginspage="""
	ResponseWrite http
	ResponseWrite "://www.adobe.com/go/getflashplayer"" " & vbcrlf & _
		"			   src=""libs/swf/Preloader.swf"" " & vbcrlf & _
		"			   width=""100%"" " & vbcrlf & _
		"			   height=""100%"" " & vbcrlf & _
		"			   id="""
	ResponseWrite ArrayElement(params,"chartname")
	ResponseWrite """ " & vbcrlf & _
		"			   name="""
	ResponseWrite ArrayElement(params,"chartname")
	ResponseWrite """ " & vbcrlf & _
		"			   bgColor=""#FFFFFF"" " & vbcrlf & _
		"			   allowScriptAccess=""always"" " & vbcrlf & _
		"			   flashvars=""swfFile="
	ResponseWrite "dchartdata.asp%3Fchartname%3D" & CSmartStr(ArrayElement(params,"chartname"))
	ResponseWrite """ />" & vbcrlf & _
		"	</object>				" & vbcrlf & _
		"</noscript>" & vbcrlf
	ResponseWrite "<s"
	ResponseWrite "cript type=""text/javascript"">" & vbcrlf & _
		"	//<![CDATA[" & vbcrlf & _
		"	document.write('<center>');" & vbcrlf & _
		"	document.write(""You need to have Adobe Flash Player 9 (or above) to view the chart.<br /><br />"");" & vbcrlf & _
		"	document.write(""<a href=\"""
	ResponseWrite http
	ResponseWrite "://www.adobe.com/go/getflashplayer\""><img border=\""0\"" src=\"""
	ResponseWrite http
	ResponseWrite "://www.adobe.com/images/shared/download_buttons/get_flash_player.gif\"" /></a><br />"");" & vbcrlf & _
		"	document.write('</center>');" & vbcrlf & _
		"	//]]>" & vbcrlf & _
		"</script>" & vbcrlf
	ResponseWrite "<s"
	ResponseWrite "cript type=""text/javascript"" language=""javascript"" src=""libs/js/AnyChart.js""></script>" & vbcrlf
	ResponseWrite "<s"
	ResponseWrite "cript type=""text/javascript"" language=""javascript"">" & vbcrlf & _
		"	//<![CDATA[" & vbcrlf & _
		"	var chart = new AnyChart('libs/swf/AnyChart.swf','libs/swf/Preloader.swf');" & vbcrlf & _
		"	chart.width = '"
	ResponseWrite width
	ResponseWrite "';" & vbcrlf & _
		"	chart.height = '"
	ResponseWrite height
	ResponseWrite "';" & vbcrlf & _
		"	chart.wMode='opaque';" & vbcrlf & vbcrlf & _
		"	var xmlFile = 'dchartdata.asp%3Fchartname%3D"
	ResponseWrite jsreplace(ArrayElement(params,"chartname"))
	ResponseWrite "';" & vbcrlf & _
		"	xmlFile += '%26ctype%3D"
	ResponseWrite ArrayElement(params,"ctype")
	ResponseWrite "';" & vbcrlf & _
		"	chart.setXMLFile(xmlFile);" & vbcrlf & _
		"	chart.write('"
	ResponseWrite ArrayElement(params,"chartname")
	ResponseWrite "');" & vbcrlf & _
		"	if("""
	ResponseWrite refresh
	ResponseWrite """!=""0"")" & vbcrlf & _
		"		setInterval('refreshChart()',"
	ResponseWrite refresh
	ResponseWrite ");" & vbcrlf & _
		"	function refreshChart()" & vbcrlf & _
		"	{" & vbcrlf & _
		"		page='dchartdata.asp?chartname="
	ResponseWrite jsreplace(ArrayElement(params,"chartname"))
	ResponseWrite "';" & vbcrlf & _
		"		params={" & vbcrlf & _
		"				action:'refresh'," & vbcrlf & _
		"				rndval:Math.random()" & vbcrlf & _
		"				};" & vbcrlf & _
		"		$.get(page,params,function(xml)" & vbcrlf & _
		"			{" & vbcrlf & _
		"				var arr = new Array();" & vbcrlf & _
		"				arr=xml.split(""\n"");" & vbcrlf & _
		"				for(i=0; i<arr.length;i+=2)" & vbcrlf & _
		"				{" & vbcrlf & _
		"					chart.removeSeries(arr[i]);" & vbcrlf & _
		"					chart.addSeries(arr[i+1]);" & vbcrlf & _
		"					chart.updatePointData(arr[i]+""_gauge"",arr[i]+""_point"",{value: arr[i+1]});" & vbcrlf & _
		"				}" & vbcrlf & _
		"				chart.refresh();" & vbcrlf & _
		"			});" & vbcrlf & vbcrlf & _
		"	}" & vbcrlf & _
		"	//"
	ResponseWrite "]]>" & vbcrlf & _
		"</script>" & vbcrlf & _
		"</div>" & vbcrlf
End Function
Function xt_include(ByVal params)
	if bValue(asp_file_exists(getabspath(ArrayElement(params,"file")))) then
		asp_include getabspath(ArrayElement(params,"file")),false
	end if
End Function
Function xt_label(ByVal params)
	ResponseWrite htmlspecialchars(GetFieldLabel(ArrayElement(params,"custom1"),ArrayElement(params,"custom2")))
End Function
Function xt_custom(ByVal params)
	ResponseWrite GetCustomLabel(ArrayElement(params,"custom1"))
End Function
Function xt_caption(ByVal params)
	ResponseWrite GetTableCaption(ArrayElement(params,"custom1"))
End Function
Function xtempl_get_func_output(ByRef var)
	Dim params,func,out
	if not bValue(asp_strlen(ArrayElement(var,"func"))) then
		xtempl_get_func_output = ""
		Exit Function
	end if
	ob_start 
	doAssignment params,ArrayElement(var,"params")
	doAssignment func,ArrayElement(var,"func")
	xtempl_call_func func,params
	doAssignmentByRef out,ob_get_contents()
	ob_end_clean 
	doAssignmentByRef xtempl_get_func_output,out
	Exit Function
End Function
%>
