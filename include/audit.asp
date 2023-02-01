<%

'------ Class AuditTrailTable ------
Class AuditTrailTable
	Public logTableName
	Public TableObj
	Public params
	Public strLogin
	Public strFailLogin
	Public strLogout
	Public strChPass
	Public strAdd
	Public strEdit
	Public strDelete
	Public strAccess
	Public strKeysHeader
	Public strFieldsHeader
	Public columnDate
	Public columnTime
	Public columnIP
	Public columnUser
	Public columnTable
	Public columnAction
	Public columnKey
	Public columnField
	Public columnOldValue
	Public columnNewValue
	Public attLogin
	Public timeLogin
	Public Function init_AuditTrailTable()
		DoAssignmentByRef init_AuditTrailTable,method_AuditTrailTable_AuditTrailTable(me)
	End Function
	Public Function LogLogin()
		DoAssignmentByRef LogLogin,method_AuditTrailTable_LogLogin(me)
	End Function
	Public Function LogLoginFailed_p1(ByVal pUsername)
		DoAssignmentByRef LogLoginFailed_p1,method_AuditTrailTable_LogLoginFailed(me,pUsername)
	End Function
	Public Function LogLogout()
		DoAssignmentByRef LogLogout,method_AuditTrailTable_LogLogout(me)
	End Function
	Public Function LogChPassword()
		DoAssignmentByRef LogChPassword,method_AuditTrailTable_LogChPassword(me)
	End Function
	Public Function LogAdd_p3(ByVal str_table,ByVal values,ByVal keys)
		DoAssignmentByRef LogAdd_p3,method_AuditTrailTable_LogAdd(me,str_table,values,keys)
	End Function
	Public Function LogEdit_p4(ByVal str_table,ByVal newvalues,ByVal oldvalues,ByVal keys)
		DoAssignmentByRef LogEdit_p4,method_AuditTrailTable_LogEdit(me,str_table,newvalues,oldvalues,keys)
	End Function
	Public Function LogDelete_p3(ByVal str_table,ByVal values,ByVal keys)
		DoAssignmentByRef LogDelete_p3,method_AuditTrailTable_LogDelete(me,str_table,values,keys)
	End Function
	Public Function LogAddEvent_p3(ByVal message,ByVal description,ByVal stable)
		DoAssignmentByRef LogAddEvent_p3,method_AuditTrailTable_LogAddEvent(me,message,description,stable)
	End Function
	Public Function LogAddEvent_p2(ByVal message,ByVal description)
		DoAssignmentByRef LogAddEvent_p2,method_AuditTrailTable_LogAddEvent(me,message,description,"")
	End Function
	Public Function LogAddEvent_p1(ByVal message)
		DoAssignmentByRef LogAddEvent_p1,method_AuditTrailTable_LogAddEvent(me,message,"","")
	End Function
	Public Function LoginSuccessful()
		DoAssignmentByRef LoginSuccessful,method_AuditTrailTable_LoginSuccessful(me)
	End Function
	Public Function LoginUnsuccessful_p1(ByVal pUsername)
		DoAssignmentByRef LoginUnsuccessful_p1,method_AuditTrailTable_LoginUnsuccessful(me,pUsername)
	End Function
	Public Function LoginAccess()
		DoAssignmentByRef LoginAccess,method_AuditTrailTable_LoginAccess(me)
	End Function
	Public Function logValueEnable_p1(ByVal table)
		DoAssignmentByRef logValueEnable_p1,method_AuditTrailTable_logValueEnable(me,table)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"logTableName", logTableName
		setArrElement out,"TableObj", TableObj
		setArrElement out,"params", params
		setArrElement out,"strLogin", strLogin
		setArrElement out,"strFailLogin", strFailLogin
		setArrElement out,"strLogout", strLogout
		setArrElement out,"strChPass", strChPass
		setArrElement out,"strAdd", strAdd
		setArrElement out,"strEdit", strEdit
		setArrElement out,"strDelete", strDelete
		setArrElement out,"strAccess", strAccess
		setArrElement out,"strKeysHeader", strKeysHeader
		setArrElement out,"strFieldsHeader", strFieldsHeader
		setArrElement out,"columnDate", columnDate
		setArrElement out,"columnTime", columnTime
		setArrElement out,"columnIP", columnIP
		setArrElement out,"columnUser", columnUser
		setArrElement out,"columnTable", columnTable
		setArrElement out,"columnAction", columnAction
		setArrElement out,"columnKey", columnKey
		setArrElement out,"columnField", columnField
		setArrElement out,"columnOldValue", columnOldValue
		setArrElement out,"columnNewValue", columnNewValue
		setArrElement out,"attLogin", attLogin
		setArrElement out,"timeLogin", timeLogin
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment logTableName, dict("logTableName")
		DoAssignment TableObj, dict("TableObj")
		DoAssignment params, dict("params")
		DoAssignment strLogin, dict("strLogin")
		DoAssignment strFailLogin, dict("strFailLogin")
		DoAssignment strLogout, dict("strLogout")
		DoAssignment strChPass, dict("strChPass")
		DoAssignment strAdd, dict("strAdd")
		DoAssignment strEdit, dict("strEdit")
		DoAssignment strDelete, dict("strDelete")
		DoAssignment strAccess, dict("strAccess")
		DoAssignment strKeysHeader, dict("strKeysHeader")
		DoAssignment strFieldsHeader, dict("strFieldsHeader")
		DoAssignment columnDate, dict("columnDate")
		DoAssignment columnTime, dict("columnTime")
		DoAssignment columnIP, dict("columnIP")
		DoAssignment columnUser, dict("columnUser")
		DoAssignment columnTable, dict("columnTable")
		DoAssignment columnAction, dict("columnAction")
		DoAssignment columnKey, dict("columnKey")
		DoAssignment columnField, dict("columnField")
		DoAssignment columnOldValue, dict("columnOldValue")
		DoAssignment columnNewValue, dict("columnNewValue")
		DoAssignment attLogin, dict("attLogin")
		DoAssignment timeLogin, dict("timeLogin")
	End Sub
' end serialize
End Class
' AuditTrailTable implementation 
Function method_AuditTrailTable_AuditTrailTable(byref this_object)
	this_object.logTableName = ""
	this_object.strLogin = "login"
	this_object.strFailLogin = "failed login"
	this_object.strLogout = "logout"
	this_object.strChPass = "change password"
	this_object.strAdd = "add"
	this_object.strEdit = "edit"
	this_object.strDelete = "delete"
	this_object.strAccess = "access"
	this_object.strKeysHeader = "---Keys"
	this_object.strFieldsHeader = "---Fields"
	this_object.columnDate = "Date"
	this_object.columnTime = "Time"
	this_object.columnIP = "IP"
	this_object.columnUser = "User"
	this_object.columnTable = "Table"
	this_object.columnAction = "Action"
	this_object.columnKey = "Key field"
	this_object.columnField = "Field"
	this_object.columnOldValue = "Old value"
	this_object.columnNewValue = "New value"
	this_object.attLogin = 0
	this_object.timeLogin = 0
	Dim userid,var_SESSION,var_SERVER
	doClassAssignmentByRef this_object,"TableObj",dal.Table(this_object.logTableName)
	userid = ""
	if bValue(Session("UserID")) then
		doAssignment userid,Session("UserID")
	end if
	doClassAssignment this_object,"params",CreateDictionary2(Empty,GetRequestValue(Request.ServerVariables,"REMOTE_ADDR"),Empty,userid)
End Function
Function method_AuditTrailTable_LogLogin(byref this_object)
	Dim retval,table,arr,var_SERVER,var_SESSION
End Function
Function method_AuditTrailTable_LogLoginFailed(byref this_object,ByVal pUsername)
	Dim retval,table,arr,var_SERVER
End Function
Function method_AuditTrailTable_LogLogout(byref this_object)
	Dim var_SESSION,retval,table,arr
End Function
Function method_AuditTrailTable_LogChPassword(byref this_object)
	Dim retval,table,arr
End Function
Function method_AuditTrailTable_LogAdd(byref this_object,ByVal str_table,ByVal values,ByVal keys)
	Dim retval,table,arr,str,idx,val,strFields,this,v
	retval = true
	doAssignment table,str_table
	Set arr = (CreateDictionary())
	if bValue(globalEvents.exists_p1("OnAuditLog")) then
		doAssignmentByRef retval,globalEvents.OnAuditLog_p6(this_object.strAdd,this_object.params,table,keys,values,arr)
	end if
	if bValue(retval) then
		str = ""
		if IsLess(0,asp_count(keys)) then
			str = CSmartStr(str) & (CSmartStr(this_object.strKeysHeader) & vbcrlf)
			GetCollectionBounds keys,i_audit_loopIdx2,i_audit_loopMax2
			do while i_audit_loopIdx2<=i_audit_loopMax2
				idx = GetCollectionKey(keys,i_audit_loopIdx2)
				doAssignment val,ArrayElement(keys,idx)
				str = CSmartStr(str) & (((CSmartStr(idx) & " : ") & CSmartStr(val)) & vbcrlf)
				i_audit_loopIdx2=i_audit_loopIdx2+1
			loop
		end if
		strFields = ""
		if bValue(this_object.logValueEnable_p1(str_table)) then
			GetCollectionBounds values,i_audit_loopIdx3,i_audit_loopMax3
			do while i_audit_loopIdx3<=i_audit_loopMax3
				idx = GetCollectionKey(values,i_audit_loopIdx3)
				doAssignment val,ArrayElement(values,idx)
				if not IsEqual(val,"") and not bValue(asp_array_key_exists(idx,keys)) then
					strFields = CSmartStr(strFields) & (CSmartStr(idx) & " [new]: ")
					if bValue(IsBinaryType(GetFieldType(idx,str_table))) then
						v = "<binary value>"
					else
						doAssignmentByRef v,asp_str_replace(CreateDictionary3(Empty,vbcrlf,Empty,vblf,Empty,"	")," ",val)
						if IsLess(300,asp_strlen(v)) then
							doAssignmentByRef v,asp_substr(val,0,300)
						end if
					end if
					strFields = CSmartStr(strFields) & (CSmartStr(v) & vbcrlf)
				end if
				i_audit_loopIdx3=i_audit_loopIdx3+1
			loop
		end if
		if not IsEqual(strFields,"") then
			str = CSmartStr(str) & ((CSmartStr(this_object.strFieldsHeader) & vbcrlf) & CSmartStr(strFields))
		end if
		doClassAssignment this_object.TableObj,"datetime",now()
		doClassAssignment this_object.TableObj,"ip",ArrayElement(this_object.params,0)
		doClassAssignment this_object.TableObj,"user",ArrayElement(this_object.params,1)
		doClassAssignment this_object.TableObj,"table",str_table
		doClassAssignment this_object.TableObj,"action",this_object.strAdd
		doClassAssignment this_object.TableObj,"description",str
		this_object.TableObj.Add 
	end if
	doAssignmentByRef method_AuditTrailTable_LogAdd,retval
	Exit Function
End Function
Function method_AuditTrailTable_LogEdit(byref this_object,ByVal str_table,ByVal newvalues,ByVal oldvalues,ByVal keys)
	Dim retval,table,str,idx,val,strFields,this,v,var_type,newvalue,oldvalue
	retval = true
	doAssignment table,str_table
	if bValue(globalEvents.exists_p1("OnAuditLog")) then
		doAssignmentByRef retval,globalEvents.OnAuditLog_p6(this_object.strEdit,this_object.params,table,keys,newvalues,oldvalues)
	end if
	if bValue(retval) then
		str = ""
		if IsLess(0,asp_count(keys)) then
			str = CSmartStr(str) & (CSmartStr(this_object.strKeysHeader) & vbcrlf)
			GetCollectionBounds newvalues,i_audit_loopIdx4,i_audit_loopMax4
			do while i_audit_loopIdx4<=i_audit_loopMax4
				idx = GetCollectionKey(newvalues,i_audit_loopIdx4)
				doAssignment val,ArrayElement(newvalues,idx)
				if bValue(asp_array_key_exists(idx,keys)) then
					if not IsEqual(val,ArrayElement(oldvalues,idx)) then
						str = CSmartStr(str) & (((CSmartStr(idx) & " [old]: ") & CSmartStr(ArrayElement(oldvalues,idx))) & vbcrlf)
						str = CSmartStr(str) & (((CSmartStr(idx) & " [new]: ") & CSmartStr(val)) & vbcrlf)
					else
						str = CSmartStr(str) & (((CSmartStr(idx) & " : ") & CSmartStr(val)) & vbcrlf)
					end if
				end if
				i_audit_loopIdx4=i_audit_loopIdx4+1
			loop
		end if
		strFields = ""
		if bValue(this_object.logValueEnable_p1(str_table)) then
			v = ""
			GetCollectionBounds newvalues,i_audit_loopIdx5,i_audit_loopMax5
			i_audit_exitLoop5=false
			do while i_audit_loopIdx5<=i_audit_loopMax5
				i_audit_exitLoop5=false
				do
					idx = GetCollectionKey(newvalues,i_audit_loopIdx5)
					doAssignment val,ArrayElement(newvalues,idx)
					doAssignmentByRef var_type,GetFieldType(idx,str_table)
					if bValue(IsBinaryType(var_type)) then
						exit do
					end if
					if bValue(IsDateFieldType(var_type)) then
						doAssignmentByRef newvalue,format_datetime_custom(db2time(ArrayElement(newvalues,idx)),"yyyy-MM-dd HH:mm:ss")
						doAssignmentByRef oldvalue,format_datetime_custom(db2time(ArrayElement(oldvalues,idx)),"yyyy-MM-dd HH:mm:ss")
					else
						doAssignment newvalue,ArrayElement(newvalues,idx)
						doAssignment oldvalue,ArrayElement(oldvalues,idx)
					end if
					if not IsEqual(newvalue,oldvalue) and not bValue(asp_array_key_exists(idx,keys)) then
						strFields = CSmartStr(strFields) & (CSmartStr(idx) & " [old]: ")
						if bValue(IsBinaryType(var_type)) then
							v = "<binary value>"
						else
							doAssignmentByRef v,asp_str_replace(CreateDictionary3(Empty,vbcrlf,Empty,vblf,Empty,"	")," ",oldvalue)
							if IsLess(300,asp_strlen(v)) then
								doAssignmentByRef v,asp_substr(v,0,300)
							end if
						end if
						strFields = CSmartStr(strFields) & (CSmartStr(v) & vbcrlf)
						strFields = CSmartStr(strFields) & (CSmartStr(idx) & " [new]: ")
						if bValue(IsBinaryType(var_type)) then
							v = "<binary value>"
						else
							doAssignmentByRef v,asp_str_replace(CreateDictionary3(Empty,vbcrlf,Empty,vblf,Empty,"	")," ",newvalue)
							if IsLess(300,asp_strlen(v)) then
								doAssignmentByRef v,asp_substr(v,0,300)
							end if
						end if
						strFields = CSmartStr(strFields) & (CSmartStr(v) & vbcrlf)
					end if
				loop while false
				if i_audit_exitLoop5 then _
					exit do
				i_audit_loopIdx5=i_audit_loopIdx5+1
			loop
			v = ""
		end if
		if not IsEqual(strFields,"") then
			str = CSmartStr(str) & ((CSmartStr(this_object.strFieldsHeader) & vbcrlf) & CSmartStr(strFields))
		end if
		doClassAssignment this_object.TableObj,"datetime",now()
		doClassAssignment this_object.TableObj,"ip",ArrayElement(this_object.params,0)
		doClassAssignment this_object.TableObj,"user",ArrayElement(this_object.params,1)
		doClassAssignment this_object.TableObj,"table",str_table
		doClassAssignment this_object.TableObj,"action",this_object.strEdit
		doClassAssignment this_object.TableObj,"description",str
		this_object.TableObj.Add 
	end if
	doAssignmentByRef method_AuditTrailTable_LogEdit,retval
	Exit Function
End Function
Function method_AuditTrailTable_LogDelete(byref this_object,ByVal str_table,ByVal values,ByVal keys)
	Dim retval,table,arr,str,idx,val,strFields,this,v
	retval = true
	doAssignment table,str_table
	Set arr = (CreateDictionary())
	if bValue(globalEvents.exists_p1("OnAuditLog")) then
		doAssignmentByRef retval,globalEvents.OnAuditLog_p6(this_object.strDelete,this_object.params,table,keys,values,arr)
	end if
	if bValue(retval) then
		str = ""
		if IsLess(0,asp_count(keys)) then
			str = CSmartStr(str) & (CSmartStr(this_object.strKeysHeader) & vbcrlf)
			GetCollectionBounds keys,i_audit_loopIdx6,i_audit_loopMax6
			do while i_audit_loopIdx6<=i_audit_loopMax6
				idx = GetCollectionKey(keys,i_audit_loopIdx6)
				doAssignment val,ArrayElement(keys,idx)
				str = CSmartStr(str) & (((CSmartStr(idx) & " : ") & CSmartStr(val)) & vbcrlf)
				i_audit_loopIdx6=i_audit_loopIdx6+1
			loop
		end if
		strFields = ""
		if bValue(this_object.logValueEnable_p1(str_table)) then
			v = ""
			GetCollectionBounds values,i_audit_loopIdx7,i_audit_loopMax7
			do while i_audit_loopIdx7<=i_audit_loopMax7
				idx = GetCollectionKey(values,i_audit_loopIdx7)
				doAssignment val,ArrayElement(values,idx)
				if not IsEqual(val,"") and not bValue(asp_array_key_exists(idx,keys)) then
					strFields = CSmartStr(strFields) & (CSmartStr(idx) & " [old]: ")
					if bValue(IsBinaryType(GetFieldType(idx,str_table))) then
						v = "<binary value>"
					else
						doAssignmentByRef v,asp_str_replace(CreateDictionary3(Empty,vbcrlf,Empty,vblf,Empty,"	")," ",val)
						if IsLess(300,asp_strlen(v)) then
							doAssignmentByRef v,asp_substr(v,0,300)
						end if
					end if
					strFields = CSmartStr(strFields) & (CSmartStr(v) & vbcrlf)
				end if
				i_audit_loopIdx7=i_audit_loopIdx7+1
			loop
		end if
		if not IsEqual(strFields,"") then
			str = CSmartStr(str) & ((CSmartStr(this_object.strFieldsHeader) & vbcrlf) & CSmartStr(strFields))
		end if
		doClassAssignment this_object.TableObj,"datetime",now()
		doClassAssignment this_object.TableObj,"ip",ArrayElement(this_object.params,0)
		doClassAssignment this_object.TableObj,"user",ArrayElement(this_object.params,1)
		doClassAssignment this_object.TableObj,"table",str_table
		doClassAssignment this_object.TableObj,"action",this_object.strDelete
		doClassAssignment this_object.TableObj,"description",str
		this_object.TableObj.Add 
	end if
	doAssignmentByRef method_AuditTrailTable_LogDelete,retval
	Exit Function
End Function
Function method_AuditTrailTable_LogAddEvent(byref this_object,ByVal message,ByVal description,ByVal stable)
	Dim retval,table,arr
	retval = true
	doAssignment table,stable
	Set arr = (CreateDictionary())
	if bValue(globalEvents.exists_p1("OnAuditLog")) then
		doAssignmentByRef retval,globalEvents.OnAuditLog_p6(message,this_object.params,table,keys,values,arr)
	end if
	if bValue(retval) then
		doClassAssignment this_object.TableObj,"datetime",now()
		doClassAssignment this_object.TableObj,"ip",ArrayElement(this_object.params,0)
		doClassAssignment this_object.TableObj,"user",ArrayElement(this_object.params,1)
		doClassAssignment this_object.TableObj,"table",stable
		doClassAssignment this_object.TableObj,"action",message
		doClassAssignment this_object.TableObj,"description",description
		this_object.TableObj.Add 
	end if
	doAssignmentByRef method_AuditTrailTable_LogAddEvent,retval
	Exit Function
End Function
Function method_AuditTrailTable_LoginSuccessful(byref this_object)
	Dim var_SERVER
	if IsLess(0,this_object.attLogin) and IsLess(0,this_object.timeLogin) then
		doClassAssignment this_object.TableObj,"ip",GetRequestValue(Request.ServerVariables,"REMOTE_ADDR")
		doClassAssignment this_object.TableObj,"action",this_object.strAccess
		this_object.TableObj.Delete 
	end if
End Function
Function method_AuditTrailTable_LoginUnsuccessful(byref this_object,ByVal pUsername)
	Dim var_SERVER
	if IsLess(0,this_object.attLogin) and IsLess(0,this_object.timeLogin) then
		doClassAssignment this_object.TableObj,"datetime",now()
		doClassAssignment this_object.TableObj,"ip",GetRequestValue(Request.ServerVariables,"REMOTE_ADDR")
		doClassAssignment this_object.TableObj,"user",pUsername
		this_object.TableObj.table = ""
		doClassAssignment this_object.TableObj,"action",this_object.strAccess
		this_object.TableObj.description = ""
		this_object.TableObj.Add 
	end if
End Function
Function method_AuditTrailTable_LoginAccess(byref this_object)
	Dim rstmp,i,data,firstAccess
	if IsLess(0,this_object.attLogin) and IsLess(0,this_object.timeLogin) then
		doAssignmentByRef rstmp,this_object.TableObj.Query_p2(((((CSmartStr(AddFieldWrappers("ip")) & "='") & CSmartStr(GetRequestValue(Request.ServerVariables,"REMOTE_ADDR"))) & "' and ") & CSmartStr(AddFieldWrappers("action"))) & "='access'",CSmartStr(AddFieldWrappers("id")) & " asc")
		i = 0
		do while bValue(doAssignmentByRef(data,db_fetch_array(rstmp)))
			if IsLessOrEqual(CSmartDbl(secondsPassedFrom(ArrayElement(data,"datetime")))/60,this_object.timeLogin) then
				if IsEqual(i,0) then
					doAssignment firstAccess,ArrayElement(data,"datetime")
				end if
				i = CSmartDbl(i)+1
			end if
		loop
		if IsLessOrEqual(this_object.attLogin,i) then
			doAssignmentByRef method_AuditTrailTable_LoginAccess,asp_ceil(CSmartDbl(this_object.timeLogin)-CSmartDbl(secondsPassedFrom(firstAccess))/60)
			Exit Function
		else
			method_AuditTrailTable_LoginAccess = false
			Exit Function
		end if
	else
		method_AuditTrailTable_LoginAccess = false
		Exit Function
	end if
End Function
Function method_AuditTrailTable_logValueEnable(byref this_object,ByVal table)
	if IsEqual(table,"payment") then
		method_AuditTrailTable_logValueEnable = false
		Exit Function
	end if
End Function

'------ Class AuditTrailFile ------
Class AuditTrailFile
	Public logfile
	Public strLogin
	Public strFailLogin
	Public strLogout
	Public strChPass
	Public strAdd
	Public strEdit
	Public strDelete
	Public strAccess
	Public strKeysHeader
	Public strFieldsHeader
	Public columnDate
	Public columnTime
	Public columnIP
	Public columnUser
	Public columnTable
	Public columnAction
	Public columnKey
	Public columnField
	Public columnOldValue
	Public columnNewValue
	Public params
	Public Function init_AuditTrailFile()
		DoAssignmentByRef init_AuditTrailFile,method_AuditTrailFile_AuditTrailFile(me)
	End Function
	Public Function LogLogin()
		DoAssignmentByRef LogLogin,method_AuditTrailFile_LogLogin(me)
	End Function
	Public Function LogLoginFailed_p1(ByVal pUsername)
		DoAssignmentByRef LogLoginFailed_p1,method_AuditTrailFile_LogLoginFailed(me,pUsername)
	End Function
	Public Function LogLogout()
		DoAssignmentByRef LogLogout,method_AuditTrailFile_LogLogout(me)
	End Function
	Public Function LogChPassword()
		DoAssignmentByRef LogChPassword,method_AuditTrailFile_LogChPassword(me)
	End Function
	Public Function LogAdd_p3(ByVal str_table,ByVal values,ByVal keys)
		DoAssignmentByRef LogAdd_p3,method_AuditTrailFile_LogAdd(me,str_table,values,keys)
	End Function
	Public Function LogEdit_p4(ByVal str_table,ByVal newvalues,ByVal oldvalues,ByVal keys)
		DoAssignmentByRef LogEdit_p4,method_AuditTrailFile_LogEdit(me,str_table,newvalues,oldvalues,keys)
	End Function
	Public Function LogDelete_p3(ByVal str_table,ByVal values,ByVal keys)
		DoAssignmentByRef LogDelete_p3,method_AuditTrailFile_LogDelete(me,str_table,values,keys)
	End Function
	Public Function CreateLogFile()
		DoAssignmentByRef CreateLogFile,method_AuditTrailFile_CreateLogFile(me)
	End Function
	Public Function LogAddEvent_p3(ByVal message,ByVal description,ByVal str_table)
		DoAssignmentByRef LogAddEvent_p3,method_AuditTrailFile_LogAddEvent(me,message,description,str_table)
	End Function
	Public Function LogAddEvent_p2(ByVal message,ByVal description)
		DoAssignmentByRef LogAddEvent_p2,method_AuditTrailFile_LogAddEvent(me,message,description,"")
	End Function
	Public Function LogAddEvent_p1(ByVal message)
		DoAssignmentByRef LogAddEvent_p1,method_AuditTrailFile_LogAddEvent(me,message,"","")
	End Function
	Public Function LoginAccess()
		DoAssignmentByRef LoginAccess,method_AuditTrailFile_LoginAccess(me)
	End Function
	Public Function LoginSuccessful()
		DoAssignmentByRef LoginSuccessful,method_AuditTrailFile_LoginSuccessful(me)
	End Function
	Public Function LoginUnsuccessful_p1(ByVal pUsername)
		DoAssignmentByRef LoginUnsuccessful_p1,method_AuditTrailFile_LoginUnsuccessful(me,pUsername)
	End Function
	Public Function logValueEnable_p1(ByVal table)
		DoAssignmentByRef logValueEnable_p1,method_AuditTrailFile_logValueEnable(me,table)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"logfile", logfile
		setArrElement out,"strLogin", strLogin
		setArrElement out,"strFailLogin", strFailLogin
		setArrElement out,"strLogout", strLogout
		setArrElement out,"strChPass", strChPass
		setArrElement out,"strAdd", strAdd
		setArrElement out,"strEdit", strEdit
		setArrElement out,"strDelete", strDelete
		setArrElement out,"strAccess", strAccess
		setArrElement out,"strKeysHeader", strKeysHeader
		setArrElement out,"strFieldsHeader", strFieldsHeader
		setArrElement out,"columnDate", columnDate
		setArrElement out,"columnTime", columnTime
		setArrElement out,"columnIP", columnIP
		setArrElement out,"columnUser", columnUser
		setArrElement out,"columnTable", columnTable
		setArrElement out,"columnAction", columnAction
		setArrElement out,"columnKey", columnKey
		setArrElement out,"columnField", columnField
		setArrElement out,"columnOldValue", columnOldValue
		setArrElement out,"columnNewValue", columnNewValue
		setArrElement out,"params", params
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment logfile, dict("logfile")
		DoAssignment strLogin, dict("strLogin")
		DoAssignment strFailLogin, dict("strFailLogin")
		DoAssignment strLogout, dict("strLogout")
		DoAssignment strChPass, dict("strChPass")
		DoAssignment strAdd, dict("strAdd")
		DoAssignment strEdit, dict("strEdit")
		DoAssignment strDelete, dict("strDelete")
		DoAssignment strAccess, dict("strAccess")
		DoAssignment strKeysHeader, dict("strKeysHeader")
		DoAssignment strFieldsHeader, dict("strFieldsHeader")
		DoAssignment columnDate, dict("columnDate")
		DoAssignment columnTime, dict("columnTime")
		DoAssignment columnIP, dict("columnIP")
		DoAssignment columnUser, dict("columnUser")
		DoAssignment columnTable, dict("columnTable")
		DoAssignment columnAction, dict("columnAction")
		DoAssignment columnKey, dict("columnKey")
		DoAssignment columnField, dict("columnField")
		DoAssignment columnOldValue, dict("columnOldValue")
		DoAssignment columnNewValue, dict("columnNewValue")
		DoAssignment params, dict("params")
	End Sub
' end serialize
End Class
' AuditTrailFile implementation 
Function method_AuditTrailFile_AuditTrailFile(byref this_object)
	this_object.logfile = "audit.log"
	this_object.strLogin = "login"
	this_object.strFailLogin = "failed login"
	this_object.strLogout = "logout"
	this_object.strChPass = "change password"
	this_object.strAdd = "add"
	this_object.strEdit = "edit"
	this_object.strDelete = "delete"
	this_object.strAccess = "access"
	this_object.strKeysHeader = "---Keys"
	this_object.strFieldsHeader = "---Fields"
	this_object.columnDate = "Date"
	this_object.columnTime = "Time"
	this_object.columnIP = "IP"
	this_object.columnUser = "User"
	this_object.columnTable = "Table"
	this_object.columnAction = "Action"
	this_object.columnKey = "Key field"
	this_object.columnField = "Field"
	this_object.columnOldValue = "Old value"
	this_object.columnNewValue = "New value"
	Dim userid,var_SESSION,var_SERVER
	userid = ""
	if bValue(Session("UserID")) then
		doAssignment userid,Session("UserID")
	end if
	doClassAssignment this_object,"params",CreateDictionary2(Empty,GetRequestValue(Request.ServerVariables,"REMOTE_ADDR"),Empty,userid)
End Function
Function method_AuditTrailFile_LogLogin(byref this_object)
	Dim retval,table,arr,fp,this,str
End Function
Function method_AuditTrailFile_LogLoginFailed(byref this_object,ByVal pUsername)
	Dim retval,table,arr,fp,this,str
End Function
Function method_AuditTrailFile_LogLogout(byref this_object)
	Dim var_SESSION,retval,table,arr,fp,this,str
End Function
Function method_AuditTrailFile_LogChPassword(byref this_object)
	Dim retval,table,arr,fp,this,str
End Function
Function method_AuditTrailFile_LogAdd(byref this_object,ByVal str_table,ByVal values,ByVal keys)
	Dim retval,table,arr,key,val,fp,this,str,str_add,idx,v
	retval = true
	doAssignment table,str_table
	Set arr = (CreateDictionary())
	if bValue(globalEvents.exists_p1("OnAuditLog")) then
		doAssignmentByRef retval,globalEvents.OnAuditLog_p6(this_object.strAdd,this_object.params,table,keys,values,arr)
	end if
	if bValue(retval) then
		if IsLess(0,asp_count(keys)) then
			key = ""
			GetCollectionBounds keys,i_audit_loopIdx9,i_audit_loopMax9
			do while i_audit_loopIdx9<=i_audit_loopMax9
				idx = GetCollectionKey(keys,i_audit_loopIdx9)
				doAssignment val,ArrayElement(keys,idx)
				if not IsEqual(key,"") then
					key = CSmartStr(key) & ","
				end if
				key = CSmartStr(key) & CSmartStr(val)
				i_audit_loopIdx9=i_audit_loopIdx9+1
			loop
		end if
		doAssignmentByRef fp,this_object.CreateLogFile()
		str = (((((((((((CSmartStr(format_datetime_custom(db2time(now()),"MMM dd,yyyy")) & CSmartStr(chr(9))) & CSmartStr(format_datetime_custom(db2time(now()),"HH:mm:ss"))) & CSmartStr(chr(9))) & CSmartStr(ArrayElement(this_object.params,0))) & CSmartStr(chr(9))) & CSmartStr(ArrayElement(this_object.params,1))) & CSmartStr(chr(9))) & CSmartStr(table)) & CSmartStr(chr(9))) & CSmartStr(this_object.strAdd)) & CSmartStr(chr(9))) & CSmartStr(key)
		str_add = ""
		if bValue(this_object.logValueEnable_p1(str_table)) then
			GetCollectionBounds values,i_audit_loopIdx10,i_audit_loopMax10
			do while i_audit_loopIdx10<=i_audit_loopMax10
				idx = GetCollectionKey(values,i_audit_loopIdx10)
				doAssignment val,ArrayElement(values,idx)
				if not IsEqual(val,"") and not bValue(asp_array_key_exists(idx,keys)) then
					v = ""
					if bValue(IsBinaryType(GetFieldType(idx,str_table))) then
						v = "<binary value>" & vbcrlf
					else
						doAssignmentByRef v,asp_str_replace(CreateDictionary3(Empty,vbcrlf,Empty,vblf,Empty,"	")," ",val)
						if IsLess(300,asp_strlen(v)) then
							doAssignmentByRef v,asp_substr(v,0,300)
						end if
					end if
					str_add = CSmartStr(str_add) & (((((CSmartStr(str) & CSmartStr(chr(9))) & CSmartStr(idx)) & CSmartStr(chr(9))) & CSmartStr(v)) & vbcrlf)
				end if
				i_audit_loopIdx10=i_audit_loopIdx10+1
			loop
		else
			str_add = CSmartStr(str_add) & (CSmartStr(str) & vbcrlf)
		end if
		if bValue(fp) then
			fputs fp,str_add
			asp_fclose fp
		end if
	end if
	doAssignmentByRef method_AuditTrailFile_LogAdd,retval
	Exit Function
End Function
Function method_AuditTrailFile_LogEdit(byref this_object,ByVal str_table,ByVal newvalues,ByVal oldvalues,ByVal keys)
	Dim retval,table,key,val,fp,this,str,putsValue,str_add,var_type,idx,newvalue,oldvalue,v1,v2
	retval = true
	doAssignment table,str_table
	if bValue(globalEvents.exists_p1("OnAuditLog")) then
		doAssignmentByRef retval,globalEvents.OnAuditLog_p6(this_object.strEdit,this_object.params,table,keys,newvalues,oldvalues)
	end if
	if bValue(retval) then
		if IsLess(0,asp_count(keys)) then
			key = ""
			GetCollectionBounds keys,i_audit_loopIdx11,i_audit_loopMax11
			do while i_audit_loopIdx11<=i_audit_loopMax11
				idx = GetCollectionKey(keys,i_audit_loopIdx11)
				doAssignment val,ArrayElement(keys,idx)
				if not IsEqual(key,"") then
					key = CSmartStr(key) & ","
				end if
				key = CSmartStr(key) & CSmartStr(val)
				i_audit_loopIdx11=i_audit_loopIdx11+1
			loop
		end if
		doAssignmentByRef fp,this_object.CreateLogFile()
		str = (((((((((((CSmartStr(format_datetime_custom(db2time(now()),"MMM dd,yyyy")) & CSmartStr(chr(9))) & CSmartStr(format_datetime_custom(db2time(now()),"HH:mm:ss"))) & CSmartStr(chr(9))) & CSmartStr(ArrayElement(this_object.params,0))) & CSmartStr(chr(9))) & CSmartStr(ArrayElement(this_object.params,1))) & CSmartStr(chr(9))) & CSmartStr(table)) & CSmartStr(chr(9))) & CSmartStr(this_object.strEdit)) & CSmartStr(chr(9))) & CSmartStr(key)
		putsValue = true
		str_add = ""
		if bValue(this_object.logValueEnable_p1(str_table)) then
			GetCollectionBounds newvalues,i_audit_loopIdx12,i_audit_loopMax12
			i_audit_exitLoop12=false
			do while i_audit_loopIdx12<=i_audit_loopMax12
				i_audit_exitLoop12=false
				do
					idx = GetCollectionKey(newvalues,i_audit_loopIdx12)
					doAssignment val,ArrayElement(newvalues,idx)
					doAssignmentByRef var_type,GetFieldType(idx,str_table)
					if bValue(IsBinaryType(var_type)) then
						exit do
					end if
					if bValue(IsDateFieldType(var_type)) then
						doAssignmentByRef newvalue,format_datetime_custom(db2time(ArrayElement(newvalues,idx)),"yyyy-MM-dd HH:mm:ss")
						doAssignmentByRef oldvalue,format_datetime_custom(db2time(ArrayElement(oldvalues,idx)),"yyyy-MM-dd HH:mm:ss")
					else
						doAssignment newvalue,ArrayElement(newvalues,idx)
						doAssignment oldvalue,ArrayElement(oldvalues,idx)
					end if
					if not IsEqual(newvalue,oldvalue) then
						v1 = ""
						if bValue(IsBinaryType(var_type)) then
							v1 = "<binary value>"
						else
							doAssignmentByRef v1,asp_str_replace(CreateDictionary3(Empty,vbcrlf,Empty,vblf,Empty,"	")," ",oldvalue)
							if IsLess(300,asp_strlen(v1)) then
								doAssignmentByRef v1,asp_substr(v1,0,300)
							end if
						end if
						v2 = ""
						if bValue(IsBinaryType(var_type)) then
							v2 = "<binary value>"
						else
							doAssignmentByRef v2,asp_str_replace(CreateDictionary3(Empty,vbcrlf,Empty,vblf,Empty,"	")," ",newvalue)
							if IsLess(300,asp_strlen(v2)) then
								doAssignmentByRef v2,asp_substr(v2,0,300)
							end if
						end if
						str_add = CSmartStr(str_add) & (((((((CSmartStr(str) & CSmartStr(chr(9))) & CSmartStr(idx)) & CSmartStr(chr(9))) & CSmartStr(v1)) & CSmartStr(chr(9))) & CSmartStr(v2)) & vbcrlf)
					end if
				loop while false
				if i_audit_exitLoop12 then _
					exit do
				i_audit_loopIdx12=i_audit_loopIdx12+1
			loop
		else
			str_add = CSmartStr(str_add) & (CSmartStr(str) & vbcrlf)
		end if
		if bValue(fp) then
			fputs fp,str_add
			asp_fclose fp
		end if
	end if
	doAssignmentByRef method_AuditTrailFile_LogEdit,retval
	Exit Function
End Function
Function method_AuditTrailFile_LogDelete(byref this_object,ByVal str_table,ByVal values,ByVal keys)
	Dim retval,table,arr,key,val,fp,this,str,str_add,v,idx
	retval = true
	doAssignment table,str_table
	Set arr = (CreateDictionary())
	if bValue(globalEvents.exists_p1("OnAuditLog")) then
		doAssignmentByRef retval,globalEvents.OnAuditLog_p6(this_object.strDelete,this_object.params,table,keys,values,arr)
	end if
	if bValue(retval) then
		if IsLess(0,asp_count(keys)) then
			key = ""
			GetCollectionBounds keys,i_audit_loopIdx13,i_audit_loopMax13
			do while i_audit_loopIdx13<=i_audit_loopMax13
				idx = GetCollectionKey(keys,i_audit_loopIdx13)
				doAssignment val,ArrayElement(keys,idx)
				if not IsEqual(key,"") then
					key = CSmartStr(key) & ","
				end if
				key = CSmartStr(key) & CSmartStr(val)
				i_audit_loopIdx13=i_audit_loopIdx13+1
			loop
		end if
		doAssignmentByRef fp,this_object.CreateLogFile()
		str = (((((((((((CSmartStr(format_datetime_custom(db2time(now()),"MMM dd,yyyy")) & CSmartStr(chr(9))) & CSmartStr(format_datetime_custom(db2time(now()),"HH:mm:ss"))) & CSmartStr(chr(9))) & CSmartStr(ArrayElement(this_object.params,0))) & CSmartStr(chr(9))) & CSmartStr(ArrayElement(this_object.params,1))) & CSmartStr(chr(9))) & CSmartStr(table)) & CSmartStr(chr(9))) & CSmartStr(this_object.strDelete)) & CSmartStr(chr(9))) & CSmartStr(key)
		str_add = ""
		if bValue(this_object.logValueEnable_p1(str_table)) then
			GetCollectionBounds values,i_audit_loopIdx14,i_audit_loopMax14
			do while i_audit_loopIdx14<=i_audit_loopMax14
				idx = GetCollectionKey(values,i_audit_loopIdx14)
				doAssignment val,ArrayElement(values,idx)
				v = ""
				if bValue(IsBinaryType(GetFieldType(idx,str_table))) then
					v = "<binary value>"
				else
					doAssignmentByRef v,asp_str_replace(CreateDictionary3(Empty,vbcrlf,Empty,vblf,Empty,"	")," ",val)
					if IsLess(300,asp_strlen(v)) then
						doAssignmentByRef v,asp_substr(v,0,300)
					end if
				end if
				if bValue(fp) then
					str_add = CSmartStr(str_add) & (((((CSmartStr(str) & CSmartStr(chr(9))) & CSmartStr(idx)) & CSmartStr(chr(9))) & CSmartStr(v)) & vbcrlf)
				end if
				i_audit_loopIdx14=i_audit_loopIdx14+1
			loop
		else
			str_add = CSmartStr(str) & vbcrlf
		end if
		if bValue(fp) then
			fputs fp,str_add
			asp_fclose fp
		end if
	end if
	doAssignmentByRef method_AuditTrailFile_LogDelete,retval
	Exit Function
End Function
Function method_AuditTrailFile_CreateLogFile(byref this_object)
	Dim p,logfileName,logfileExt,tn,fp,str
	doAssignmentByRef p,asp_strrpos(this_object.logfile,".",empty)
	doAssignmentByRef logfileName,asp_substr(this_object.logfile,0,p)
	doAssignmentByRef logfileExt,asp_substr(this_object.logfile,CSmartDbl(p)+1,empty)
	tn = (((CSmartStr(logfileName) & "_") & CSmartStr(format_datetime_custom(db2time(now()),"yyyyMMdd"))) & ".") & CSmartStr(logfileExt)
	doAssignmentByRef fp,asp_fopen(tn,"a")
	if bValue(fp) then
		if not bValue(filesize(tn)) then
			str = ((((((((((((((((((CSmartStr(this_object.columnDate) & CSmartStr(chr(9))) & CSmartStr(this_object.columnTime)) & CSmartStr(chr(9))) & CSmartStr(this_object.columnIP)) & CSmartStr(chr(9))) & CSmartStr(this_object.columnUser)) & CSmartStr(chr(9))) & CSmartStr(this_object.columnTable)) & CSmartStr(chr(9))) & CSmartStr(this_object.columnAction)) & CSmartStr(chr(9))) & CSmartStr(this_object.columnKey)) & CSmartStr(chr(9))) & CSmartStr(this_object.columnField)) & CSmartStr(chr(9))) & CSmartStr(this_object.columnOldValue)) & CSmartStr(chr(9))) & CSmartStr(this_object.columnNewValue)) & vbcrlf
			if bValue(fp) then
				fputs fp,str
			end if
		end if
	end if
	doAssignmentByRef method_AuditTrailFile_CreateLogFile,fp
	Exit Function
End Function
Function method_AuditTrailFile_LogAddEvent(byref this_object,ByVal message,ByVal description,ByVal str_table)
	Dim retval,table,arr,fp,this,str,params
	retval = true
	doAssignment table,str_table
	Set arr = (CreateDictionary())
	if bValue(globalEvents.exists_p1("OnAuditLog")) then
		doAssignmentByRef retval,globalEvents.OnAuditLog_p6(message,this_object.params,table,arr,arr,arr)
	end if
	if bValue(retval) then
		doAssignmentByRef fp,this_object.CreateLogFile()
		str = ((((((((((((CSmartStr(format_datetime_custom(db2time(now()),"MMM dd,yyyy")) & CSmartStr(chr(9))) & CSmartStr(format_datetime_custom(db2time(now()),"HH:mm:ss"))) & CSmartStr(chr(9))) & CSmartStr(ArrayElement(params,0))) & CSmartStr(chr(9))) & CSmartStr(ArrayElement(params,1))) & CSmartStr(chr(9))) & CSmartStr(table)) & CSmartStr(chr(9))) & CSmartStr(message)) & CSmartStr(chr(9))) & CSmartStr(description)) & vbcrlf
		if bValue(fp) then
			fputs fp,str
			asp_fclose fp
		end if
	end if
	doAssignmentByRef method_AuditTrailFile_LogAddEvent,retval
	Exit Function
End Function
Function method_AuditTrailFile_LoginAccess(byref this_object)
	method_AuditTrailFile_LoginAccess = false
	Exit Function
End Function
Function method_AuditTrailFile_LoginSuccessful(byref this_object)
	method_AuditTrailFile_LoginSuccessful = true
	Exit Function
End Function
Function method_AuditTrailFile_LoginUnsuccessful(byref this_object,ByVal pUsername)
	method_AuditTrailFile_LoginUnsuccessful = true
	Exit Function
End Function
Function method_AuditTrailFile_logValueEnable(byref this_object,ByVal table)
	if IsEqual(table,"payment") then
		method_AuditTrailFile_logValueEnable = false
		Exit Function
	end if
End Function
%>
