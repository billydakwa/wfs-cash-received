<%

'------ Class oLocking ------
Class oLocking
	Public lockTableName
	Public TableObj
	Public ConfirmTime
	Public UnlockTime
	Public ConfirmAdmin
	Public ConfirmUser
	Public LockAdmin
	Public LockUser
	Public UserID
	Public Function init_oLocking()
		DoAssignmentByRef init_oLocking,method_oLocking_oLocking(me)
	End Function
	Public Function LockRecord_p2(ByVal strtable,ByVal keys)
		DoAssignmentByRef LockRecord_p2,method_oLocking_LockRecord(me,strtable,keys)
	End Function
	Public Function UnlockRecord_p3(ByVal strtable,ByVal keys,ByVal sid)
		DoAssignmentByRef UnlockRecord_p3,method_oLocking_UnlockRecord(me,strtable,keys,sid)
	End Function
	Public Function ConfirmLock_p3(ByVal strtable,ByVal keys,ByRef message)
		DoAssignmentByRef ConfirmLock_p3,method_oLocking_ConfirmLock(me,strtable,keys,message)
	End Function
	Public Function GetLockInfo_p4(ByVal strtable,ByVal keys,ByVal links,ByVal pageid)
		DoAssignmentByRef GetLockInfo_p4,method_oLocking_GetLockInfo(me,strtable,keys,links,pageid)
	End Function
	Public Function UnlockAdmin_p3(ByVal strtable,ByVal keys,ByVal startEdit)
		DoAssignmentByRef UnlockAdmin_p3,method_oLocking_UnlockAdmin(me,strtable,keys,startEdit)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"lockTableName", lockTableName
		setArrElement out,"TableObj", TableObj
		setArrElement out,"ConfirmTime", ConfirmTime
		setArrElement out,"UnlockTime", UnlockTime
		setArrElement out,"ConfirmAdmin", ConfirmAdmin
		setArrElement out,"ConfirmUser", ConfirmUser
		setArrElement out,"LockAdmin", LockAdmin
		setArrElement out,"LockUser", LockUser
		setArrElement out,"UserID", UserID
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment lockTableName, dict("lockTableName")
		DoAssignment TableObj, dict("TableObj")
		DoAssignment ConfirmTime, dict("ConfirmTime")
		DoAssignment UnlockTime, dict("UnlockTime")
		DoAssignment ConfirmAdmin, dict("ConfirmAdmin")
		DoAssignment ConfirmUser, dict("ConfirmUser")
		DoAssignment LockAdmin, dict("LockAdmin")
		DoAssignment LockUser, dict("LockUser")
		DoAssignment UserID, dict("UserID")
	End Sub
' end serialize
End Class
' oLocking implementation 
Function method_oLocking_oLocking(byref this_object)
	this_object.lockTableName = ""
	this_object.ConfirmTime = 250
	this_object.UnlockTime = 300
	Dim var_SESSION
	this_object.ConfirmAdmin = "Administrator %s aborted your edit session"
	this_object.ConfirmUser = "Your edit session timed out"
	this_object.LockAdmin = "Record is edited by %s during %s minutes"
	this_object.LockUser = "Record is edited by another user"
	doClassAssignmentByRef this_object,"TableObj",dal.Table(this_object.lockTableName)
	if not IsEmpty(Session("UserID")) and not bValue(isnull(Session("UserID"))) then
		doClassAssignment this_object,"UserID",Session("UserID")
	else
		this_object.UserID = "Guest"
	end if
End Function
Function method_oLocking_LockRecord(byref this_object,ByVal strtable,ByVal keys)
	Dim skeys,val,sdate,arrDelete,rstmp,data
	skeys = ""
	GetCollectionBounds keys,i_locking_loopIdx2,i_locking_loopMax2
	do while i_locking_loopIdx2<=i_locking_loopMax2
		ind = GetCollectionKey(keys,i_locking_loopIdx2)
		doAssignment val,ArrayElement(keys,ind)
		if bValue(asp_strlen(skeys)) then
			skeys = CSmartStr(skeys) & "&"
		end if
		skeys = CSmartStr(skeys) & CSmartStr(asp_rawurlencode(val))
		i_locking_loopIdx2=i_locking_loopIdx2+1
	loop
	doAssignmentByRef sdate,now()
	doClassAssignment this_object.TableObj,"startdatetime",sdate
	doClassAssignment this_object.TableObj,"confirmdatetime",sdate
	doClassAssignment this_object.TableObj,"sessionid",session_id()
	doClassAssignment this_object.TableObj,"table",strtable
	doClassAssignment this_object.TableObj,"keys",skeys
	doClassAssignment this_object.TableObj,"userid",this_object.UserID
	this_object.TableObj.action = 1
	this_object.TableObj.Add 
	Set arrDelete = (CreateDictionary())
	doAssignmentByRef rstmp,this_object.TableObj.Query_p2(((((((((CSmartStr(AddFieldWrappers("table")) & "='") & CSmartStr(db_addslashes(strtable))) & "' and ") & CSmartStr(AddFieldWrappers("keys"))) & "='") & CSmartStr(db_addslashes(skeys))) & "' and ") & CSmartStr(AddFieldWrappers("action"))) & "=1",CSmartStr(AddFieldWrappers("id")) & " asc")
	do while bValue(doAssignmentByRef(data,db_fetch_array(rstmp)))
		if IsLess(this_object.UnlockTime,secondsPassedFrom(ArrayElement(data,"confirmdatetime"))) then
			setArrElement arrDelete,asp_count(arrDelete),ArrayElement(data,"id")
		else
			GetCollectionBounds arrDelete,i_locking_loopIdx4,i_locking_loopMax4
			do while i_locking_loopIdx4<=i_locking_loopMax4
				ind = GetCollectionKey(arrDelete,i_locking_loopIdx4)
				doAssignment val,ArrayElement(arrDelete,ind)
				doClassAssignment this_object.TableObj,"id",val
				this_object.TableObj.Delete 
				i_locking_loopIdx4=i_locking_loopIdx4+1
			loop
			if IsEqual(ArrayElement(data,"sessionid"),session_id()) then
				method_oLocking_LockRecord = true
				Exit Function
			else
				doClassAssignment this_object.TableObj,"sessionid",session_id()
				this_object.TableObj.action = 1
				doClassAssignment this_object.TableObj,"table",strtable
				doClassAssignment this_object.TableObj,"keys",skeys
				this_object.TableObj.Delete 
				method_oLocking_LockRecord = false
				Exit Function
			end if
		end if
	loop
	method_oLocking_LockRecord = false
	Exit Function
End Function
Function method_oLocking_UnlockRecord(byref this_object,ByVal strtable,ByVal keys,ByVal sid)
	Dim skeys,val
	if IsEqual(sid,"") then
		doAssignmentByRef sid,session_id()
	end if
	skeys = ""
	GetCollectionBounds keys,i_locking_loopIdx5,i_locking_loopMax5
	do while i_locking_loopIdx5<=i_locking_loopMax5
		ind = GetCollectionKey(keys,i_locking_loopIdx5)
		doAssignment val,ArrayElement(keys,ind)
		if bValue(asp_strlen(skeys)) then
			skeys = CSmartStr(skeys) & "&"
		end if
		skeys = CSmartStr(skeys) & CSmartStr(asp_rawurlencode(val))
		i_locking_loopIdx5=i_locking_loopIdx5+1
	loop
	doClassAssignment this_object.TableObj,"table",strtable
	doClassAssignment this_object.TableObj,"keys",skeys
	doClassAssignment this_object.TableObj,"sessionid",sid
	this_object.TableObj.action = 1
	this_object.TableObj.Delete 
End Function
Function method_oLocking_ConfirmLock(byref this_object,ByVal strtable,ByVal keys,ByRef message)
	Dim skeys,val,sdate,rstmp,myfound,newid,oldid,newdate,olddate,otherfound,tempfound,data,this
	skeys = ""
	GetCollectionBounds keys,i_locking_loopIdx6,i_locking_loopMax6
	do while i_locking_loopIdx6<=i_locking_loopMax6
		ind = GetCollectionKey(keys,i_locking_loopIdx6)
		doAssignment val,ArrayElement(keys,ind)
		if bValue(asp_strlen(skeys)) then
			skeys = CSmartStr(skeys) & "&"
		end if
		skeys = CSmartStr(skeys) & CSmartStr(asp_rawurlencode(val))
		i_locking_loopIdx6=i_locking_loopIdx6+1
	loop
	doAssignmentByRef sdate,now()
	doClassAssignment this_object.TableObj,"startdatetime",sdate
	doClassAssignment this_object.TableObj,"confirmdatetime",sdate
	doClassAssignment this_object.TableObj,"sessionid",session_id()
	doClassAssignment this_object.TableObj,"table",strtable
	doClassAssignment this_object.TableObj,"keys",skeys
	doClassAssignment this_object.TableObj,"userid",this_object.UserID
	this_object.TableObj.action = 1
	this_object.TableObj.Add 
	doAssignmentByRef rstmp,this_object.TableObj.Query_p2(((((((((CSmartStr(AddFieldWrappers("table")) & "='") & CSmartStr(db_addslashes(strtable))) & "' and ") & CSmartStr(AddFieldWrappers("keys"))) & "='") & CSmartStr(db_addslashes(skeys))) & "' and ") & CSmartStr(AddFieldWrappers("action"))) & "=1",CSmartStr(AddFieldWrappers("id")) & " asc")
	myfound = 0
	newid = 0
	oldid = 0
	newdate = ""
	olddate = ""
	otherfound = 0
	tempfound = 0
	i_locking_exitLoop7=false
	do while bValue(doAssignmentByRef(data,db_fetch_array(rstmp)))
		i_locking_exitLoop7=false
		do
			if IsEqual(ArrayElement(data,"sessionid"),session_id()) then
				doAssignment oldid,newid
				doAssignment newid,ArrayElement(data,"id")
				doAssignment newdate,ArrayElement(data,"confirmdatetime")
				doAssignment olddate,newdate
				myfound = CSmartDbl(myfound)+1
				doAssignment otherfound,tempfound
				tempfound = 0
				exit do
			end if
			tempfound = CSmartDbl(tempfound)+1
		loop while false
		if i_locking_exitLoop7 then _
			exit do
	loop
	if IsLess(1,myfound) and not bValue(otherfound) then
		doClassAssignment this_object.TableObj,"id",oldid
		doClassAssignment this_object.TableObj,"confirmdatetime",now()
		this_object.TableObj.Update 
		doClassAssignment this_object.TableObj,"id",newid
		this_object.TableObj.Delete 
		method_oLocking_ConfirmLock = true
		Exit Function
	else
		if IsLess(1,myfound) and bValue(otherfound) then
			if IsLess(CSmartDbl(this_object.UnlockTime)-5,secondsPassedFrom(olddate)) then
				this_object.UnlockRecord_p3 strtable,keys,session_id()
				doAssignment message,this_object.ConfirmUser
				method_oLocking_ConfirmLock = false
				Exit Function
			else
				doClassAssignment this_object.TableObj,"id",oldid
				doClassAssignment this_object.TableObj,"confirmdatetime",now()
				this_object.TableObj.Update 
				doClassAssignment this_object.TableObj,"id",newid
				this_object.TableObj.Delete 
				method_oLocking_ConfirmLock = true
				Exit Function
			end if
		else
			this_object.UnlockRecord_p3 strtable,keys,session_id()
			doAssignmentByRef rstmp,this_object.TableObj.Query_p2(((((((((((((CSmartStr(AddFieldWrappers("table")) & "='") & CSmartStr(db_addslashes(strtable))) & "' and ") & CSmartStr(AddFieldWrappers("keys"))) & "='") & CSmartStr(db_addslashes(skeys))) & "' and ") & CSmartStr(AddFieldWrappers("sessionid"))) & "<>'") & CSmartStr(session_id())) & "' and ") & CSmartStr(AddFieldWrappers("action"))) & "=2",CSmartStr(AddFieldWrappers("id")) & " asc")
			if bValue(doAssignmentByRef(data,db_fetch_array(rstmp))) then
				doAssignmentByRef message,mysprintf(this_object.ConfirmAdmin,CreateDictionary1(Empty,ArrayElement(data,"userid")))
			else
				doAssignment message,this_object.ConfirmUser
			end if
			doClassAssignment this_object.TableObj,"table",strtable
			doClassAssignment this_object.TableObj,"keys",skeys
			this_object.TableObj.action = 2
			this_object.TableObj.Delete 
			method_oLocking_ConfirmLock = false
			Exit Function
		end if
	end if
End Function
Function method_oLocking_GetLockInfo(byref this_object,ByVal strtable,ByVal keys,ByVal links,ByVal pageid)
	Dim page,skeys,val,rstmp,data,sdate,arrDateTime,str
	page = CSmartStr(GetTableURL(strtable)) & "_edit.asp"
	skeys = ""
	GetCollectionBounds keys,i_locking_loopIdx8,i_locking_loopMax8
	do while i_locking_loopIdx8<=i_locking_loopMax8
		ind = GetCollectionKey(keys,i_locking_loopIdx8)
		doAssignment val,ArrayElement(keys,ind)
		if bValue(asp_strlen(skeys)) then
			skeys = CSmartStr(skeys) & "&"
		end if
		skeys = CSmartStr(skeys) & CSmartStr(asp_rawurlencode(val))
		i_locking_loopIdx8=i_locking_loopIdx8+1
	loop
	doAssignmentByRef rstmp,this_object.TableObj.Query_p2(((((((((((((CSmartStr(AddFieldWrappers("table")) & "='") & CSmartStr(db_addslashes(strtable))) & "' and ") & CSmartStr(AddFieldWrappers("keys"))) & "='") & CSmartStr(db_addslashes(skeys))) & "' and ") & CSmartStr(AddFieldWrappers("sessionid"))) & "<>'") & CSmartStr(session_id())) & "' and ") & CSmartStr(AddFieldWrappers("action"))) & "=1",CSmartStr(AddFieldWrappers("id")) & " asc")
	if bValue(doAssignmentByRef(data,db_fetch_array(rstmp))) then
		doAssignmentByRef sdate,now()
		doAssignmentByRef arrDateTime,db2time(ArrayElement(data,"startdatetime"))
		doAssignmentByRef str,mysprintf(this_object.LockAdmin,CreateDictionary2(Empty,ArrayElement(data,"userid"),Empty,round(CSmartDbl(secondsPassedFrom(ArrayElement(data,"startdatetime")))/60,2)))
		if bValue(links) then
			str = CSmartStr(str) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			str = CSmartStr(str) & (((((((((((("<a class=admin_links href='#' onclick=""UnlockAdmin('" & CSmartStr(page)) & "','") & CSmartStr(htmlspecialchars(jsreplace(strtable)))) & "','") & CSmartStr(htmlspecialchars(jsreplace(skeys)))) & "','") & CSmartStr(ArrayElement(data,"sessionid"))) & "','no',") & CSmartStr(pageid)) & ");return false;"">") & CSmartStr("Unlock record")) & "</a>")
			str = CSmartStr(str) & (((((((((((("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a class=admin_links href='#' onclick=""UnlockAdmin('" & CSmartStr(page)) & "','") & CSmartStr(htmlspecialchars(jsreplace(strtable)))) & "','") & CSmartStr(htmlspecialchars(jsreplace(skeys)))) & "','") & CSmartStr(ArrayElement(data,"sessionid"))) & "','yes',") & CSmartStr(pageid)) & ");return false;"">") & CSmartStr("Edit record")) & "</a>")
		end if
		doAssignmentByRef method_oLocking_GetLockInfo,str
		Exit Function
	else
		method_oLocking_GetLockInfo = ""
		Exit Function
	end if
End Function
Function method_oLocking_UnlockAdmin(byref this_object,ByVal strtable,ByVal keys,ByVal startEdit)
	Dim skeys,val,sdate,rstmp
	skeys = ""
	GetCollectionBounds keys,i_locking_loopIdx9,i_locking_loopMax9
	do while i_locking_loopIdx9<=i_locking_loopMax9
		ind = GetCollectionKey(keys,i_locking_loopIdx9)
		doAssignment val,ArrayElement(keys,ind)
		if bValue(asp_strlen(skeys)) then
			skeys = CSmartStr(skeys) & "&"
		end if
		skeys = CSmartStr(skeys) & CSmartStr(asp_rawurlencode(val))
		i_locking_loopIdx9=i_locking_loopIdx9+1
	loop
	doAssignmentByRef sdate,now()
	if bValue(startEdit) then
		doClassAssignment this_object.TableObj,"startdatetime",sdate
		doClassAssignment this_object.TableObj,"confirmdatetime",sdate
		doClassAssignment this_object.TableObj,"sessionid",session_id()
		doClassAssignment this_object.TableObj,"table",strtable
		doClassAssignment this_object.TableObj,"keys",skeys
		doClassAssignment this_object.TableObj,"userid",this_object.UserID
		this_object.TableObj.action = 1
		this_object.TableObj.Add 
	end if
	doAssignmentByRef rstmp,CustomQuery(((((((((((((((("delete from " & CSmartStr(AddTableWrappers(this_object.lockTableName))) & " where ") & CSmartStr(AddFieldWrappers("table"))) & "='") & CSmartStr(db_addslashes(strtable))) & "' and ") & CSmartStr(AddFieldWrappers("keys"))) & "='") & CSmartStr(db_addslashes(skeys))) & "' and ") & CSmartStr(AddFieldWrappers("action"))) & "=1 and ") & CSmartStr(AddFieldWrappers("sessionid"))) & "<>'") & CSmartStr(session_id())) & "' ")
	doAssignmentByRef rstmp,CustomQuery(((((((("delete from " & CSmartStr(AddTableWrappers(this_object.lockTableName))) & " where ") & CSmartStr(AddFieldWrappers("startdatetime"))) & "<'") & CSmartStr(format_datetime_custom(adddays(db2time(now()),-2),"yyyy-MM-dd HH:mm:ss"))) & "' and ") & CSmartStr(AddFieldWrappers("action"))) & "=2")
	doClassAssignment this_object.TableObj,"startdatetime",sdate
	doClassAssignment this_object.TableObj,"confirmdatetime",sdate
	doClassAssignment this_object.TableObj,"sessionid",session_id()
	doClassAssignment this_object.TableObj,"table",strtable
	doClassAssignment this_object.TableObj,"keys",skeys
	doClassAssignment this_object.TableObj,"userid",this_object.UserID
	this_object.TableObj.action = 2
	this_object.TableObj.Add 
End Function
%>
