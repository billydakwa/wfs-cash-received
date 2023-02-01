<%
Function db_addslashes(ByVal str)
	doAssignmentByRef db_addslashes,asp_str_replace("'","''",str)
	Exit Function
End Function
Function db_stripslashes(ByVal str)
	doAssignmentByRef db_stripslashes,asp_str_replace("''","'",str)
	Exit Function
End Function
Function db_addslashesbinary(ByVal str)
	db_addslashesbinary = "0x" & CSmartStr(bin2hex(str))
	Exit Function
End Function
Function db_stripslashesbinary(ByVal str)
	Dim pos,pos1
	doAssignmentByRef pos,asp_strpos(str,".Picture",empty)
	if IsFalse(pos) or IsLess(300,pos) then
		doAssignmentByRef db_stripslashesbinary,str
		Exit Function
	end if
	doAssignmentByRef pos1,asp_strpos(str,"BM",pos)
	if IsFalse(pos1) or IsLess(300,pos1) then
		doAssignmentByRef db_stripslashesbinary,str
		Exit Function
	end if
	doAssignmentByRef db_stripslashesbinary,asp_substr(str,pos1,empty)
	Exit Function
End Function
Function AddFieldWrappers(ByVal strName)
	if IsEqual(asp_substr(strName,0,1),strLeftWrapper) then
		doAssignmentByRef AddFieldWrappers,strName
		Exit Function
	end if
	AddFieldWrappers = (CSmartStr(strLeftWrapper) & CSmartStr(strName)) & CSmartStr(strRightWrapper)
	Exit Function
End Function
Function AddTableWrappers(ByVal strName)
	Dim arr,ret
	if IsEqual(asp_substr(strName,0,1),strLeftWrapper) then
		doAssignmentByRef AddTableWrappers,strName
		Exit Function
	end if
	doAssignmentByRef arr,explode(".",strName)
	ret = (CSmartStr(strLeftWrapper) & CSmartStr(ArrayElement(arr,0))) & CSmartStr(strRightWrapper)
	if IsLess(1,asp_count(arr)) then
		ret = CSmartStr(ret) & ((("." & CSmartStr(strLeftWrapper)) & CSmartStr(ArrayElement(arr,1))) & CSmartStr(strRightWrapper))
	end if
	doAssignmentByRef AddTableWrappers,ret
	Exit Function
End Function
Function RemoveFieldWrappers(ByVal strName)
	if IsEqual(asp_substr(strName,0,1),strLeftWrapper) then
		doAssignmentByRef RemoveFieldWrappers,asp_substr(strName,1,CSmartDbl(asp_strlen(strName))-2)
		Exit Function
	end if
	doAssignmentByRef RemoveFieldWrappers,strName
	Exit Function
End Function
Function RemoveTableWrappers(ByVal strName)
	Dim arr,ret
	if not IsEqual(asp_substr(strName,0,1),strLeftWrapper) then
		doAssignmentByRef RemoveTableWrappers,strName
		Exit Function
	end if
	doAssignmentByRef arr,explode(".",strName)
	doAssignmentByRef ret,asp_substr(ArrayElement(arr,0),1,CSmartDbl(asp_strlen(ArrayElement(arr,0)))-2)
	if IsLess(1,asp_count(arr)) then
		ret = CSmartStr(ret) & ("." & CSmartStr(asp_substr(ArrayElement(arr,1),1,CSmartDbl(asp_strlen(ArrayElement(arr,1)))-2)))
	end if
	doAssignmentByRef RemoveTableWrappers,ret
	Exit Function
End Function
Function db_upper(ByVal dbval)
	db_upper = ("upper(" & CSmartStr(dbval)) & ")"
	Exit Function
End Function
Function db_datequotes(ByVal val)
	db_datequotes = ("convert(datetime,'" & CSmartStr(val)) & "',120)"
	Exit Function
End Function
Function db_field2char(ByVal value,ByVal var_type)
	if bValue(IsCharType(var_type)) then
		doAssignmentByRef db_field2char,value
		Exit Function
	end if
	if not bValue(IsDateFieldType(var_type)) then
		db_field2char = ("convert(varchar," & CSmartStr(value)) & ")"
		Exit Function
	end if
	db_field2char = ("convert(varchar(50)," & CSmartStr(value)) & ", 120)"
	Exit Function
End Function
Function db_field2time(ByVal value,ByVal var_type)
	doAssignmentByRef db_field2time,value
	Exit Function
End Function
%>
