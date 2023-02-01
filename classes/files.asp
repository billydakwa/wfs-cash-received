<%

'------ Class MoveFile ------
Class MoveFile
	Public sourceFilename
	Public destFilename
	Public destPath
	Public destPathIsAbsolute
	Public Function init_MoveFile_p4(ByVal source,ByVal name,ByVal path,ByVal abs)
		DoAssignmentByRef init_MoveFile_p4,method_MoveFile_MoveFile(me,source,name,path,abs)
	End Function
	Public Function Move()
		DoAssignmentByRef Move,method_MoveFile_Move(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"sourceFilename", sourceFilename
		setArrElement out,"destFilename", destFilename
		setArrElement out,"destPath", destPath
		setArrElement out,"destPathIsAbsolute", destPathIsAbsolute
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment sourceFilename, dict("sourceFilename")
		DoAssignment destFilename, dict("destFilename")
		DoAssignment destPath, dict("destPath")
		DoAssignment destPathIsAbsolute, dict("destPathIsAbsolute")
	End Sub
' end serialize
End Class
' MoveFile implementation 
Function method_MoveFile_MoveFile(byref this_object,ByVal source,ByVal name,ByVal path,ByVal abs)
	doClassAssignment this_object,"sourceFilename",source
	doClassAssignment this_object,"destFilename",name
	doClassAssignment this_object,"destPath",path
	doClassAssignment this_object,"destPathIsAbsolute",abs
End Function
Function method_MoveFile_Move(byref this_object)
	Dim path,last
	doAssignment path,this_object.destPath
	if not bValue(this_object.destPathIsAbsolute) then
		doAssignmentByRef path,getabspath(path)
	end if
	doAssignmentByRef last,asp_substr(path,CSmartDbl(asp_strlen(path))-1,empty)
	if not IsEqual(last,"/") and not IsEqual(last,"\") then
		path = CSmartStr(path) & "/"
	end if
	runner_move_uploaded_file this_object.sourceFilename,CSmartStr(path) & CSmartStr(this_object.destFilename)
End Function

'------ Class SaveFile ------
Class SaveFile
	Public fileContents
	Public destFilename
	Public destPath
	Public destPathIsAbsolute
	Public Function init_SaveFile_p4(ByVal contents,ByVal name,ByVal path,ByVal abs)
		DoAssignmentByRef init_SaveFile_p4,method_SaveFile_SaveFile(me,contents,name,path,abs)
	End Function
	Public Function Save()
		DoAssignmentByRef Save,method_SaveFile_Save(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"fileContents", fileContents
		setArrElement out,"destFilename", destFilename
		setArrElement out,"destPath", destPath
		setArrElement out,"destPathIsAbsolute", destPathIsAbsolute
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment fileContents, dict("fileContents")
		DoAssignment destFilename, dict("destFilename")
		DoAssignment destPath, dict("destPath")
		DoAssignment destPathIsAbsolute, dict("destPathIsAbsolute")
	End Sub
' end serialize
End Class
' SaveFile implementation 
Function method_SaveFile_SaveFile(byref this_object,ByVal contents,ByVal name,ByVal path,ByVal abs)
	doClassAssignment this_object,"fileContents",contents
	doClassAssignment this_object,"destFilename",name
	doClassAssignment this_object,"destPath",path
	doClassAssignment this_object,"destPathIsAbsolute",abs
End Function
Function method_SaveFile_Save(byref this_object)
	Dim path,last
	doAssignment path,this_object.destPath
	if not bValue(this_object.destPathIsAbsolute) then
		doAssignmentByRef path,getabspath(path)
	end if
	doAssignmentByRef last,asp_substr(path,CSmartDbl(asp_strlen(path))-1,empty)
	if not IsEqual(last,"/") and not IsEqual(last,"\") then
		path = CSmartStr(path) & "/"
	end if
	runner_save_file CSmartStr(path) & CSmartStr(this_object.destFilename),this_object.fileContents
End Function

'------ Class DeleteFile ------
Class DeleteFile
	Public destFilename
	Public destPath
	Public destPathIsAbsolute
	Public Function init_DeleteFile_p3(ByVal name,ByVal path,ByVal abs)
		DoAssignmentByRef init_DeleteFile_p3,method_DeleteFile_DeleteFile(me,name,path,abs)
	End Function
	Public Function Delete()
		DoAssignmentByRef Delete,method_DeleteFile_Delete(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"destFilename", destFilename
		setArrElement out,"destPath", destPath
		setArrElement out,"destPathIsAbsolute", destPathIsAbsolute
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment destFilename, dict("destFilename")
		DoAssignment destPath, dict("destPath")
		DoAssignment destPathIsAbsolute, dict("destPathIsAbsolute")
	End Sub
' end serialize
End Class
' DeleteFile implementation 
Function method_DeleteFile_DeleteFile(byref this_object,ByVal name,ByVal path,ByVal abs)
	doClassAssignment this_object,"destFilename",name
	doClassAssignment this_object,"destPath",path
	doClassAssignment this_object,"destPathIsAbsolute",abs
End Function
Function method_DeleteFile_Delete(byref this_object)
	Dim path,last
	doAssignment path,this_object.destPath
	if not bValue(this_object.destPathIsAbsolute) then
		doAssignmentByRef path,getabspath(path)
	end if
	doAssignmentByRef last,asp_substr(path,CSmartDbl(asp_strlen(path))-1,empty)
	if not IsEqual(last,"/") and not IsEqual(last,"\") then
		path = CSmartStr(path) & "/"
	end if
	runner_delete_file CSmartStr(path) & CSmartStr(this_object.destFilename)
End Function
%>
