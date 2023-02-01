<%
set dal_info=server.CreateObject("Scripting.Dictionary")

class tDAL
	public acc_awb
	public payment
	public part
	public awb
    public lstTables
    Public sub class_initialize()
        set lstTables=server.CreateObject("Scripting.Dictionary")
    end sub
    public sub FillTablesList()
		if asp_count(lstTables)>0 then
			exit sub
		end if
      set tmpTables=server.CreateObject("Scripting.Dictionary")
      tmpTables.Add "name","acc_awb"
      tmpTables.Add "varname","acc_awb"
	  lstTables.Add asp_count(lstTables),tmpTables
      set tmpTables=server.CreateObject("Scripting.Dictionary")
      tmpTables.Add "name","payment"
      tmpTables.Add "varname","payment"
	  lstTables.Add asp_count(lstTables),tmpTables
      set tmpTables=server.CreateObject("Scripting.Dictionary")
      tmpTables.Add "name","part"
      tmpTables.Add "varname","part"
	  lstTables.Add asp_count(lstTables),tmpTables
      set tmpTables=server.CreateObject("Scripting.Dictionary")
      tmpTables.Add "name","awb"
      tmpTables.Add "varname","awb"
	  lstTables.Add asp_count(lstTables),tmpTables
	end sub
	Function Table(strTable)
	    call FillTablesList()
        for each tbl in lstTables.Keys
            if ucase(strTable)=ucase(lstTables(tbl)("name")) then
                call CreateClass(lstTables(tbl))
                Execute "set Table=" & lstTables(tbl)("varname")
                exit function
            end if
        next
'	check table names without  and other prefixes
         for each tbl in lstTables.Keys
            if ucase(cutprefix(strTable))=ucase(cutprefix(lstTables(tbl)("name"))) then
                call CreateClass(lstTables(tbl))
                Execute "set Table=" & lstTables(tbl)("varname")
                exit function
            end if
        next
	End Function
	function GetTablesList
		call FillTablesList()
		set res=server.CreateObject("Scripting.Dictionary")
		for each key in lstTables.keys
			res(asp_count(res))=lstTables(key)("name")
		next
		set GetTablesList=res
	end function
	
	function GetFieldsList(stable)
		dim tbl
		set tbl = Table(stable)
		set GetFieldsList=tbl.GetFieldsList()
	end function
	
	function GetFieldType(stable,field)
		dim tbl
		set tbl = Table(stable)
		GetFieldType=tbl.GetFieldType(field)
	end function
	
	function GetDBTableKeys(stable)
		dim tbl
		set tbl = Table(stable)
		set GetDBTableKeys=tbl.GetDBTableKeys()
	end function
	
	function CreateClass(tbl)
        dim res,field
	    execute("res=vartype(" & tbl("varname") & ")")
		if res=vbObject then
			exit function
		end if
        asp_include "include/dal/" & GoodFieldName(tbl("name")) & ".asp",true
'//	create class and object

		command="Class class_" & GoodFieldName(tbl("name")) & vbcrlf &_
		"    Public m_TableName"& vbcrlf &_
        "	Public m_Param"& vbcrlf &_
        "	Public m_Value"& vbcrlf &_
        "	Public m_ChangedFields" & vbcrlf
        arrFields=dal_info(tbl("name")).keys()
	    for each field in arrFields
		    command=command & " Public m_fld" & GoodFieldName(field) & vbcrlf
	    next
	    command=command & "	Public sub class_initialize()"& vbcrlf &_
    "		set m_Param=CreateObject(""Scripting.Dictionary"")"& vbcrlf &_
    "		set m_Value=CreateObject(""Scripting.Dictionary"")"& vbcrlf &_
    "		set m_ChangedFields=CreateObject(""Scripting.Dictionary"")"& vbcrlf &_
    "		m_TableName = """&replace(tbl("name"),"""","""""")&""""& vbcrlf &_
    "	end sub"& vbcrlf &_
    "    function GetFieldsList()"& vbcrlf &_
    "    	set GetFieldsList=asp_array_keys(dal_info(m_TableName),empty)"& vbcrlf &_
    "    end function"& vbcrlf &_
    "    function GetFieldType(field)"& vbcrlf &_
    "    	if not asp_array_key_exists(field,dal_info(m_TableName)) then"& vbcrlf &_
    "    		GetFieldType=200"& vbcrlf &_
    "    		exit function"& vbcrlf &_
    "    	end if"& vbcrlf &_
    "    	GetFieldType=dal_info(m_TableName)(field)(""nType"")"& vbcrlf &_
    "    end function"& vbcrlf &_
    "    function GetDBTableKeys()"& vbcrlf &_
	"    dim fname, ret"& vbcrlf &_
    "    	set ret=server.CreateObject(""Scripting.Dictionary"")"& vbcrlf &_
    "    	if not asp_array_key_exists(m_TableName,dal_info) or not asp_is_array(dal_info(m_TableName)) then"& vbcrlf &_
	"    		set GetDBTableKeys=ret"& vbcrlf &_
	"    		exit function"& vbcrlf &_
	"    	end if"& vbcrlf &_
	"  		for each fname in dal_info(m_TableName).Keys"& vbcrlf &_
	"			if bValue(dal_info(m_TableName)(fname)(""key"")) then"& vbcrlf &_
	"				ret(asp_count(ret))=fname"& vbcrlf &_
	"			end if"& vbcrlf &_
	"		next"& vbcrlf &_
	"		set GetDBTableKeys=ret"& vbcrlf &_
	"	end function"& vbcrlf &_
    "    Public Function TableName()"& vbcrlf &_
    "        TableName=AddTableWrappers(m_TableName)"& vbcrlf &_
    "    End Function"& vbcrlf &_
    "   "& vbcrlf &_
	"    function PrepareValue(value,stype)"& vbcrlf &_
	"    	if IsDateFieldType(stype) then "& vbcrlf &_
	"    		if vartype(value)=vbDate then _"& vbcrlf &_
	"    			value=localdatetime2db(value,"""")"& vbcrlf &_
	"    		PrepareValue=db_datequotes(value)"& vbcrlf &_
	"    	elseif NeedQuotes(stype) then"& vbcrlf &_
	"    		PrepareValue=""'"" & db_addslashes(value) & ""'"" "& vbcrlf &_
	"    	else"& vbcrlf &_
	"    		PrepareValue=value"& vbcrlf &_
	"    	end if"& vbcrlf &_
	"    end function"& vbcrlf &_	
    "    Public Sub Add()"& vbcrlf &_
    "		DALAdd dal."& tbl("varname") & vbcrlf &_
    "	end sub"& vbcrlf &_
    "    Public Sub Delete()"& vbcrlf &_
    "		DALDelete dal."& tbl("varname") & vbcrlf &_
    "	end sub"& vbcrlf &_
    "    Public Sub Update()"& vbcrlf &_
    "		DALUpdate dal."& tbl("varname") & vbcrlf &_
    "	end sub"& vbcrlf &_
    "    Public Function QueryAll()"& vbcrlf &_
    "		set QueryAll = DALQueryAll(dal."& tbl("varname") &")"& vbcrlf &_
    "	end function"& vbcrlf &_
    "    Public Function Query(sw,ob)"& vbcrlf &_
    "		set Query = DALQuery(dal."& tbl("varname") &",sw,ob)"& vbcrlf &_
    "	end function"& vbcrlf &_
    "    Public Function Query_p2(sw,ob)"& vbcrlf &_
    "		set Query_p2 = Query(sw,ob)"& vbcrlf &_
    "	end function"& vbcrlf &_
    "    Public Function FetchByID()"& vbcrlf &_
    "		set FetchByID=DALFetchByID(dal."& tbl("varname") &")"& vbcrlf &_
    "	end function"& vbcrlf &_
    "    Public Sub Reset()"& vbcrlf &_
    "		DALReset dal."& tbl("varname") &""& vbcrlf &_
    "	end sub"& vbcrlf &_
    "    Public Property Let Param(p_Fields,p_Data)"& vbcrlf &_
    "		DALLetParam dal."& tbl("varname") &",p_Fields,p_Data"& vbcrlf &_
    "	end Property"& vbcrlf &_
    "    Public Property Let Value(p_Fields,p_Data)"& vbcrlf &_
    "		DALLetValue dal."& tbl("varname") &",p_Fields,p_Data"& vbcrlf &_
    "	end Property"& vbcrlf
	    for each field in arrFields
		    command=command & "    Public Property Let "&dal_info(tbl("name"))(field)("varname")&"(p_Data)"& vbcrlf &_
    "        m_fld"&GoodFieldName(field)&" = p_Data"& vbcrlf &_
    "        m_ChangedFields(""" & replace(field,"""","""""") & """)=true"& vbcrlf &_
    "    End Property"& vbcrlf
        next
        command=command & "End Class"& vbcrlf &_
        "set "& tbl("varname") & " = new class_"& GoodFieldName(tbl("name")) &""
        execute command
	end function
end class

dim dal
set dal = new tDAL

Function CustomQuery(dalSQL)
	dim dal_rs
    Set dal_rs = server.CreateObject("ADODB.Recordset")
    dal_rs.open dalSQL, dbConnection
    set CustomQuery=dal_rs
End Function

Function UsersTableName()
    UsersTableName=""
End Function

function cutprefix(tablename)
	dim pos
	pos=instr(tablename,".")
	if pos>0 then
		cutprefix=mid(tablename,pos+1)
	else
		cutprefix=tablename
	end if
end function

sub DALAdd(obj)
	dim rs,insertValues,insretFields,field,fld,tableinfo,arrFields,command
	insertValue=""
	insretParam=""
	set tableinfo=dal_info(obj.m_TableName)
	arrFields=tableinfo.Keys()

	if SQLUpdateMode then
	'	prepare WHERE params
		for each fld in arrFields
			if obj.m_ChangedFields.Exists(fld) then
				command="obj.Value(fld) = obj.m_fld" & GoodFieldName(fld)
				execute command
			end if
			for each field in obj.m_Value
				if ucase(fld)=ucase(field) then
					insertValues=insertValues & obj.PrepareValue(obj.m_Value(fld),tableinfo(fld)("nType")) & ", "
					insertFields=insertFields & AddFieldWrappers(fld) & ", "
					exit for
				end if
			next
		next
	 
		if bValue(insertValues) then _
			insertValues = mid(insertValues,1,len(insertValues)-2)
		if bValue(insertFields) then _
			insertFields = mid(insertFields,1,len(insertFields)-2)
		if bValue(insertValues) and bValue(insertFields) then
			dalSQL = "insert into " & AddTableWrappers(obj.TableName()) & " (" & insertFields & ") values (" & insertValues & ")"
			db_exec dalSQL,dbConnection
		end if
		DALReset(obj)
		exit sub
	end if
	
	
	Set rs = server.CreateObject("ADODB.Recordset")
	rs.Open "select * from " & obj.TableName() & " where 1=0", dbConnection, 1,2
	rs.Addnew
	
	for each fld in arrFields
		if obj.m_ChangedFields.Exists(fld) then
			command="rs(fld)=obj.m_fld" & GoodFieldName(fld)
			execute command
		elseif obj.m_Value.Exists(fld) then 
			rs(fld)=obj.m_Value(fld)
		elseif obj.m_Param.Exists(fld) then
			rs(fld)=obj.m_Param(fld)
		end if
	next
'	do add
	rs.Update
	rs.Close
	DALReset(obj)
End Sub


sub DALDelete(obj)
    dim deleteFields
	deleteFields=""
	dim tableinfo
	set tableinfo=dal_info(obj.m_TableName)
	dim arrFields
	arrFields=tableinfo.keys()
	dim value,isset
'	prepare WHERE params
	dim command
	for each fld in arrFields
		isset=true
		if obj.m_ChangedFields.Exists(fld) then
           	command="value = obj.m_fld" & GoodFieldName(fld)
			execute command
		elseif obj.m_Param.Exists(fld) then
			value=obj.m_Param(fld)
		elseif obj.m_Value.Exists(fld) then 
			value=obj.m_Value(fld)
		else
			isset=false
		end if
		if isset then
			if NeedQuotes(tableinfo(fld)("nType")) then 
            	deleteFields=deleteFields & AddFieldWrappers(fld) & "='" & db_addslashes(value) & "' and "
			else
            	deleteFields=deleteFields & AddFieldWrappers(fld) & "=" & value & " and "
			end if
		end if
	next
'	do delete
	if deleteFields<>"" then
		deleteFields=left(deleteFields,len(deleteFields)-5)
		dalSQL="delete from " & obj.TableName() & " where " & deleteFields
		dbConnection.Execute dalSQL
	end if
	DALReset(obj)
end sub

sub DALUpdate(obj)
	dim updateParam,updateValue,field,fld,arrFields,tableinfo,command
	updateParam=""
	updateValue=""
	set tableinfo=dal_info(obj.m_TableName)
	arrFields=tableinfo.keys()
'	prepare WHERE params
	for each fld in arrFields
		if obj.m_ChangedFields.Exists(fld) then
			if tableinfo(fld).Exists("key") then
				command="obj.Param(fld) = obj.m_fld" & GoodFieldName(fld)
			else
				command="obj.Value(fld) = obj.m_fld" & GoodFieldName(fld)
			end if
			execute command
		end if
		for each field in obj.m_Param
			if ucase(fld)=ucase(field) then
				updateParam=updateParam & AddFieldWrappers(fld) & "=" & obj.PrepareValue(obj.m_Param(fld),tableinfo(fld)("nType")) & " and "
				exit for
			end if
		next
		for each field in obj.m_Value
			if ucase(fld)=ucase(field) then
				updateValue=updateValue & AddFieldWrappers(fld) & "=" & obj.PrepareValue(obj.m_Value(fld),tableinfo(fld)("nType")) & " and "
				exit for
			end if
		next
	next
	 
	if SQLUpdateMode then
		if bValue(updateParam) then _
			updateParam = mid(updateParam,1,len(updateParam)-5)
		if bValue(updateValue) then _
			updateValue = mid(updateValue,1,len(updateValue)-5)
		if bValue(updateValue) and bValue(updateParam) then
			dalSQL = "update " & AddTableWrappers(obj.TableName()) & " set " & updateValue & " where " & updateParam
			db_exec dalSQL,dbConnection
		else
			DALReset(obj)
		end if
		exit sub
	end if
	
	
	if updateParam="" then
		DALReset(obj)
		exit sub
	end if

'	construct SQL, do select
	dim rs
	Set rs = server.CreateObject("ADODB.Recordset")
	dim strSQL
	updateParam = mid(updateParam,1,len(updateParam)-5)
	strSQL = "select * from " & obj.TableName() & " where " & updateParam
	rs.Open strSQL, dbConnection, 1,2
	while not rs.EOF
    '	prepare values
	    for each fld in arrFields
		    if obj.m_ChangedFields.Exists(fld) and not tableinfo(fld).Exists("key") then
			    command="rs(fld)=obj.m_fld" & GoodFieldName(fld)
			    execute command
		    elseif obj.m_Value.Exists(fld) then 
			    rs(fld)=obj.m_Value(fld)
		    end if
	    next
	    rs.Update
	    rs.movenext
	wend
	rs.Close
	DALReset(obj)
End Sub

Function DALQueryAll(obj)
	set DALQueryAll=DALQuery(obj,"","")
End Function

function DALQuery(obj,sw,ob)
	dim dalSQL,rs
	dalSQL = "select * from " & obj.TableName()
	if sw<>"" then _
		dalSQL = dalSQL & " where " & sw
	if ob<>"" then _
		dalSQL = dalSQL & " order by " & ob
	
	Set rs = server.CreateObject("ADODB.Recordset")
	rs.open dalSQL, dbConnection
	set DALQuery=rs
End Function

Function DALFetchByID(obj)
    dim dalWhere
	dalWhere=""
	dim tableinfo
	set tableinfo=dal_info(obj.m_TableName)
	dim arrFields
	arrFields=tableinfo.keys()
	dim value,isset
'	prepare WHERE params
	dim command
	for each fld in arrFields
		isset=true
		if obj.m_ChangedFields.Exists(fld) then
           	command="value = obj.m_fld" & GoodFieldName(fld)
			execute command
		elseif obj.m_Param.Exists(fld) then
			value=obj.m_Param(fld)
		elseif obj.m_Value.Exists(fld) then 
			value=obj.m_Value(fld)
		else
			isset=false
		end if
		if isset then
			if NeedQuotes(tableinfo(fld)("nType")) then 
            	dalWhere=dalWhere & AddFieldWrappers(fld) & "='" & db_addslashes(value) & "' and "
			else
            	dalWhere=dalWhere & AddFieldWrappers(fld) & "=" & value & " and "
			end if
		end if
	next
	DALReset(obj)
'	do fetch
	if dalWhere="" then
		exit function
	end if
	dalWhere=left(dalWhere,len(dalWhere)-5)
	dalSQL="select * from " & obj.TableName() & " where " & dalWhere
	rs.open dalSQL, dbConnection
	set DALFetchByID = rs
End Function

sub DALReset(obj)
	obj.m_ChangedFields.RemoveAll
	obj.m_Param.RemoveAll
	obj.m_Value.RemoveAll
End Sub

sub DALLetParam(obj,p_Fields,p_Data)
	dim tableinfo
	set tableinfo=dal_info(obj.m_TableName)
	dim arrFields
	arrFields=tableinfo.keys()
	dim value,isset
'	assign params
	for each fld in arrFields
		if ucase(fld)=ucase(p_Fields) then
			obj.m_Param(fld)=p_Data
			exit sub
		end if
	next
End sub

sub DALLetValue(obj,p_Fields,p_Data)
	dim tableinfo
	set tableinfo=dal_info(obj.m_TableName)
	dim arrFields
	arrFields=tableinfo.keys()
	dim value,isset
'	assign values
	for each fld in arrFields
		if ucase(fld)=ucase(p_Fields) then
			obj.m_Value(fld)=p_Data
			exit sub
		end if
	next
End sub
	
%>