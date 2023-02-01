<%
Set dbConn = Server.CreateObject("ADODB.Connection")

function getImportTableName(contrl_name)
	dim parsed,fso,psFilePath,strConn,adoWbkAsDatabase,tableName
	parsed=0
	contrl_name=replace(contrl_name,"file_","")
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	psFilePath = GetRequestForm("value_" & contrl_name)
	sPath = (server.mappath("tmp"))
	psFilePath = fso.BuildPath(sPath,psFilePath)

	runner_save_file psFilePath, GetRequestForm("file_" & contrl_name)
	if UCase(Right(psFilePath,4))=".XLS" then
		strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & psFilePath & ";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"";"
		dbConn.Open strConn
		Set adoWbkAsDatabase = Server.CreateObject("ADOX.Catalog")
		adoWbkAsDatabase.ActiveConnection = dbConn
		Set adoTables = adoWbkAsDatabase.Tables
		tableName=addTableWrappers(adoTables(0).Name)
		getImportTableName = tableName
	else
		strConn = "Driver={Microsoft Text Driver (*.txt; *.csv)};Dbq=" & sPath & ";Extensions=asc,csv,tab,txt" 
		dbConn.Open strConn
		tableName = GetRequestForm("value_" & contrl_name)
		getImportTableName = sPath & "/" & tableName
	end if
end function

function openImportExcelFile(TableName)
	dim rs
	doAssignmentByRef rs,db_query("select * from " & TableName, dbConn)
	doAssignment openImportExcelFile,rs
end function

function getImportExcelFields(rs)
	dim fields, i, field
	set fields = CreateDictionary()
	doAssignmentByRef data,db_fetch_array(rs)
	For Each field In data.keys
		fields(asp_count(fields)) = trim(field)
	Next 
	rs.movefirst
	doAssignment getImportExcelFields,fields
end function

function getImportExelData(rs,fields)
	dim j,i,arr,ret
	i=0
	while bValue(doAssignmentByRef(data,db_fetch_numarray(rs)))
		set arr = CreateDictionary()
		for j = 0 to asp_count(data.keys)
			if bValue(IsDateFieldType(GetFieldType(fields(j),""))) then
				data(j)=str_format_datetime(db2time(data(j)))
			end if
			arr(fields(j)) = data(j)
		next
		ret = InsertRecord(arr, i)
		i=i+1
		total_records=total_records+1
	wend
	getImportExelData=ret
end function
function db_exec_import(sql,conn)
	dim err_num
	On Error Resume Next
	conn.Execute sql
	err_num=Err.Number
	Err.Clear
	On Error Goto 0
	if err_num=0 then
		db_exec_import=true
	else
		db_exec_import=false
	end if
end function	

function getImportCVSFields(uploadfile)
	dim handle,arr,k
	doAssignmentByRef handle,asp_fopen(uploadfile,"r")
	doAssignment arr,explode(",",handle.ReadLine)
	for each k in arr.keys
		if left(arr(k),1)="""" and right(arr(k),1)="""" then 
			arr(k)=mid(arr(k),2,len(arr(k))-2)
			arr(k)=replace(arr(k),"""""","""")
		end if
	next 
	doAssignment getImportCVSFields,arr
end function

function readCSVLine(byref handle)
	dim line,str,evenQuotes,pos
'	read lines from the file
	line=""
	evenQuotes = true
	do
		if handle.AtEndOfStream then
			exit do
		end if
		str = handle.ReadLine
		if len(line)>0 then
			line = line & vbcrlf
		end if
		line = line & str
		pos=0
' count quotes
		do while true
			pos = instr(pos+1,str,"""")
			if pos=0 then 
				exit do
			end if
			evenQuotes = not evenQuotes
		loop
	loop while not evenQuotes

	readCSVLine = trim(line)
end function

function fgetcsv(handle,blen,delim)
	dim out,line, inQuotes, c, i, value, valueBegin, valueLen
	line = readCSVLine(handle)
	if len(line)=0 then
		fgetcsv=false
		exit function
	end if
	valueBegin = 1
	inQuotes = false
	set out = CreateDictionary()
'	cut values
	for i=1 to len(line)
		c = mid(line,i,1)
		if c="""" then
			inQuotes = not inQuotes
		end if
		if not inQuotes and c=delim or i=len(line) then
			valueLen = i - valueBegin
			if i=len(line) then
				valueLen = valueLen + 1
			end if
			value = mid(line,valueBegin, valueLen)
			valueBegin = i+1
	'prepare value
			if len(value)>1 and left(value,1)="""" and right(value,1)="""" then
				value=mid(value,2,len(value)-2)
				value=replace(value,"""""","""")
			end if
			value = trim(value)
			out.Add out.Count, value
		end if
	next
	set fgetcsv = out
end function
%>