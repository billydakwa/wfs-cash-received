<!--#include file="json.asp"-->
<%
sortgroup=0
sortorder="a"

dim output_buffer(20)
ob_enabled=0

set included_files = CreateDictionary()

errorhappened = false

sub sendmail(email, subject, message)
    dim tmpDict
    set tmpDict = CreateObject("Scripting.Dictionary")
    tmpDict("to")=email
    tmpDict("subject")=subject
    tmpDict("body")=message
    runner_mail tmpDict
end sub

' ASPRunnerPro mail function.
' "params" is a Scripting.Dictionary object with input parameters.
' The following parameters are supported:
'	"from" - Sender email address. If none specified an email address from the wizard will be used.
'	"to" - Receiver email address.
'	"body" - Plain text message body.
'	"htmlbody" - Html message body (do not use 'body' parameter in this case).
' Setting character set is not supported.
'
' Returns a Scripting.Dictionary object with the following data:
' "mailed" - indicates wheter mail sent or not
' "source" - error source (a COM object usually)
' "number" - error number
' "description" - error description
' "message" - formatted message with information above
Function runner_mail(params)
	On Error Resume Next
	Dim email_from, email_to, email_body, email_htmlbody, email_charset, email_ishtml, email_cc, email_bcc, csmtpserver, csmtpport, csmtppassword, csmtpuser

	csmtpserver = "localhost"
	csmtpport = CSmartLng("25")
	csmtppassword = ""
	csmtpuser = ""
	
	If VarType(params("from")) = vbEmpty or VarType(params("from")) = vbNull Then
		email_from = ""
	Else
		email_from = params("from")
	End If

	If VarType(params("to")) = vbEmpty or VarType(params("to")) = vbNull Then
		strMessage = "Email address is empty. Cannot send email."
		Exit Function
	Else
		email_to = params("to")
	End If
	
	email_cc=""
	If VarType(params("cc")) <> vbEmpty and VarType(params("cc")) <> vbNull Then
		email_cc = params("cc")
	End If
	
	email_bcc=""
	If VarType(params("bcc")) <> vbEmpty and VarType(params("bcc")) <> vbNull Then
		email_bcc = params("bcc")
	End If

	
	email_ishtml = false
	email_subject = params("subject")
	email_body = ""
	
	If VarType(params("body")) = vbEmpty or VarType(params("body")) = vbNull Then
		If Not (VarType(params("htmlbody")) = vbEmpty or VarType(params("htmlbody")) = vbNull) Then
			email_body = params("htmlbody")
		End If
		email_ishtml = true
	Else
		email_body = params("body")
	End If
	
	Version = Request.ServerVariables("SERVER_SOFTWARE")
	If InStr(Version, "Microsoft-IIS") > 0 Then
		i = InStr(Version, "/")
		If i > 0 Then
			IISVer = Trim(Mid(Version, i+1))
		End If
	End If

	Err.Clear
	dim myMail
'	Roadmap for CDO library
'	http://msdn.microsoft.com/en-us/library/ms978698.aspx
	Set myMail=CreateObject("CDO.Message")
	If err.Number=0 Then
        myMail.Subject = email_subject 
        myMail.From = email_from
        myMail.To = email_to
        
        if email_cc<>"" then _
            myMail.Cc = email_cc
        if email_bcc<>"" then _
            myMail.Bcc = email_bcc
            
		If email_ishtml Then
			myMail.HTMLBody = email_body
		Else
			myMail.TextBody = email_body
		End If
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")=2
        'Name or IP of remote SMTP server
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver")=csmtpserver
        'Server port
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=csmtpport 

	if csmtpport = 465 then
	   myMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1
	   myMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	end if

        ' SMTP username and passwords
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = csmtppassword
        myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = csmtpuser
		if csmtpuser<>"" then _
			myMail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        myMail.Configuration.Fields.Update
        myMail.Send
        Set myMail = Nothing 
	Else
		Set myMail = Server.CreateObject("CDONTS.NewMail")
		myMail.From = email_from
		myMail.To = email_to
        if email_cc<>"" then _
            myMail.Cc = email_cc
        if email_bcc<>"" then _
            myMail.Bcc = email_bcc
		myMail.Subject = email_subject
		myMail.Body = email_body
		
		If email_ishtml Then
			myMail.BodyFormat = 0
		Else
			myMail.BodyFormat = 1
		End If
		
		myMail.Send
		Set myMail = Nothing
	End If

	dim result
	set result = CreateObject("Scripting.Dictionary")
	if Err.Number<>0 then
		result("mailed") = False
		result("source") = Err.Source
		result("number") = Err.Number
		result("description") = Err.Description
		result("message") = "Error happened sending email to " & email_to & "<br>" & Err.Source & "<br>" & Err.Number & "<br>" & Err.Description
		Set runner_mail = result
		Err.Clear
	Else
		result("mailed") = True
	end if
	Set runner_mail = result
	on error goto 0
End Function

'//// TEST CODE
'Dim csmtpserver, csmtpport, csmtppassword, csmtpuser
'csmtpserver = "localhost"
'csmtpport = 25
'csmtppassword = "123"
'csmtpuser = "user"
'Dim dicTest
'Set dicTest = CreateObject("Scripting.Dictionary")
'dicTest.Add "from", "user@test.com"
'dicTest.Add "to", "to@test.com"
'dicTest.Add "htmlbody", "<html><body>??? ?????</body></html>"
'dicTest.Add "subject", "Hello"
'runner_mail(dicTest)


'///////////////////////////////////////////////////////////////////////////////
Sub printfile(filename)
	if instr(filename,"\")=0 then
		filename=getabspath(filename)
	end if
	Dim objStream
	set objStream = Server.CreateObject("ADODB.Stream")
	objStream.Type = 1
	objStream.Open
	objStream.LoadFromFile filename
	Response.BinaryWrite objStream.Read
	set objStream = Nothing
End Sub

'///////////////////////////////////////////////////////////////////////////////
function CreateThumbnail(value, size, ext)
	dim jpeg
	SafeCreateObject "Persits.Jpeg", jpeg
	if isnull(jpeg) then
		CreateThumbnail=value
		exit function
	end if
	on error resume next
	Jpeg.OpenBinary value
	if err.number<>0 then
		CreateThumbnail=value
		on error goto 0
		exit function
	end if
	on error goto 0
	dim sx,sy
	sx = Jpeg.OriginalWidth
	sy = Jpeg.OriginalHeight
	if sx<=size and sy<=size or sx=0 or sy=0 then
		CreateThumbnail=value
		exit function
	end if
	if sx>=sy then
		jpeg.Height=sy*size/sx
		jpeg.Width=size
	else
		jpeg.Width=sx*size/sy
		jpeg.Height=size
	end if

	dim ret
	CreateThumbnail=Jpeg.Binary
end function

sub SafeCreateObject(name,object)
	on error resume next
	set object = server.CreateObject(name)
	if err.Number<>0 then
		object=null
	end if
	on error goto 0
end sub

'///////////////////////////////////////////////////////////////////////////////
Function myfile_get_contents(filename,p)
	myfile_get_contents=""
	dim stream
	set stream=Server.CreateObject("ADODB.Stream")
	stream.CharSet=cCharset
	stream.type=2
	on error resume next
	stream.Open
	if err.Number<>0 then
		err.Clear
		set stream=nothing
		on error goto 0
		exit function
	end if 
	on error goto 0
	stream.LoadFromFile Filename
	myfile_get_contents = stream.ReadText
	stream.Close
	set stream=nothing
End Function

Sub myfile_put_contents(filename, contents)
	Const adTypeBinary = 1
	Const adSaveCreateOverWrite = 2
  
	'Create Stream object
	Dim BinaryStream
	Set BinaryStream = CreateObject("ADODB.Stream")
  
	'Specify stream type - we want To save binary data.
	BinaryStream.Type = adTypeBinary
  
	'Open the stream And write binary data To the object
	BinaryStream.Open
	BinaryStream.Write ByteArray
  
	'Save binary data To disk
	BinaryStream.SaveToFile FileName, adSaveCreateOverWrite
	
	Set BinaryStream = Nothing
End Sub

'///////////////////////////////////////////////////////////////////////////////
Function myfile_exists(filename)
	Set fso = CreateObject("Scripting.FileSystemObject")
	myfile_exists = fso.FileExists(getabspath(filename))
	set fso = Nothing
End Function


'///////////////////////////////////////////////////////////////////////////////
Sub runner_delete_file(strFileName)
	Set fso = CreateObject("Scripting.FileSystemObject")
	if fso.FileExists(strFileName) then
		fso.DeleteFile(strFileName)
	end if
	set fso = Nothing
End Sub

'///////////////////////////////////////////////////////////////////////////////
Function mysprintf(format, params)
	Dim c, informat, formatchar, intzerobegin, intdecimals, formatnum, out
	informat = false
	intzerobegin = false
	intdecimals = 0
	formatnum = 0
	out = ""
	
	For i = 1 To Len(format)
		c = Mid(format, i, 1)
		Select Case c
		Case "%"
			If informat Then
				' error
				Response.Write "Invalid character in format"
				Response.End
			Else
				informat = true
			End If
		Case "s"
			If informat Then
				out = out & params(formatnum)
				
				informat = false
				intzerobegin = false
				intdecimals = 0
				formatnum = formatnum + 1
			Else
				out = out & "s"
			End If
		Case "d"
			If informat Then
				If intdecimals > 0 And Not intzerobegin Then
					' error
					' format "%4d" (for example) is wrong
					' (should be "%04d")
					Response.Write "Wrong decimal format"
					Response.End
				End If
				
				Dim s, ndot
				s = CStr(params(formatnum))
				ndot = InStr(1, s, ".")
				If ndot > 0 Then
					s = Mid(s, 1, ndot - 1)
				End If
				If intdecimals > 0 And Len(s) < intdecimals Then
					s = String(intdecimals - Len(s), "0") & s
				End If
				out = out & s
				
				informat = false
				intzerobegin = false
				intdecimals = 0
				formatnum = formatnum + 1
			Else
				out = out & "d"
			End If
		Case Else
			If informat Then
				If c = "0" Then
					intzerobegin = true
				ElseIf c = "1" Or c = "2" Or c = "3" Or c = "4" Or c = "5" _
						Or c = "6" Or c = "7" Or c = "8" Or c = "9" Then
							intdecimals = CLng(c)
				Else
					' error
					Response.Write "Invalid character in format"
					Response.End
				End If
			Else
				out = out & c
			End If
		End Select
	Next
	mysprintf = out
End Function

'//// TEST CODE
'Response.Write mysprintf("%d", Array(1)) & "<br/>"
'Response.Write mysprintf("%05d", Array(1)) & "<br/>"
'Response.Write mysprintf("%04d-%02d-%02d %02d:%02d:%02d", Array(2000, 2, 1, 13, 11, 59)) & "<br/>"
'Response.Write mysprintf("s%sss", Array("Hello")) & "<br/>"
'Response.Write mysprintf("s-%s..%s-s", Array("Hello", "again")) & "<br/>"
'Response.Write mysprintf("s-%%s..%s-s", Array("Hello", "again")) & "<br/>"
'Response.Write mysprintf("%5d", Array(1)) & "<br/>"

function GetRequestValue(byref arr,byval key)
		if typename(arr)="IRequest" then
		    doAssignment GetRequestValue,GetRequestValue(request.QueryString,key)
		    if vartype(GetRequestValue)=vbEmpty then
    		    doAssignment GetRequestValue,GetRequestValue(RequestForm(),key)
    		end if
    		exit function
		end if
		if vartype(arr(key))=vbEmpty then
			if vartype(arr(key & "[]"))<>vbEmpty then
				doAssignment GetRequestValue,arr(key & "[]")
				exit function
			end if
		end if
		GetRequestValue=arr(key)
end function

function RequestForm()
	if left(request.ServerVariables("CONTENT_TYPE"),9)="multipart" then
		if formParsed<>1 then
			if ParseMultiPartForm()=true then _
			    formParsed=1
		end if
	end if
	if formParsed<>1 then
		set RequestForm=request.form
	else
		set RequestForm=myRequest
	end if
end function

function GetCollectionBounds(byref arr, byref first, byref last)
	if IsDictionary(arr) then
		first=0
		last=arr.count-1
		exit function
	end if
	if IsArray(arr) then
		first=lbound(arr)
		last=ubound(arr)
		exit function
	end if
	if typename(arr)="IRequestDictionary" or typename(arr)="IStringList" then
		first=1
		last=arr.count
		exit function
	end if
	if typename(arr)="ISessionObject" then
		first=1
		last=arr.contents.count
		exit function
	end if
end function

function GetCollectionKey(byref arr,byval index)
	if IsDictionary(arr) then
		GetCollectionKey = arr.keys()(index)
		exit function
	end if
	if typename(arr)="IRequestDictionary" then
		GetCollectionKey = arr.key(index)
		exit function
	end if
	if typename(arr)="IStringList" then
		GetCollectionKey=index
		exit function
	end if
	if typename(arr)="ISessionObject" then
		GetCollectionKey = arr.contents.key(index)
		exit function
	end if
end function

Function Unicode2Bytes(str)
    dim ind	
    For ind = 1 To len(str) 
        Unicode2Bytes = Unicode2Bytes& ChrB(Asc(Mid(str, ind, 1)))
    Next 
End Function

Function SupposeImageType(file)
	If LenB(file) > 1 And MidB(file, 1, 2) = chrb(asc("B")) & chrb(asc("M"))  Then
    	SupposeImageType = "image/bmp"
		Exit Function
	End If
	If LenB(file) > 2 And MidB(file, 1, 3) = chrb(asc("G")) & chrb(asc("I"))& chrb(asc("F"))  Then
		SupposeImageType = "image/gif"
		Exit Function
    End If
	if LenB(file) > 3 and  MidB(file, 1, 3) = chrb(&Hff) & chrb(&Hd8) & chrb(&Hff) then
		SupposeImageType = "image/jpeg"
		Exit Function
    End If
	if LenB(file) > 8 and MidB(file, 1, 8) = chrb(&H89) & chrb(&H50) & chrb(&H4e) & chrb(&H47) _
											& chrb(&H0d) & chrb(&H0a) & chrb(&H1a) & chrb(&H0a)  then
		SupposeImageType = "image/png"
		Exit Function
    End If
    SupposeImageType=""
End Function


function bValue(ByVal val)
	dim vt
	vt = vartype(val)
	if vt=vbEmpty or vt=vbNull then
		bValue=false
	elseif vt=vbBoolean then
		bValue=val
	elseif vt=vbString then
		bValue=(Len(val)>0 and val<>"0")
	elseif IsNumeric(val) then
		bValue=CBool(val)
	elseif vt=vbObject then
		if IsDictionary(val) then
			if val.Count>0 then 
				bValue=true
			else
				bValue=false
			end if
			exit function
		end if
		bValue=true
	else
		bValue=true
	end if
end function

function doAssignment(ByRef var,ByRef value)
	if not isobject(value) then
		var=value
		doAssignment=value
	elseif IsDictionary(value) then
		copyDictionary value,var
		set doAssignment=var
	else
		set var=value
		set doAssignment=value
	end if
end function

function doAssignmentByRef(ByRef var,ByRef value)
	if not isobject(value) then
		var=value
		doAssignmentByRef=value
	else
		set var=value
		set doAssignmentByRef=value
	end if
end function


function setArrElement(ByRef arr,ByVal key,ByRef value)
	if not isobject(arr) then _
		set arr=CreateDictionary()
	dim tval
	doAssignment tval,value
	if vartype(key)=vbString then
		if IsNumeric(key) then
			key=CLng(key)
		end if
	end if
	if not IsObject(value) then
		arr(key)=tval
		setArrElement=tval
	else
		set arr(key)=tval
		setArrElement=bValue(tval)
	end if
end function


function setArrElementByRef(ByRef arr,Byval key,ByRef value)
	if isempty(arr) then _
		set arr=CreateDictionary()
	if vartype(key)=vbString then
		if IsNumeric(key) then
			key=CLng(key)
		end if
	end if
	if not IsObject(value) then
		arr(key)=value
		setArrElementByRef=value
	else
		set arr(key)=value
		setArrElementByRef=bValue(value)
	end if
end function


function doClassAssignmentByRef(ByRef obj,ByVal key,ByRef value)
	dim str1,str2
	str1=""
	str2=""
	if IsObject(value) then 
		str1="Set "
		str2="Set "
	end if
	str1=str1 & "obj." & key & " = value"
	str2=str2 & "doClassAssignmentByRef = value"
	Execute str1
	Execute str2
end function

function doClassAssignment(ByRef obj,ByVal key,ByRef value)
	dim str1,str2,tval
	doAssignment tval,value
	str1=""
	str2=""
	if IsObject(value) then 
		str1="Set "
		str2="Set "
	end if
	str1=str1 & "obj." & key & " = tval"
	str2=str2 & "doClassAssignment = tval"
	Execute str1
	Execute str2
end function

function setArrElementN_Int(ByRef arr,byref pkeys,ByRef value,byreference)
    dim tarr,i
    ensureArrayCreated arr,pkeys(0)
    set tarr=arr(pkeys(0))
    for i=1 to pkeys.count-2
        ensureArrayCreated tarr,pkeys(i)
        set tarr=tarr(pkeys(i))
    next
    lastkey = pkeys(pkeys.count-1)
    if isEmpty(lastkey) then
        lastkey=asp_count(tarr)
    end if
    if byreference then 
        setArrElementByRef tarr,lastkey,value
    else
        setArrElement tarr,lastkey,value
    end if
    doAssignmentByRef setArrElementN_Int,value
end function

function setArrElementByRefN(ByRef arr,byref pkeys,ByRef value)
	doAssignmentByRef setArrElementByRefN,setArrElementN_Int(arr,pkeys,value,true)
end function

function setArrElementN(ByRef arr,byref pkeys,ByRef value)
	doAssignmentByRef setArrElementN,setArrElementN_Int(arr,pkeys,value,false)
end function


function ensureArrayCreated(byref arr, byval key)
	if not IsObject(arr) then _
		set arr=CreateDictionary
	if isobject(arr(key)) then
		exit function
	end if
	set arr(key) = CreateDictionary()
end function


function postInc(ByRef var)
	postInc=var
	var=var+1
end function
function postDec(ByRef var)
	postDec=var
	var=var-1
end function
function preInc(ByRef var)
	var=var+1
	preInc=var
end function
function preDec(ByRef var)
	var=var-1
	preDec=var
end function


'	array function routines

function CreateDictionary()
	set CreateDictionary=Server.CreateObject("Scripting.Dictionary")
end function

function GetCreateDictionaryString(n,numeric)
	dim body
	dim funcname
	if not numeric then
		funcname="CreateDictionary"&n
	else
		funcname="CreateArray"&n
	end if
	body = "set "&funcname&"=Server.CreateObject(""Scripting.Dictionary"")" & vbcrlf
	body = body & "dim counter" &  vbcrlf
	body = body & "counter=0" & vbcrlf
	dim i,params
	for i=1 to n
		if i>1 then _
			params= params & ","
		if not numeric then
			params = params & "name" &i & ",param" & i
			body = body & "	if not isEmpty(name"&i&") then "&vbcrlf &_
			"setArrElement "&funcname&",name"&i&",param"&i&vbcrlf &_
			"else "&vbcrlf
		else
			params = params & "param" & i
		end if
		body = body & "setArrElement "&funcname&",counter,param"&i&vbcrlf &_
		"counter=counter+1" & vbcrlf
		if not numeric then
			body = body & "end if"&vbcrlf
		end if
	next
	GetCreateDictionaryString  = "function "&funcname&"("&params&")"&vbcrlf & body & vbcrlf & "end function"
end function

dim arrsizes(6)
arrsizes(0)=1
arrsizes(1)=2
arrsizes(2)=3
arrsizes(3)=4
arrsizes(4)=5
arrsizes(5)=6

dim dictsizes(13)
dictsizes(0)=1
dictsizes(1)=2
dictsizes(2)=3
dictsizes(3)=4
dictsizes(4)=5
dictsizes(5)=6
dictsizes(6)=7
dictsizes(7)=8
dictsizes(8)=9
dictsizes(9)=12
dictsizes(10)=16
dictsizes(11)=24
dictsizes(12)=42

for nCDF=0 to ubound(arrsizes)
	execute GetCreateDictionaryString(arrsizes(nCDF),true)
next
for nCDF=0 to ubound(dictsizes)
	execute GetCreateDictionaryString(dictsizes(nCDF),false)
next


function CreateClass(classname,pcount,param1,param2,param3,param4,param5,param6,param7)
dim str,i
	str="set CreateClass = new " & classname & vbcrlf
	str=str+"CreateClass.init_" & classname 
	if pcount>0 then _ 
		str=str & "_p" & pcount
	for i = 1 to pcount
		str=str+" param" & i
		if i<pcount then
			str=str+","
		end if
	next
	execute str
end function

function IIF(ByVal cond,ByRef expr1,ByRef expr2)
	if bValue(cond) then
		if vartype(expr1)=vbObject then
	        set iif=expr1
	    else
		    iif=expr1
		end if
	else
		if vartype(expr2)=vbObject then
	        set iif=expr2
	    else
		    iif=expr2
		end if
	end if
end function

function IsSet(ByRef var)
	if VarType(var)=vbEmpty or VarType(var)=vbNull then
		IsSet=false
	elseif VarType(var)=vbBoolean then
		IsSet=var
	end if
end function

function asp_stripos(str,substr,start)
    asp_stripos=asp_strpos(lcase(str),lcase(substr),start)
end function

function asp_strpos(str,substr,start)
	dim pos
	if isnull(str) then
		asp_strpos=str
		exit function
	end if
	if isEmpty(start) then start=0
	pos=instr(start+1,str,substr)
	if isEmpty(pos) then	
		asp_strpos=false
		exit function
	end if
	if pos=0 then
		asp_strpos=false
		exit function
	end if
	asp_strpos=pos-1
end function

function asp_strrpos(str,substr,a)
	dim pos
	pos=instrrev(str,substr)
	if isEmpty(pos) then	
		asp_strrpos=false
		exit function
	end if
	if pos=0 then
		asp_strrpos=false
		exit function
	end if
	asp_strrpos=pos-1
end function


function asp_array_splice(p_arr,offset,length)
    if offset>=0 and length>=0 then
        dim tmpDict, i, l
        l=0
        set tmpDict=CreateObject("Scripting.Dictionary")
        for each i in p_arr.keys
            if i<offset or i>=offset+length then
                setArrElement tmpDict,l,ArrayElement(p_arr,i)
                l=l+1
            end if
        next
        set p_arr=tmpDict
    end if
end function

function asp_unsetElement(ByRef p_arr,ByVal p_key)
	if IsDictionary(p_arr) then
		if vartype(p_key)=vbString then
			if IsNumeric(p_key) and len(p_key)<10 then _
				p_key=CLng(p_key)
		end if
	    if p_arr.exists(p_key) then _
		    p_arr.remove p_key
		exit function
	end if
	if TypeName(p_arr)="ISessionObject" then
		Session.contents.remove p_key
		exit function
	end if
	if TypeName(p_arr)="IRequestDictionary" then
		p_arr(p_key)=empty
	end if
end function

function asp_array_key_exists(ByVal p_key,ByRef p_arr)
    if IsDictionary(p_arr) then
		if vartype(p_key)=vbString then
			if IsNumeric(p_key) and len(p_key)<10 then _
				p_key=CLng(p_key)
		end if
        if p_arr.Exists(p_key) then
            asp_array_key_exists=true
            exit function
        else
            asp_array_key_exists=false
            exit function
        end if
    end if
	if TypeName(p_arr)="IRequest" then
		asp_array_key_exists = vartype(Request.QueryString(p_key))<>vbEmpty or  vartype(RequestForm()(p_key))<>vbEmpty
		exit function
	else
		on error resume next
		asp_array_key_exists = false
		asp_array_key_exists = vartype(p_arr(p_key))<>vbEmpty
		exit function
	end if
    asp_array_key_exists=false
end function

function asp_array_unique(p_arr)
    dim tmpDict, key, nkey, flag
    set tmpDict=CreateObject("Scripting.Dictionary")
    for each key in p_arr
        flag=0
        for each nkey in tmpDict
            if cstr(p_arr(key))=cstr(tmpDict(nkey)) then 
				flag=1
			end if
        next
        if flag=0 then 
			setArrElement tmpDict,key,p_arr(key)
		end if
    next
    set asp_array_unique=tmpDict
end function

function asp_is_array(p_arr)
    if IsObject(p_arr) then
		if IsDictionary(p_arr) then
		    asp_is_array=true
		    exit function
		end if
		if TypeName(p_arr)="IStringList" then
		    asp_is_array=true
		    exit function
		end if
    end if
    asp_is_array=false
end function

function asp_array_keys(p_arr, p_search)
    dim key, tmpDict
    set tmpDict=CreateObject("Scripting.Dictionary")
    for each key in p_arr.Keys
        if key=p_search or isEmpty(p_search) then
            tmpDict(tmpDict.Count)=key
        end if
    next
    set asp_array_keys=tmpDict
end function

function asp_count(byref p_arr)
	if not IsObject(p_arr) then
		if isArray(p_arr) then
		    asp_count=ubound(p_arr)
		else
        	if not isnull(p_arr) and not IsEmpty(p_arr) then
    	        asp_count=1
	        else
	    	    asp_count=0
		    end if
	    end if
	else
		if IsDictionary(p_arr) then
		    asp_count=p_arr.Count
			exit function
		end if
		dim tname
		tname=typename(p_arr)
		if tname="ISessionObject" then
		    asp_count=p_arr.Contents.Count
		elseif tname="IRequestDictionary" or tname="IStringList" or tname="IRequest" then
		    asp_count=p_arr.Count
		end if
	end if
end function

function asp_strlen(str)
    asp_strlen=len(CSmartStr(str))
end function

Function utf8_substr(ByVal str,ByVal from,ByVal len)
	doAssignmentByRef utf8_substr,asp_substr(str,from,len)
	Exit Function
End Function

function asp_substr(str,start,slen)
	if IsNull(str) or isEmpty(str) then
		asp_substr=false
		exit function
	end if
    dim tmpstr
    if isEmpty(slen) or slen>=0 then
        if len(str)<=start then
            asp_substr=false
            exit function
        end if
        if start>=0 then
            if not isEmpty(slen) and start+1+slen<=len(str) then
                asp_substr=mid(str,start+1,slen)
            else
                asp_substr=mid(str,start+1)
            end if
            exit function
        else
            dim tstart
            tstart=len(str)+start+1
            if tstart<1 then
                tstart=1
            end if
            if not isEmpty(slen) and slen<=abs(start) then
                asp_substr=mid(str,tstart,slen)
            else
                asp_substr=mid(str,tstart)
            end if
            exit function
        end if
    else
        slen=abs(slen)
        if start>=0 then
            tmpstr=mid(str,start+1)
        else
            start=abs(start)
            tmpstr=right(str,start)
        end if
        if len(tmpstr)<slen then
            asp_substr=false
            exit function
        end if
        asp_substr=left(tmpstr,len(tmpstr)-slen)
        exit function
    end if
end function

function asp_is_numeric(val)
    asp_is_numeric=(IsNumeric(val) and not IsEmpty(val))
end function

function asp_strtolower(str)
    asp_strtolower=LCase(str)
end function

function asp_strtoupper(str)
    asp_strtoupper=UCase(str)
end function

function asp_strrev(str)
    dim i, newstr
    for i=1 to len(str)
        newstr=mid(str,i,1) & newstr
    next
    asp_strrev=newstr
end function

function asp_nl2br(str)
	if IsNull(str) or isEmpty(str) then
		asp_nl2br=""
		exit function
	end if
	asp_nl2br=replace(str,vbcrlf,"<br />")
end function    

function asp_rawurlencode(str)
    asp_rawurlencode=SafeURLEncode(str)
end function    

function asp_rawurldecode(str)
	if IsNull(str) or isEmpty(str) then
		asp_rawurldecode=""
		exit function
	end if
	str = Replace(str, "+", " ")
    For i = 1 To Len(str) 
		sT = Mid(str, i, 1) 
		If sT = "%" Then 
			If i+2 <= Len(str) Then 
				sR = sR & Chr(CLng("&H" & Mid(str, i+1, 2))) 
				i = i+2 
			End If 
		Else 
			sR = sR & sT 
		End If 
	Next 
	asp_rawurldecode = sR 
end function


function asp_substr_replace(str, str_replace, start, length)
    dim tstr
	if IsNull(str) or isEmpty(str) then
		asp_substr_replace=""
		exit function
	end if
    if isEmpty(length) then _
        length=len(str)
    if len(str)=0 then
        asp_substr_replace=str_replace
        exit function
    end if
    if start>=0 and start+1>len(str) or start<0 and abs(start)>len(str) then
        asp_substr_replace=str
        exit function
    end if
    if start>=0 and length>=0 then
        asp_substr_replace=left(str,start) & str_replace & mid(str,start+length+1)
        exit function
    elseif start<0 and length>=0 then
        asp_substr_replace=mid(str,1,len(str)+start) & str_replace & mid(str,-start+length)
        exit function
    elseif start>=0 and length<0 then
        asp_substr_replace=left(str,start) & str_replace & right(str,abs(length))
        exit function
    elseif start<0 and length<0 then
        asp_substr_replace=mid(str,1,len(str)+start) & str_replace & right(str,abs(length))
        exit function        
    end if
end function

function asp_str_replace(str_search,str_replace,str)
    dim i, val, str_val
    if VarType(str)<>vbObject then 
		str=CSmartStr(str)
	end if
    if VarType(str_search)=vbString and VarType(str_replace)=vbString and VarType(str)=vbString then
        asp_str_replace=replace(str,str_search,str_replace)
        exit function
    elseif VarType(str_search)=vbString and VarType(str_replace)=vbString and VarType(str)=vbObject then
        for each val in str.Keys
            str(val)=replace(str(val),str_search,str_replace)
        next
        set asp_str_replace=str
        exit function
    elseif VarType(str_search)=vbObject and VarType(str_replace)=vbString and VarType(str)=vbString then
        for each val in str_search.Keys
            str=replace(str,str_search(val),str_replace)
        next
        asp_str_replace=str
        exit function
    elseif VarType(str_search)=vbObject and VarType(str_replace)=vbString and VarType(str)=vbObject then
        for each str_val in str.Keys
            for each val in str_search.Keys
                str(str_val)=replace(str(str_val),str_search(val),str_replace)
            next
        next
        set asp_str_replace=str
        exit function
    elseif VarType(str_search)=vbObject and VarType(str_replace)=vbObject and VarType(str)=vbString then
        for i=0 to str_search.Count
            if i<=str_replace.Count then
                str=replace(str,str_search.Item(i),str_replace.Item(i))
            else
                str=replace(str,str_search.Item(i),"")
            end if
        next
        asp_str_replace=str
        exit function
    elseif VarType(str_search)=vbObject and VarType(str_replace)=vbObject and VarType(str)=vbObject then
        for each val in str.Keys
            for i=0 to str_search.Count
                if i<=str_replace.Count then
                    val=replace(val,str_search.Item(i),str_replace.Item(i))
                else
                    val=replace(val,str_search.Item(i),"")
                end if
            next
        next
        set asp_str_replace=str
        exit function
    end if
end function

function asp_sizeof(p_arr)
    asp_sizeof=asp_count(p_arr)
end function

function asp_dirname(str)
    dim p1, p2, s
    if instr(1,str,"/")=0 and instr(1,str,"\")=0 then
        asp_dirname="."
    else
        if right(str,1)="/" or right(str,1)="\" then str=left(str,len(str)-1)
        p1=instrrev(str,"/")
        p2=instrrev(str,"\")
        if p1=0 and p2=0 then
            asp_dirname=str
        elseif p1>p2 then
            str=left(str,p1-1)
        else
            str=left(str,p2-1)
        end if
        asp_dirname=str
    end if
end function

function asp_file_exists(filename)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(filename) Then
        asp_file_exists=true
    Else
        If fso.FolderExists(filename) Then
            asp_file_exists=true
        else
            asp_file_exists=false
        end if
    End If
    set fso = Nothing
end function

function asp_header(str)
    dim p
    p=instr(1,str,":")
    if lcase(trim(left(str,p-1)))="location" then
        response.Redirect trim(mid(str,p+1))
    elseif lcase(trim(left(str,p-1)))="content-type" then
        response.ContentType=trim(mid(str,p+1))
    else
        Response.AddHeader trim(left(str,p-1)),trim(mid(str,p+1))
    end if
end function

function explode(patt,str)
    set explode=asp_split(patt,str)
end function

function asp_split(patt,str)
    dim arr, i
    set dict=CreateObject("Scripting.Dictionary")
    arr=split(str,patt)
    for i=0 to ubound(arr)
        setArrElement dict,i,arr(i)
    next
    set asp_split=dict
end function

function asp_ceil(val)
    if val=int(val) then
        asp_Ceil=val
    else
        asp_Ceil=int(val)+1
    end if
end function

function asp_floor(val)
    asp_floor=int(val)
end function

function asp_urldecode(sConvert)
	if IsNull(sConvert) or isEmpty(sConvert) then
		asp_urldecode=""
		exit function
	end if
    Dim aSplit
    Dim sOutput
    Dim I
     
    ' convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ")
     
    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")
     
    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        sOutput = sOutput & _
          Chr("&H" & Left(aSplit(i + 1), 2)) &_
          Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If
     
    asp_urldecode = sOutput
end function

function db_close(conn)
    conn.Close
    Set conn = Nothing
end function

function db_error()
	db_error=Err.Description
end function

function db_exec(sSQL,conn)
   if IsIdentical(dDebug,true) then response.write sSQL & "<br>"
    conn.Execute sSQL
    call ReportError
end function

function db_query(sSQL,conn)
    dim asp_rs
   if IsIdentical(dDebug,true) then response.write sSQL & "<br>"
    Set asp_rs = server.CreateObject("ADODB.Recordset")
    asp_rs.Open sSQL,conn
    call ReportError
    set db_query=asp_rs
end function

function db_query_direct(sSQL,conn,a)
    set db_query_direct=db_query(sSQL,conn)
end function

function db_fetch_array(asp_rs)
	doAssignmentByRef db_fetch_array,db_fetch_array_int(asp_rs,true)
end function

function db_fetch_numarray(asp_rs)
	doAssignmentByRef db_fetch_numarray,db_fetch_array_int(asp_rs,false)
end function

function NumberWithZero(n)
	if n<10 then
		NumberWithZero="0" & n
	else
		NumberWithZero=n
	end if
end function

function db_fetch_array_int(asp_rs,byname)
    dim field, tdate
    if asp_rs.EOF then
        db_fetch_array_int=false
        exit function
    end if
    dim rsDict
	dim i,value
	i=0
    set rsDict=CreateObject("Scripting.Dictionary")
    For Each field in asp_rs.Fields
        value=asp_rs.Fields(field.Name).Value
        if isnull(value) then 
            value=null
        elseif IsTimeType(field.type) then
			tdate=asp_rs.Fields(field.Name).Value
			value=NumberWithZero(hour(tdate)) & ":" & NumberWithZero(minute(tdate)) & ":" & NumberWithZero(second(tdate))
		elseif IsDateFieldType(asp_rs.Fields(field.Name).Type) then
			tdate=asp_rs.Fields(field.Name).Value
            if hour(tdate)=0 and minute(tdate)=0 and second(tdate)=0 then
                value=year(tdate) & "-" & NumberWithZero(month(tdate)) & "-" & NumberWithZero(day(tdate))
            elseif year(tdate)=1899 and month(tdate)=12 and day(tdate)=30 then
                value=NumberWithZero(hour(tdate)) & ":" & NumberWithZero(minute(tdate)) & ":" & NumberWithZero(second(tdate))
            else
                value=year(tdate) & "-" & NumberWithZero(month(tdate)) & "-" & NumberWithZero(day(tdate)) & " " & NumberWithZero(hour(tdate)) & ":" & NumberWithZero(minute(tdate)) & ":" & NumberWithZero(second(tdate))
            end if
		elseif vartype(value)=14 then
		    value=CDbl(value)
        end if
		if byname then
			rsDict(field.Name)=value
		else
			rsDict(i)=value
		end if
		i=i+1
    next
    asp_rs.Movenext
    set db_fetch_array_int=rsDict
end function

function asp_session_unset()
    session.Abandon
end function
    

Function CSmartDbl(strValue)
	dim vt
	vt = vartype(strValue)
	if vt>=2 and vt<=5 then
		CSmartDbl = strValue
		exit function
	end if
    if vt=vbBoolean then
		if strValue=true then
    	    CSmartDbl=1
        	exit function
	    end if
	end if
    On Error Resume Next    
    CSmartDbl = CDbl(strValue)
    if Err.Number<>0 then
	    Err.Clear
    	if InStr(strValue, ".")>0 then 
	       CSmartDbl = CDbl(Replace(strValue, ".", ","))
	    elseif InStr(strValue, ",")>0 then 
	        CSmartDbl = CDbl(Replace(strValue, ",", "."))
	    end if
	    Err.Clear
    end if
    On Error Goto 0
End Function

Function CSmartLng(strValue)
    if strValue=true then
        CSmartLng=1
        exit function
    end if
    On Error Resume Next    
    CSmartLng = CLng(strValue)
    if Err.Number<>0 then
	    Err.Clear
		CSmartLng=0
	    Err.Clear
    end if
    On Error Goto 0
End Function

Function CSmartLng(strValue)
    On Error Resume Next
    CSmartLng = CLng(strValue)
    if Err.Number<>0 then
	    Err.Clear
    	CSmartLng = 0
    end if
    On Error Goto 0
End Function

Function CSmartStr(Value)
    dim vt
	if isnull(Value) then
		CSmartStr=""
		exit function
	end if
	vt = vartype(Value)
	if vt=vbString then
		CSmartStr = Value
		exit function
	end if
	if vt=vbBoolean then
		if value then
			CSmartStr="-1"
		else
			CSmartStr=""
		end if
		exit function
	end if
	if vt=vbDate then
		CSmartStr = dbvalue(Value)
		exit function
	end if
    if vt=vbEmpty then _
        CSmartStr=""
    On Error Resume Next
    CSmartStr = CStr(Value)
    if Err.Number<>0 then
	    Err.Clear
	    CSmartStr =""
    end if
    On Error Goto 0
End Function

function CSmartDate(value)
	if isempty(value) or isnull(value) then
		CSmartDate=null
		exit function
	end if 
    On Error Resume Next
	CSmartDate=CDate(value)
    if Err.Number<>0 then
	    Err.Clear
	    CSmartDate =null
    end if
    if vartype(CSmartDate)=vbEmpty then _
        CSmartDate=null
    On Error Goto 0
end function

function db_pageseek(qhandle,pagesize,page)
	db_dataseek qhandle,(page-1)*pagesize
end function

function db_dataseek(qhandle,row)
   dim i
   i=0
   while i<row
   		qhandle.movenext
		i=i+1
   wend
end function

function asp_in_array(p_val,p_arr,p_strict)
	dim ret
	ret=asp_array_search(p_val, p_arr, p_strict)
	if IsIdentical(ret,false) then
		asp_in_array=false
	else
		asp_in_array=true
	end if
end function

function asp_array_search(p_val, p_arr, p_strict)
    dim key
    on error resume next
	for each key in p_arr.keys
        if not p_strict then
			if IsEqual(p_arr(key),p_val) then
    	        asp_array_search=key
        	    exit function
	        end if
		else
			if IsIdentical(p_arr(key),p_val) then
    	        asp_array_search=key
        	    exit function
	        end if
		end if
    next
	on error goto 0
    asp_array_search=false
end function

sub ReportError
if Err.number<>0 then
	response.flush
%>
</form>
<p align=center><font size=+2>ASP <%="error happened"%></font></p>
<table border="0" cellpadding="3" cellspacing="1" width="700" bgcolor="#000000" align="center">
<tr><td bgcolor="#ccccff" colspan=2 align=middle><font size=+1><b><%="Technical information" %></b></font></td></tr>
<tr bgcolor="#cccccc"><td width=130 bgcolor="#ccccff"><b>Error number</b></td><td align="left"><%=Err.Number%></td></tr>
<tr bgcolor="#cccccc"><td bgcolor="#ccccff"><b><%="Error description" %></b></td><td align="left"><font color=#cc3300><%=Err.Description%></font></td></tr>
<tr bgcolor="#cccccc"><td bgcolor="#ccccff"><b><%="URL" %></b></td><td align="left"><%=htmlspecialchars(Request.ServerVariables("URL"))%></td></tr>
<% if strSQL<>"" then %>
<tr bgcolor="#cccccc"><td bgcolor="#ccccff" ><b><%="SQL query" %></b></td><td align="left"><%=strSQL%></td></tr>
<% end if %>
<% if strMoreInfo<>"" then %>
<tr bgcolor="#cccccc"><td bgcolor="#ccccff" ><b>Additional info</b></td><td align="left"><%=strMoreInfo%></td></tr>
<% end if %>
</table>
    <form target=_new action="http://www.xlinesoft.com/asprunner/errors/default.asp" method="post" name="frmerror">
	<input type='hidden' name='ErrorNumber' value="<%=Err.Number%>" />
	<input type='hidden' name='Description' value="<%=Err.Description%>" />
	<input type='hidden' name='SQL' value="<%=dSQL%>" />
    </form>
<p align=center>
<a href="#" onClick="document.forms.frmerror.submit();return false;"><font size=3><b>More info on this error</b></font></a>
</p>

<%
	Response.End
end if
end sub

Function SafeURLEncode(str)
	if IsNull(str) or isEmpty(str) then
		SafeURLEncode=""
		exit function
	end if
    SafeURLEncode = replace(server.urlencode(CStr(str)),"+","%20")
End Function

Function htmlspecialchars(str)
    Dim ret,first,last
    if asp_is_array(str) and asp_count(str)>0 then 
		GetCollectionBounds str,first,last
		ret = CSmartStr(str(first))
    else
		ret=CSmartStr(str)
	end if	
    if len(ret)>0 then
		ret = Replace(ret, "&", "&amp;")
		ret = Replace(ret, """", "&quot;")
		ret = Replace(ret, "'", "&#039;")
		ret = Replace(ret, "<", "&lt;")
		ret = Replace(ret, ">", "&gt;")
	end if
    htmlspecialchars = ret
End Function

function asp_number_format(n,d,a,b)
    asp_number_format = FormatNumber(CSmartDbl(n),d,0,0,0)
end function

function asp_setcookie(name,val,ttime)
    Response.Cookies(name)=val
	Response.Cookies(name).Expires = DateAdd("yyyy", 1, Now())
end function

function isFalse(str)
	isFalse=false
	if vartype(str)=vbBoolean then 
		isFalse=not str
	end if
end function

function asp_join(term, arr)
    dim k, str
    str=""
    for each k in arr.keys
        str=str & arr(k) & term
    next
    if len(str)>0 then _
        str=left(str,len(str)-len(term))
    asp_join=str
end function

function GetUploadedFileContents(name)
	GetUploadedFileContents = GetRequestForm(name)
end function

function GetUploadedFileName(name)
	GetUploadedFileName = ""
end function

function asp_intval(val)
    asp_intval=CSmartLng(val)
end function    

function asp_array_splice(ByRef p_arr,offset,length)
    if offset>=0 and length>=0 then
        dim tmpDict, i, l
        l=0
        set tmpDict=CreateObject("Scripting.Dictionary")
        for each i in p_arr.keys
            if i<offset or i>=offset+length then
                setArrElement tmpDict,l,p_arr(i)
                l=l+1
            end if
        next
        set p_arr=tmpDict
    end if
end function


function DoUpdateRecord(byval table,byref evalues,byref blobfields,byval strWhereClause, byval pageid, byref pageObject)
	if SQLUpdateMode then 
		DoUpdateRecord = DoUpdateRecordSQL(table,evalues, blobfields,strWhereClause, pageid, pageObject)
		Exit Function
	end if 
	dim rs,strSQL,status
	Set rs = server.CreateObject("ADODB.Recordset")
	strSQL=gSQLWhere(strWhereClause,"")
	LogInfo(strSQL)


    on error resume next
    rs.Open strSQL, conn, 1,2
    call report_edit_error
	dim fields,keys,editformat,ftype,isAbs
	set keys = GetTableKeys("")
	fields=evalues.keys
	for each f in fields
		editformat = GetEditFormat(f,"")
		ftype=GetFieldType(f,"")
		if IsFalse(asp_array_search(f,keys,false)) or IsUpdatable(rs(f)) then
'	update field
			strValue = evalues.Item(f)
            if errorhappened then _
				exit for
            if isnull(strValue) then _
                strValue=""
   	        ctype = GetRequestForm("type_" & GoodFieldName(f) & "_" & pageid)

'            if editformat=EDIT_FORMAT_FILE then
'	            If ctype = "upload1" Then
'	            '	delete file
'		            rs(f) = Null
'					isAbs = GetFieldData("",f,"Absolute",false)
'	    	        set pageObject.filesToDelete(pageObject.filesToDelete.Count) = CreateClass("DeleteFile",3,GetRequestForm("filename_" & GoodFieldName(f) & "_" & pageid),GetUploadFolder(f,""),isAbs,Empty,Empty,Empty,Empty)
'	        	    if GetCreateThumbnail(f,"") then 
'						set pageObject.filesToDelete(pageObject.filesToDelete.Count) = CreateClass("DeleteFile",3,GetThumbnailPrefix(f,"") & GetRequestForm("filename_" & GoodFieldName(f) & "_" & pageid),GetUploadFolder(f,""),isAbs,Empty,Empty,Empty,Empty)
'					end if
'	            end if
'	            If ctype = "upload2" Then
    	        ' write file
'	    	        rs(f)= strValue
'	        	    if strValue<>"" then 
'						isAbs = GetFieldData("",f,"Absolute", false)
'						dim contents
'						contents = GetRequestForm("value_" & GoodFieldName(f) & "_" & pageid)
'						if ResizeOnUpload(f,"") then
'							contents = CreateThumbnail(contents,GetNewImageSize(f,""),CheckImageExtension(strValue))
'						end if
'						set pageObject.filesToSave(pageObject.filesToSave.Count) = CreateClass("SaveFile",4,contents, strValue,GetUploadFolder(f,""),isAbs,Empty,Empty,Empty)
'					end if
'	            end if
'           else
	            if IsBinaryType(ftype) then
					if len(strValue)=0 then
						rs(f).value=null
					else
		            rs(f).AppendChunk strValue
					end if
				elseif IsFloatType(ftype) then
		            if strValue<>"" then
			            rs(f) = CSmartDbl(strValue)
		            else
			            rs(f) = null
    		        end if
	            elseif IsNumberType(ftype) then
		            if strValue<>"" and IsNumeric(strValue) then 
			            rs(f) = CLng(strValue)
		            else
			            rs(f) = null
		            end if
	            elseif ischartype(ftype) then
					rs(f) = strValue			
				elseif IsDateFieldType(ftype) then
		            rs(f) = CSmartDate(strValue)
	            else
		            if strValue="" then
			            rs(f)=null
		            else
			            rs(f)=strValue
				    end if
			    end if
'		    end if
		    call report_edit_error
		end if
	next	
	if not errorhappened then
		rs.Update
		call report_edit_error
	end if
	rs.Close
	if errorhappened then
		DoUpdateRecord=false
		exit function
	end if
'	save files
	pageObject.ProcessFiles
	if inlineedit then
		status="UPDATED"
		message="" & "Record updated" & ""
		IsSaved = true
	else
		message="<div class=message><<< " & "Record updated" & " >>></div>"
	end if
	if usermessage<>"" then _
		message=usermessage
	DoUpdateRecord=true
end function


function DoInsertRecord(byval table,byref avalues,byref blobfields, byval pageid, byref pageObject)
	if SQLUpdateMode then 
		DoInsertRecord = DoInsertRecordSQL(table,avalues, blobfields, pageid, pageObject)
		Exit Function
	end if 
	dim rs,status
	Set rs = server.CreateObject("ADODB.Recordset")
		on error resume next
	rs.Open "select * from " & AddTableWrappers(table) & " where 1=0", conn, 1,2
	rs.Addnew
	call report_add_error

	dim fields,tkeys,editformat,ftype
	set tkeys = GetTableKeys("")
	fields=avalues.keys
	for each f in fields
		if errorhappened then _
			exit for
		editformat = GetEditFormat(f,"")
		ftype=GetFieldType(f,"")
		if IsFalse(asp_array_search(f,tkeys,false)) or IsUpdatable(rs(f)) then
'	insert field
			strValue = avalues.Item(f)
			if isnull(strValue) then _
				strValue=""
'			ctype = GetRequestForm("type_" & GoodFieldName(f) & "_" & pageid)
'			if editformat=EDIT_FORMAT_FILE then
'				If ctype = "upload2" Then
' write file
'					rs(f)= strValue
'				end if
'			else
				if IsBinaryType(ftype) then
					if len(strValue)=0 then
						rs(f).value=null
					else
					rs(f).AppendChunk strValue
					end if
				elseif IsFloatType(ftype) then
					if strValue<>"" then
						rs(f) = CSmartDbl(strValue)
					else
						rs(f) = null
					end if
				elseif IsNumberType(ftype) then
					if strValue<>"" and IsNumeric(strValue) then 
						rs(f) = CLng(strValue)
					else
						rs(f) = null
					end if
				elseif ischartype(ftype) then
					rs(f) = strValue			
				elseif IsDateFieldType(ftype) then
		            rs(f) = CSmartDate(strValue)
				else
					if strValue="" then
						rs(f)=null
					else
						rs(f)=strValue
					end if
			    end if
'			end if
			call report_add_error
		end if
	next
	if errorhappened then
		DoInsertRecord=false
		exit function
	end if
	rs.Update
	call report_add_error
	on error goto 0
	if errorhappened then
		DoInsertRecord=false
		exit function
	end if
'	save files
	pageObject.ProcessFiles
	if inlineadd=ADD_INLINE then
		status="ADDED"
		message="" & "Record was added" & ""
		IsSaved = true
	else
		message="<div class=message><<< " & "Record was added" & " >>></div>"
	end if
	if usermessage<>"" then _
		message=usermessage

'	get new key values
	failed_inline_add = false
	dim kk,k
	kk=tkeys.keys
	for each k in kk 
		keys(tkeys(k)) = dbvalue(rs(tkeys(k)))
	next
	rs.Close
	DoInsertRecord=true
end function


Function IsUpdatable(Field)
		if Field.Attributes and 4 or Field.Attributes and 8 then
			IsUpdatable=true
		else
			IsUpdatable=false
		end if		
End Function


function dbvalue(value)
	if isnull(value) then
		dbvalue=""
		exit function
	end if
	if vartype(value)=7 then
		dbvalue=year(value) & "-" & month(value) & "-" & day(value) & " " & hour(value) & ":" & minute(value) & ":" & second(value)
		exit function
	end if
	dbvalue=value
	exit function
end function

function GetCurrentYear()
	GetCurrentYear=year(now)
end function


Function ParseMultiPartForm
	if Request.TotalBytes = 0 then
		ParseMultiPartForm = false
		Exit Function
	end if
	ParseMultiPartForm = true
	Dim postData
	postData = Request.BinaryRead(Request.TotalBytes)
	contentType = Request.ServerVariables( "HTTP_CONTENT_TYPE")
	ctArray = split( contentType, ";")
	if trim(ctArray(0)) = "multipart/form-data" then
		errMsg = ""
		' grab the form boundry...
		bArray = split( trim( ctArray(1)), "=")
		boundry = Unicode2Bytes("--" & trim( bArray(1)))
		currentPos = 1
		inStrByte = 1
		While inStrByte > 0
			inStrByte = InStrB(currentPos, postData, boundry)
			m = inStrByte - currentPos
			If m > 1 Then
    			val = MidB(postData, currentPos, m)
    			
        		infoEnd = instrB( val, chrb(13) & chrb(10) & chrb(13) & chrb(10) )
        		if infoEnd > 0 then
        			varInfo = Bytes2String(midb( val , 1, infoEnd - 1))
        			varValue = midb( val , infoEnd + 4, lenb(val) - infoEnd - 5)

					if InStr(1, varInfo, "Content-Type") < 1 then 
						varValue=Bytes2String(varValue)
					else
						if lenb(varValue) mod 2 then varValue = varValue & chrb(0)
					end if

				strField = getFieldName(varInfo)
				if myRequest.exists(strField) then
					myRequest(strField) = myRequest(strField) & "," & varValue
        			else
					myRequest.add strField, varValue
				end if
				
        		end if
			end if
			currentPos = lenb(boundry) + inStrByte
		wend
	else
		errMsg = "Wrong encoding type!"
	end if 

End Function

Function Bytes2String(bytes)
	Dim i, byteord, nextbyteord
	For i = 1 to LenB(bytes)
		byteord = AscB(MidB(bytes, i, 1))
		If session.codepage<>65001 or byteord < &H80 Then ' Ascii
			Bytes2String= Bytes2String& Chr(byteord)
		Else ' Double-byte characters?
			if byteord >= &HC2 and byteord <= &HDF and i < LenB(bytes) then
				byteord2 = AscB(MidB(bytes, i+1, 1))
				On Error Resume Next
				charindex = (byteord-192)*64 + (byteord2-128)
				Bytes2String= Bytes2String& ChrW(charindex)
				If Err.Number <> 0 Then
					On Error GoTo 0
					Bytes2String= Bytes2String& Chr(byteord) & Chr(byteord2)
				End If
				i = i + 1
			elseif byteord >= 112 and byteord < 240 and i+1 < LenB(bytes) then
				byteord2 = AscB(MidB(bytes, i+1, 1))
				byteord3 = AscB(MidB(bytes, i+2, 1))
				On Error Resume Next
				charindex = (byteord-224)*4096 + (byteord2-128)*64 + (byteord3-128)
				Bytes2String= Bytes2String& ChrW(charindex)
				If Err.Number <> 0 Then
					On Error GoTo 0
					Bytes2String= Bytes2String& Chr(byteord) & Chr(byteord2) & Chr(byteord3)
				End If
				i = i + 2
			elseif i+2 < LenB(bytes) then
				byteord2 = AscB(MidB(bytes, i+1, 1))
				byteord3 = AscB(MidB(bytes, i+2, 1))
				byteord4 = AscB(MidB(bytes, i+3, 1))
				On Error Resume Next
				charindex = (byteord-240)*262144 + (byteord2-128)*4096 + (byteord3-128)*64 + (byteord4-128)
				Bytes2String= Bytes2String& ChrW(charindex)
				If Err.Number <> 0 Then
					On Error GoTo 0
					Bytes2String= Bytes2String& Chr(byteord) & Chr(byteord2) & Chr(byteord3) & Chr(byteord4)
				End If
				i = i + 3
			 Else 
			   Bytes2String= Bytes2String& Chr(byteord)
			 end if
		End If
	Next
End Function

function getFieldName( infoStr)
	sPos = inStr( infoStr, "name=")
	endPos = inStr( sPos + 6, infoStr, chr(34) & ";")
	if endPos = 0 then
		endPos = inStr( sPos + 6, infoStr, chr(34))
	end if
	getFieldName = mid( infoStr, sPos + 6, endPos - (sPos + 6))
end function

' This function retreives a file field's filename
function getFileName( infoStr)
	sPos = inStr( infoStr, "filename=")
	endPos = inStr( infoStr, chr(34) & crlf)
	getFileName = mid( infoStr, sPos + 10, endPos - (sPos + 10))
end function

' This function retreives a file field's mime type
function getFileType( infoStr)
	sPos = inStr( infoStr, "Content-Type: ")
	getFileType = mid( infoStr, sPos + 14)
end function

Function GetRequestForm(key)
    if isEmpty(myRequest) then
	    GetRequestForm=""
	    Exit Function
    end if

    if myRequest.Exists(key) then
	    GetRequestForm = myRequest(key)
    else
	    GetRequestForm = Request.QueryString(key)
    end if
End Function

Function Unicode2Bytes(str)
    For ind = 1 To len(str) 
         Unicode2Bytes = Unicode2Bytes& ChrB(Asc(Mid(str, ind, 1)))
    Next 
End Function


function prepare_file(value,field,controltype,postfilename,id)
		if (trim(value)="" or isnull(value)) and mid(controltype,1,5)<>"file1" then
			prepare_file=false
		else
			prepare_file=value
		end if
		if trim(postfilename)<>  "" then _
			filename=trim(postfilename)
end function
function prepare_upload(field,controltype,postfilename,value,table,id, byref pageObject)
		if controltype="upload0" then 
			prepare_upload = false
			exit function
		end if
		prepare_upload = value
		If controltype = "upload1" Then
	'	delete file
			isAbs = GetFieldData("",field,"Absolute",false)
			set pageObject.filesToDelete(pageObject.filesToDelete.Count) = CreateClass("DeleteFile",3,postfilename,GetUploadFolder(field,""),isAbs,Empty,Empty,Empty,Empty)
			if GetCreateThumbnail(field,"") then 
				set pageObject.filesToDelete(pageObject.filesToDelete.Count) = CreateClass("DeleteFile",3,GetThumbnailPrefix(field,"") & postfilename,GetUploadFolder(field,""),isAbs,Empty,Empty,Empty,Empty)
			end if
		end if
		If controltype = "upload2" Then
	' write file
			if value<>"" then 
				isAbs = GetFieldData("",field,"Absolute", false)
				dim contents
				contents = GetRequestForm("value_" & GoodFieldName(field) & "_" & id)
				if ResizeOnUpload(field,"") then
					contents = CreateThumbnail(contents,GetNewImageSize(field,""),CheckImageExtension(value))
				end if
				set pageObject.filesToSave(pageObject.filesToSave.Count) = CreateClass("SaveFile",4,contents, value,GetUploadFolder(field,""),isAbs,Empty,Empty,Empty)
			end if
		end if
end function

function FieldSubmitted(field)
	FieldSubmitted = myRequest.Exists("value_" & GoodFieldName(field)) or myRequest.Exists("value_" & GoodFieldName(field) & "[]") or myRequest.Exists("type_" & GoodFieldName(field))
end function
		   
sub report_edit_error
	if Err.number<>0 then
	    if inlineedit then
		    message ="" & "Record was NOT edited" & ". " & Err.Description
	    else  
	      message = "<div class=message><<< " & "Record was NOT edited" & " >>><br><br>" & Err.Description & "</div>"
	    end if
	  readevalues=true
	  errorhappened=true
	  err.clear
	end if
end sub

function asp_implode(p_str, p_arr)
    dim str
    str=""
    for each v in p_arr.keys
        str=str & p_arr(v) & p_str
    next
    if len(str)>0 then _
        str=left(str,len(str)-len(p_str))
    asp_implode=str
end function

function asp_strcmp(str1, str2)
    if str1<str2 then
        asp_strcmp=-1
    elseif str1>str2 then
        asp_strcmp=1
    else
        asp_strcmp=0
    end if
end function

function copyDictionary(byref from,byref out)
    dim k, d,ks,td
    Set d = CreateDictionary()
    ks=from.keys
    for each k in ks
		if IsDictionary(from(k)) then 
            copyDictionary from(k),td
            set d(k)=td
        elseif Isobject(from(k)) then
			set d(k)=from(k)
		else
            d(k)=from(k)
        end if
    next
    set out=d
end function

sub OrderTables(ByRef tables)
' order tables by tables(i)(0)
end sub

function getReportArray(name)
	getReportArray=CreateDictionary()
end function

function getChartArray(name)
	getChartArray=CreateDictionary()
end function

function asp_sort(arr)
    dim i, j, elem
	dim keys
	dim out
	set out = CreateDictionary()
	keys = arr.keys

	if asp_count(arr)<2 then exit function
    for each i in keys
        for each j in keys
			if j>i then
	            if arr(j)<arr(i) then
    	            elem=arr(i)
					doArrayAssignment arr,i,arr(j)
					doArrayAssignment arr,j,elem
		   		 end if
			end if
        next
    next

	set out = CreateDictionary()
	for each i in keys
		doArrayAssignment out, out.count, arr(i)
	next
	set arr = out
end function

Function postvalue(ByVal name)
	Dim var_POST,value,var_GET,ret,key,val
	if bValue(asp_array_key_exists(name & "[]",RequestForm())) then
		doAssignment value,GetRequestValue(RequestForm(),name)
	elseif bValue(asp_array_key_exists(name,RequestForm())) then
		doAssignment value,GetRequestValue(RequestForm(),name)
	elseif bValue(asp_array_key_exists(name & "[]",Request.QueryString)) then
		doAssignment value,GetRequestValue(Request.QueryString,name)
	elseif bValue(asp_array_key_exists(name,Request.QueryString)) then
		doAssignment value,GetRequestValue(Request.QueryString,name)
	else
		postvalue = ""
		Exit Function
	end if
	if isempty(value) then
		value=""
	end if
	if not bValue(asp_is_array(value)) then
		doAssignment postvalue,value
		Exit Function
	end if
	Set ret = CreateDictionary()
	GetCollectionBounds value,loopIdx157,loopMax157
	do while loopIdx157<=loopMax157
		key = GetCollectionKey(value,loopIdx157)
		doAssignment val,value(key)
		setArrElement ret,key,val
		loopIdx157=loopIdx157+1
	loop
	doAssignment postvalue,ret
	Exit Function
End Function


' Functions to provide encoding/decoding of strings with Base64.
' 
' Encoding: myEncodedString = base64_encode( inputString )
' Decoding: myDecodedString = base64_decode( encodedInputString )
'
' Programmed by Markus Hartsmar for ShameDesigns in 2002. 
' Email me at: mark@shamedesigns.com
' Visit our website at: http://www.shamedesigns.com/
'

	Dim Base64Chars
	Base64Chars =	"ABCDEFGHIJKLMNOPQRSTUVWXYZ" & _
			"abcdefghijklmnopqrstuvwxyz" & _
			"0123456789" & _
			"+/"


	' Functions for encoding string to Base64
	Public Function base64_encode( byVal strIn )
		Dim c1, c2, c3, w1, w2, w3, w4, n, strOut
		For n = 1 To Len( strIn ) Step 3
			c1 = Asc( Mid( strIn, n, 1 ) )
			c2 = Asc( Mid( strIn, n + 1, 1 ) + Chr(0) )
			c3 = Asc( Mid( strIn, n + 2, 1 ) + Chr(0) )
			w1 = Int( c1 / 4 ) : w2 = ( c1 And 3 ) * 16 + Int( c2 / 16 )
			If Len( strIn ) >= n + 1 Then 
				w3 = ( c2 And 15 ) * 4 + Int( c3 / 64 ) 
			Else 
				w3 = -1
			End If
			If Len( strIn ) >= n + 2 Then 
				w4 = c3 And 63 
			Else 
				w4 = -1
			End If
			strOut = strOut + mimeencode( w1 ) + mimeencode( w2 ) + _
					  mimeencode( w3 ) + mimeencode( w4 )
		Next
		base64_encode = strOut
	End Function

	Private Function mimeencode( byVal intIn )
		If intIn >= 0 Then 
			mimeencode = Mid( Base64Chars, intIn + 1, 1 ) 
		Else 
			mimeencode = ""
		End If
	End Function	


	' Function to decode string from Base64
	Public Function base64_decode( byVal strIn )
		Dim w1, w2, w3, w4, n, strOut
		For n = 1 To Len( strIn ) Step 4
			w1 = mimedecode( Mid( strIn, n, 1 ) )
			w2 = mimedecode( Mid( strIn, n + 1, 1 ) )
			w3 = mimedecode( Mid( strIn, n + 2, 1 ) )
			w4 = mimedecode( Mid( strIn, n + 3, 1 ) )
			If w2 >= 0 Then _
				strOut = strOut + _
					Chr( ( ( w1 * 4 + Int( w2 / 16 ) ) And 255 ) )
			If w3 >= 0 Then _
				strOut = strOut + _
					Chr( ( ( w2 * 16 + Int( w3 / 4 ) ) And 255 ) )
			If w4 >= 0 Then _
				strOut = strOut + _
					Chr( ( ( w3 * 64 + w4 ) And 255 ) )
		Next
		base64_decode = strOut
	End Function

	Private Function mimedecode( byVal strIn )
		If Len( strIn ) = 0 Then 
			mimedecode = -1 : Exit Function
		Else
			mimedecode = InStr( Base64Chars, strIn ) - 1
		End If
	End Function

function sortTables(ByRef tables)
end function

	
function sortMembers(ByRef rowinfo)
    gcount=rowinfo.count
	dim i
	dim keys
	dim gi
	dim tmp
	dim mindex,gcount
    set tmp = CreateObject("Scripting.Dictionary")
	if rowinfo.count=0 then exit function
'	find group index
	gcount=rowinfo(0)("usergroup_boxes")("data").Count
	gi=gcount
	for i=0 to gcount-1
		if clng(rowinfo(0)("usergroup_boxes")("data")(i)("group"))=cSmartlng(sortgroup) then
			gi = i
			exit for
		end if
	next
'	run sorting
	do while rowinfo.count>0
'	init values
        keys=rowinfo.keys
        mindex=keys(0)
        for each i in keys
            if i<>keys(0) then       
                if gi=gcount or ArrayElement(rowinfo(i)("usergroup_boxes")("data")(gi),"checked")=ArrayElement(rowinfo(mindex)("usergroup_boxes")("data")(gi),"checked") then
                    if rowinfo(i)("user")<rowinfo(mindex)("user") then _
                        mindex=i
                else
                    
                    if sortorder="a" and not bValue(ArrayElement(rowinfo(i)("usergroup_boxes")("data")(gi),"checked")) then _
                        mindex=i
                    if sortorder="d" and not bValue(ArrayElement(rowinfo(mindex)("usergroup_boxes")("data")(gi),"checked")) then _
                        mindex=i
                end if
            end if
		next
		tmp.Add tmp.Count, rowinfo(mindex)
		rowinfo.remove mindex
	loop
	set rowinfo=tmp
end function

function report_add_error
	if Err.number<>0 then
	    if inlineadd=ADD_INLINE then
		    message ="" & "Record was NOT added" & ". " & Err.Description
	    else  
		    message = "<div class=message><<< " & "Record was NOT added" & " >>><br><br>" & Err.Description & "</div>"
	    end if
	  readavalues=true
	  errorhappened=true
	  err.clear
	end if
end function

function IsEqual(a1,a2)
	dim p1,p2
	doAssignment p1,a1
	doAssignment p2,a2
	if vartype(p1)=vbNull or vartype(p1)=vbEmpty then
		IsEqual = CSmartStr(p1)=CSmartStr(p2)
		exit function
	end if
	if vartype(p1) = vartype(p2) then
		IsEqual=p1=p2
		exit function
	end if
	if vartype(p1)=vbBool then
		p2=bValue(p2)
	elseif vartype(p2)=vbBool then
		p1=bValue(p2)
	elseif vartype(p1)=vbInteger or vartype(p1)=vbLong or vartype(p1)=vbByte then
		p2=CSmartLng(p2)
	elseif vartype(p2)=vbInteger or vartype(p2)=vbLong or vartype(p2)=vbByte then
		p1=CSmartLng(p1)
	elseif vartype(p1)=vbSingle or vartype(p1)=vbDouble or vartype(p1)=vbCurrency then
		p2=CSmartDbl(p2)
	elseif vartype(p2)=vbSingle or vartype(p2)=vbDouble or vartype(p2)=vbCurrency then
		p1=CSmartDbl(p1)
	else
		p1=CSmartStr(p1)
		p2=CSmartStr(p2)
	end if
	IsEqual=p1=p2
end function

function IsIdentical(a1,a2)
	if vartype(a1)<>vartype(a2) then
		IsIdentical=false
	else
		IsIdentical=IsEqual(a1,a2)
	end if
end function


sub print_r(ByRef var)
	print_r_int var,0
end sub

sub print_r_int(ByRef var,ByVal indent)
	if not isObject(var) then
		if vartype(var)=vbBoolean then
			if var then
				responsewrite 1
			end if
		else
		ResponseWrite var
		end if
	elseif IsDictionary(var) then
		ResponseWrite "Array"
		if var.count=0 then
			ResponseWrite "()"
			exit sub
		end if
		ResponseWrite vbcrlf & space(indent) & "(" & vbcrlf
		indent=indent+4
		dim keys
		keys=var.keys
		for each k in keys
			ResponseWrite space(indent) & "["&k&"] => "
			print_r_int var(k),indent+4
			ResponseWrite vbcrlf
		next
		indent=indent-4
		ResponseWrite space(indent) & ")" & vbcrlf
	else
		ResponseWrite "["& typename(var) & "]"
	end if
end sub


Function CustomExpression(ByVal strValue,ByRef data,ByVal field,ByVal table)
	if not bValue(table) then
		doAssignment table,strTableName
	end if
	dim rs
	set rs=data
	doAssignment CustomExpression,strValue
	Exit Function
End Function

function asp_array_unshift(byref arr, byref var)
    dim tmpDict, i, a
    set tmpDict = CreateObject("Scripting.Dictionary")    
	setArrElementByRef tmpDict,0,var
    i=1
    for each a in arr.keys
        setArrElement tmpDict,i,arr(a)
        i=i+1
    next
    doAssignmentByRef arr,tmpDict
End Function

function asp_array_shift(byref arr)
    dim i
    set tmpDict = CreateObject("Scripting.Dictionary")    
    if arr.Count=0 then
        asp_array_shift=null
        exit function
    end if
    doAssignmentByRef asp_array_shift,arr(0)
    for i=0 to arr.count-2
        setArrElement arr,i,ArrayElement(arr,i+1)
    next
    arr.Remove(arr.count-1)
End Function

function rand(vmin, vmax)
    randomize
    rand=rnd(1)*(vmax-vmin)+vmin
end function

sub runner_save_file(strFileName, binData)
	Dim rsT
	Set rsT = Server.CreateObject("ADODB.Recordset")
	rsT.Fields.Append "File", 205, LenB(binData)
	rsT.Open
	rsT.AddNew
	rsT.Fields("File").AppendChunk binData
	rsT.Fields("File").AppendChunk "0"
	rsT.Update
	
	Dim stream
	Set stream = Server.CreateObject("ADODB.Stream")
	stream.Type = 1
	stream.Open
	stream.Write rsT.Fields("File").GetChunk(LenB(binData))
	stream.SaveToFile strFileName, 2

	stream.Close
	Set stream = Nothing
	rsT.Close
	Set rsT = Nothing
end sub
'//	return lookup wizard WHERE expression

function LookupWhere(ByVal field,ByVal table)
	if not bValue(table) then
		doAssignment table,strTableName
	end if
		
	LookupWhere = ""
end function

function GetDefaultValue(ByVal field,ByVal table)
	if not bValue(table) then
		doAssignment table,strTableName
	end if
		
	GetDefaultValue = ""
end function


function mdeleteIndex(i)
    mdeleteIndex=i
end function

function InArray(arr,val)
    dim i
    for i=0 to asp_count(arr)
        if arr(i)=val then
            InArray=true
            exit function
        end if
    next
    InArray=false
end function

function getabspath(filename)
	if isabspath(filename) then
		getabspath = filename
	else
	getabspath=Server.MapPath(filename)
	end if
end function

function GetMySQL4RowCount(countstr)
    dim asp_rs
    Set asp_rs = server.CreateObject("ADODB.Recordset")
    asp_rs.Open sSQL,conn,3,1
    call ReportError
	GetMySQL4RowCount = asp_rs.RecordCount
end function


function serialize(ByRef obj)
	dim arr
	set arr=obj.ASPserialize()
	arr("ASPclassname")=typename(obj)
	set serialize=arr
end function

function unserialize(ByRef arr)
	dim str
	str="set unserialize=new " & arr("ASPclassname")
	Execute str
	unserialize.ASPunserialize(arr)
end function

function asp_array_slice(ByRef arr, idx,length)
	dim out,i,keys
    if idx<0 then
        idx=asp_count(arr)+idx
    end if
    if length<0 then
        length=asp_count(arr)+length
    end if
	keys=arr.keys
	set out=CreateDictionary()
	i=0
    while (i<length or isempty(length)) and idx+i<asp_count(arr)
        setArrElement out,keys(i),ArrayElement(arr,idx+i)
		i=i+1
	wend
	set asp_array_slice=out
end function

function asp_array_intersect(ByRef arr1, ByRef arr2)
	set out=CreateDictionary()
	dim keys1,keys2,i1,i2
	keys1=arr1.keys
	keys2=arr2.keys
	for each i1 in keys1
		for each i2 in keys2
			if IsEqual(arr1(i1),arr2(i2)) then
				setArrElement out,i1,arr1(i1)
			end if
		next
	next
	set asp_array_intersect=out
end function

function strcasecmp(ByVal str1,ByVal str2)
	strcasecmp = asp_strcmp(ucase(str1),ucase(str2))
end function

function no_output_done()
	err.clear
	on error resume next
	response.addheader "TestHeader","test"
	if err.number<>0 then
		err.clear
		no_output_done=false
	else
		no_output_done=true
	end if
	on error goto 0
end function

sub flush_output()
	if response.buffer then _
		response.flush
end sub

function basename(str)
	basename=str
end function 

function fformat_number(val)
	fformat_number = str_format_number(val)
end function

function fformat_currency(val)
	fformat_currency = str_format_currency(val)
end function

function format_datetime(ttime())
	format_datetime = str_format_datetime(ttime)
end function

function fformat_time(ttime())
	fformat_time = str_format_time(ttime)
end function

Sub DoEvent(strEvent)
	On Error Resume Next
	Execute strEvent
	If Err.Number <> 13 Then
        strMoreInfo = "Event: " & strEvent
        ReportError
	End If
	On Error GoTo 0
End Sub

function IsDictionary(byref p_arr)
	if not isobject(p_arr) then
		IsDictionary=false
		exit  function
	end if
	on error resume next
	dim ret
	ret=p_arr.CompareMode
	if err.Number=0 then
		IsDictionary=true
	else
		IsDictionary=false
	end if
	on error goto 0
end function

function ArrayElement(byref p_arr, byval key)
	dim status
	if not IsObject(p_arr) and vartype(p_arr)<>vbString then
		if isArray(p_arr) then
			ArrayElement=p_arr(key)
			exit function
		end if
	end if
	if vartype(key)=vbString then
		if IsNumeric(key) then
			key=CLng(key)
		end if
	end if
 	if not isobject(p_arr) and vartype(p_arr)=vbString  then
 		ArrayElement=mid(p_arr,key+1,1)
		exit function
	end if
	on error resume next
 	if p_arr.Exists(key) then
 		DoAssignmentByRef ArrayElement,p_arr(key)
	else
		ArrayElement=Empty
	end if
	if err.number<>0 then
		ArrayElement=Empty
	end if
end function

function asp_array_reverse(arr)
    dim tmpDict,key,j
    set tmpDict = CreateObject("Scripting.Dictionary")
    set key = CreateObject("Scripting.Dictionary")
    key = arr.keys
	j=0
    for i=ubound(key) to 0 step -1
	    setArrElement tmpDict,j,arr(key(i))
		j=j+1
    next
    set asp_array_reverse=tmpDict
end function

function asp_shl(a, n)
	asp_shl=CLng(a)
	for i=1 to n 
		asp_shl=asp_shl*2
	next
end function

function asp_shl(a, n)
	asp_shl=CLng(a)
	for i=1 to n 
		asp_shl=Int(asp_shl/2)
	next
end function



function asp_array_diff(arr1,arr2)
    dim tmpDict, i, key
    set tmpDict = CreateObject("Scripting.Dictionary")
    for each key in arr1.keys
	    if not asp_in_array(arr1(key),arr2,true) then 
			tmpDict(key)=arr1(key)
		end if
    next
    set asp_array_diff=tmpDict
end function

function asp_array_values(arr)
    dim tmpDict, i, key
    set tmpDict = CreateObject("Scripting.Dictionary")
    i=0
    for each key in arr.keys
       setArrElement tmpDict,i,arr(key) 
       i=i+1
    next
    set asp_array_values=tmpDict
end function

function asp_array_merge(arr1,arr2)
    dim tmpDict, i, key
    set tmpDict = CreateObject("Scripting.Dictionary")
    i=0
    for each key in arr1.keys
	    setArrElement tmpDict,i,arr1(key)
	    i=i+1
    next
    for each key in arr2.keys
	    setArrElement tmpDict,i,arr2(key)
	    i=i+1
    next

    set asp_array_merge=tmpDict
end function

function asp_preg_match(patt,strng,arr)
    Dim regEx, Match, Matches, i, sm, p, result, modif

    if left(patt,1)="/" then patt=mid(patt,2)
    p=instrrev(patt,"/")
    if p>0 then
        modif=mid(patt,p+1)
        patt=left(patt,p-1)
    end if

    Set regEx = New RegExp

    if instr(1,modif,"i") then
        regEx.IgnoreCase = True
    else
        regEx.IgnoreCase = False
    end if

    result=patt
    
    if instr(1,modif,"U") then
        result=""
        for i=1 to len(patt)
	    if mid(patt,i,1)="*" or mid(patt,i,1)="+" then
	        if i>1 then
		    if mid(patt,i-1,1)="\" then 
		        result=result & mid(patt,i,1)
		    else
		        if i<len(patt) then
			    if mid(patt,i+1,1)="?" then 
		                result=result & mid(patt,i,1)
			        i=i+1
			    else
			        result=result & mid(patt,i,1) & "?"
			    end if
		        else
		            result=result & mid(patt,i,1) & "?"
		        end if
	            end if
	        else
		    if mid(patt,i+1,1)="?" then 
	                result=result & mid(patt,i,1)
		        i=i+1
		    else
		        result=result & mid(patt,i,1) & "?"
		    end if
	        end if
	    else
	        result=result & mid(patt,i,1)
	    end if
        next
    end if

    regEx.Pattern = result
    regEx.Global = True

    Set Matches = regEx.Execute(strng)

    if not isnull(arr) then
        set arr = CreateObject("Scripting.Dictionary")
        if Matches.Count>0 then
            Set Match = Matches(0)
            arr(0)=Match.Value
            i=1
            for each sm in Match.SubMatches
	            arr(i)=sm
	            i=i+1
            Next
        end if
    end if

    if Matches.Count>0 then
        asp_preg_match=1
    else
        asp_preg_match=0
    end if
    
end function

Function asp_preg_replace(patt,repl,strng)
    Dim regEx,p
    Set regEx = New RegExp
    
    if left(patt,1)="/" then patt=mid(patt,2)
    p=instrrev(patt,"/")
    if p>0 then
        modif=mid(patt,p+1)
        patt=left(patt,p-1)
    end if

    Set regEx = New RegExp

    if instr(1,modif,"i") then
        regEx.IgnoreCase = True
    else
        regEx.IgnoreCase = False
    end if
    
    regEx.Pattern = patt
    
    asp_preg_replace=regEx.Replace(strng, repl)
End Function

function asp_include(byval filename,once)
		filename=getabspath(filename)
	if once then
		if included_files.exists(filename) then
			exit function
		end if
	end if
	included_files(filename)=true
	ExecuteGlobal readIncludeFile(filename)
end function


function readIncludeFile(filename)
	Dim stream,out,pos,start,pos1,start1,textblock,txt,incfile,path
	pos=instrrev(filename,"\")
	path=left(filename,pos)

	set stream=Server.CreateObject("ADODB.Stream")
	stream.CharSet=cCharset
	stream.type=2
	stream.Open
	stream.LoadFromFile Filename

	file = stream.ReadText
	stream.Close
	set stream=nothing
'	cut asp wrappers and include files
	out=""
	start=1
	do while start<=len(file)
		pos=instr(start,file,"<%")

'	add text contents
		if pos=0 or pos>start then
'	handle file includes
			if pos=0 then
				textblock=mid(file,start)
			else
				textblock=mid(file,start,pos-start)
			end if
			start1=1
			do while start1<len(textblock)
				pos1=instr(start1,textblock,"<!--#incl"&"ude file=""")
'	add plain text contents
				if pos1=0 or pos1>start then
					if pos1>0 then
						txt = mid(textblock,start1,pos1-start1)
					else
						txt = mid(textblock,start1)
					end if
				end if
				txt=trim(replace(replace(txt,vbcr,""),vblf,""))
				if len(txt) then
					out = out & "ResponseWrite """ 
					out = out & replace(txt,"""","""""")
					out = out & """" & vbcrlf
				end if
				if pos1=0 then
					exit do
				end if
				start1=pos1+len("<!--#incl"&"ude file=""")
				pos1=instr(start1,textblock,"""-->")
				if pos1=0 then
					pos1=len(textblock)
				end if
'	do include
				incfile=path & mid(textblock,start1,pos1-start1)
				out=out&"asp_include """ & replace(incfile,"""","""""") & """,false" & vbcrlf
				start1=pos1+len("""-->")
			loop
		end if
		if pos=0 then 
			exit do
		end if
'	add code block
		start=pos+2
		pos=instr(start,file,"%" & ">")
		if pos=0 then 
			pos=len(file)
		end if
		out=out & vbcrlf & mid(file,start,pos-start)
		start=pos+2
	loop
	readIncludeFile=out
end function

' old style array functions

function doArrayAssignment(ByRef arr,ByRef key,ByRef value)
	dim tval
	doAssignment tval,value
	if not IsObject(value) then
		arr(key)=tval
		doArrayAssignment=tval
	else
		set arr(key)=tval
		doArrayAssignment=bValue(tval)
	end if
end function


function doArrayAssignmentByRef(ByRef arr,ByRef key,ByRef value)
	if not IsObject(value) then
		arr(key)=value
		doArrayAssignmentByRef=value
	else
		set arr(key)=value
		doArrayAssignmentByRef=bValue(value)
	end if
end function


function doArrayInArrayAssignment(ByRef arr,ByRef key,ByVal key1,ByRef value)
	ensureArrayCreated arr,key
	if isEmpty(key1) then
		key1=asp_count(arr(key))
	end if
	doArrayAssignment arr(key),key1,value
	doAssignmentByRef doArrayInArrayAssignment,value
end function

function doArrayInArrayAssignmentByRef(ByRef arr,ByRef key,ByVal key1,ByRef value)
	ensureArrayCreated arr,key
	if isEmpty(key1) then
		key1=asp_count(arr(key))
	end if
	doArrayAssignmentByRef arr(key),key1,value
	doAssignmentByRef doArrayInArrayAssignmentByRef,value
end function
' end old style

function isLess(byval arg1,byval arg2)
	if IsEmpty(arg2) or IsNull(arg2) then
		isLess=false
		exit function
	end if
	if IsEmpty(arg1) or IsNull(arg1) then
		isLess=true
		exit function
	end if
	if IsNumeric(arg1) and IsNumeric(arg2) then
		isLess=CDbl(arg1)<CDbl(arg2)
		exit function
	end if
	isLess=arg1<arg2
end function

function isLessOrEqual(arg1,arg2)
	if isEqual(arg1,arg2) then
		isLessOrEqual=true
		exit function
	end if
	if isLess(arg1,arg2) then
		isLessOrEqual=true
		exit function
	end if
	isLessOrEqual=false
end function

function callVariableMethod(byref object,byval method,byref params)
	dim command,i,strParams
	command="DoAssignmentByRef callVariableMethod,object." & method 
	if not isEmpty(params) then
		command=command & "_p" & params.Count & "("
		for i=0 to params.Count-1
			if i>0 then 
				command=command & ","
			end if
			command = command & "params("&i&")"
		next
		command=command & ")"
	end if
	Execute command
end function


function db_query_safe(sSQL,conn,byref errstr)
    dim asp_rs
    Set asp_rs = server.CreateObject("ADODB.Recordset")
    err.clear
	on error resume next
	asp_rs.Open sSQL,conn
	errstr=err.description
	if err.number=0 then
		set db_query_safe=asp_rs
	else
		set db_query_safe=false
	end if
	on error goto 0
end function

function binPrint(byref value, size)
	response.BinaryWrite value
end function

function DisplayNoImage()
	Response.ContentType = "image/gif"
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set a = fs.GetFile(Server.MapPath("images/no_image.gif"))
	Set b = a.OpenAsTextStream(1,-1)
	Response.BinaryWrite(b.Read(999999))
end function
Sub DisplayFileImage
	Response.ContentType = "image/gif"
	Set fs = CreateObject("Scripting.FileSystemObject")
	Set a = fs.GetFile(Server.MapPath("images/file.gif"))
	Set b = a.OpenAsTextStream(1,-1)
	Response.BinaryWrite(b.Read(999999))

End Sub

function asp_fclose(header)
    header.close
end function 


function asp_fopen(spath,iomode)
    Dim fso, iomode2, fp
    Set fso = CreateObject("Scripting.FileSystemObject")
    spath2=getabspath(spath)
    if iomode="a" then
        iomode2=8
    elseif iomode="w" then
        iomode2=2
    else
        iomode2=1
    end if
	on Error resume next
    If fso.FileExists(spath2) then
		set fp = fso.OpenTextFile(spath2, iomode2)
    else
        set fp = fso.CreateTextFile(spath2)
    end if
	on error goto 0
	if IsObject(fp) then
		set asp_fopen=fp
	else
		asp_fopen=false
	end if
end function

function fputs(header,str)
    header.Write(str)
end function

function filesize(filename)
    dim fs,ff
    set fs=Server.CreateObject("Scripting.FileSystemObject")
		filename=getabspath(filename)
    set ff=fs.GetFile(filename)
    res=ff.Size
    set ff=nothing
    set fs=nothing
    filesize=res
end function

function session_id()
    session_id=Session.SessionID
end function

function pow(x,y)
    pow=Exp(y* Log(x))
end function

function log10(x)
    log10=Log(x)/Log(10)
end function


function ob_start()
	ob_enabled=ob_enabled+1
	output_buffer(ob_enabled) = ""
end function

function ob_get_contents()
	ob_get_contents = output_buffer(ob_enabled)
end function

function ob_end_clean()
	output_buffer(ob_enabled) = ""
	ob_enabled = ob_enabled - 1
end function

sub ResponseWrite(str)
    if vartype(str)=vbBoolean then
        if(str) then 
            str="1"
        else
            str=""
        end if
    end if
	if ob_enabled>0 then
		output_buffer(ob_enabled) = output_buffer(ob_enabled) & str
	else
		Response.Write str
	end if
end sub

function xtempl_call_func(byval func,byref params)
	Execute func & " params"
end function

function is_string(byref val)
	is_string = (not isObject(val)) and vartype(val)=vbString
end function

function is_bool(byref val)
	is_bool = (not isObject(val)) and vartype(val)=vbBoolean
end function


function is_a(byref val,byval name)
	is_a = typename(val)=name
end function

function PropertyExists(byref obj,byval name)
	dim str
	if not isobject(obj) then
		PropertyExists=false
		exit function
	end if
	on error resume next
	str = "vartype obj." & name
	Execute str
	PropertyExists = err.Number=0 
	on error goto 0
end function

function array_pop(byref arr)
	if not IsDictionary(arr) then
		array_pop=NULL
		exit function
	end if
	if arr.Count=0 then
		array_pop=NULL
		exit function
	end if
	DoAssignmentByref array_pop,arr(arr.Count-1)
	arr.Remove(arr.Count-1)
end function
function echoBinary(byref value, byval dummy)
	response.binarywrite value
end function

function secondsPassedFrom(datetime)
    dim arrDateTime
	set arrDateTime=db2time(datetime)
	secondsPassedFrom = datediff("s",arrDateTime(0) & "-" & arrDateTime(1) & "-" & arrDateTime(2) & " " & arrDateTime(3) & ":" & arrDateTime(4) & ":" & arrDateTime(5),now())
end function

function setObjectProperty(byref obj,byval key,byref value)
	on error resume next
	doClassAssignmentByRef obj, key, value
end function

function returnError404()
    response.Status=404
end function

' checks if dictionary contains numeric keys 0-N only
function IsArrayDict(byref dict)
	dim i
	for i=0 to dict.Count-1
		if not asp_array_key_exists(i,dict) then
			IsArrayDict=false
			exit function
		end if
	next
	IsArrayDict=true
end function

function execute_events(ByRef params)
    if bValue(asp_function_exists(ArrayElement(params,"custom1"))) then
		execute ArrayElement(params,"custom1") & "(params)"
	end if
end function

function is_object(byref var)
	is_object = IsObject(var)
end function

function PrepareBlobs(byref values, byref blobfields)
	set PrepareBlobs = CreateDictionary()
	set blobfields = CreateDictionary()
end function

function ExecuteUpdate(strSQL,byref blobs,addMode)
'	exec SQL and read error message	
	error_happened=false
	on error resume next
	conn.Execute strSQL
	If err.Number=0 Then
		ExecuteUpdate=true
		exit function
	end if
	ExecuteUpdate=false
	error_happened = true
'	adding
	if addMode then
		if inlineadd<>ADD_SIMPLE then
			message="" & "Record was NOT added" & ". " & err.description
		else  
			message="<<< " & "Record was NOT added" & " >>><br><br>" & err.description
		end if
		readavalues=true
	else
		if inlineedit then
			message="" & "Record was NOT edited" & ". " & err.description
		else  
			message="<<< " & "Record was NOT edited" & " >>><br><br>" & err.description
		end if
		readevalues=true
	end if
end function



function usort(byref arr, compfuncname)
	if arr.count>1 then
		qsort arr,0,arr.count-1,compfuncname
	end if
end function


function qsortcompare(compfuncname,byref arg1,byref arg2)
	dim str
	str = "qsortcompare = " & compfuncname & "(arg1,arg2)"
	execute str
end function

function swapItems(byref arr, i1, i2)
	dim temp
	DoAssignmentByRef temp,arr(i1)
	setArrElementByRef arr,i1,arr(i2)
	setArrElementByRef arr,i2,temp
end function

function qsort(byref arr,loBound,hiBound,compfuncname)
	Dim pivot,loSwap,hiSwap,temp
'	two items
	if hiBound - loBound = 1 then
		if qsortcompare(compfuncname,arr(loBound),arr(hiBound))>0 then
			swapItems arr,loBound,hiBound
		End If
	End If
	doAssignmentByRef pivot,arr(int((loBound + hiBound) / 2))
	swapItems arr,int((loBound + hiBound) / 2),loBound
	loSwap = loBound + 1
	hiSwap = hiBound
	do
	    while loSwap < hiSwap and qsortcompare(compfuncname,arr(loSwap),pivot)<=0
	      loSwap = loSwap + 1
    	wend
	    while loSwap <= hiSwap and qsortcompare(compfuncname,arr(hiSwap),pivot)>=0
    	  hiSwap = hiSwap - 1
	    wend
	    if loSwap < hiSwap then
			swapItems arr,loSwap,hiSwap
	    End If
	loop while loSwap < hiSwap

	setArrElementByRef arr,loBound,arr(hiSwap)
	setArrElementByRef arr,hiSwap, pivot
	if loBound < (hiSwap - 1) then 
		qsort arr,loBound,hiSwap-1,compfuncname
	end if
	if hibound > hiSwap + 1  then 
		qsort arr,hiSwap+1,hiBound,compfuncname
	end if
End function

function asp_trim(str)
	asp_trim = trim(CSmartStr(str))
end function

function xtempl_include_header(xt,fname,param)
	if not asp_file_exists(getabspath(param)) then
		exit function
	end if
	if filesize(getabspath(param))>0 then
		xt.assign_function_p3 fname,"server.Execute",param
	end if
end function

function db_query_safe(sSQL,conn,byref errstr)
    dim asp_rs
    Set asp_rs = server.CreateObject("ADODB.Recordset")
    err.clear
	on error resume next
	asp_rs.Open sSQL,conn
	errstr=err.description
	if err.number=0 then
		set db_query_safe=asp_rs
	else
		set db_query_safe=false
	end if
	on error goto 0
end function

function binPrint(byref value, size)
	response.BinaryWrite value
end function

function WRGetAbsoluteFileName(filename)
	WRGetAbsoluteFileName=Server.MapPath(filename)
end function


function GetMySQLLastInsertID()
    dim rs
	Set rs = server.CreateObject("ADODB.Recordset")
	rs.Open "select LAST_INSERT_ID()",conn
	GetMySQLLastInsertID=empty
	if rs.eof then
		set rs=nothing
		exit function
	end if
	GetMySQLLastInsertID = rs.fields(0).value
	set rs=nothing
end function

Function GoodFieldName(ByVal field)
	Dim out,i,t
	field = CSmartStr(field)
	if gGoodFieldNameCache.Exists(field) then
		GoodFieldName = gGoodFieldNameCache(field)
		exit function
	end if
	out = ""
	i = 0
	do while i<asp_strlen(field)
		doAssignment t,asp_substr(field,i,1)
		if ((asc(t)<asc("a") or asc("z")<asc(t)) and (asc(t)<asc("A") or asc("Z")<asc(t))) and (asc(t)<asc("0") or asc("9")<asc(t)) then
			out = out & "_"
		else
			out = out & t
		end if
		i = i+1
	loop
	doAssignment GoodFieldName,out
	gGoodFieldNameCache(field) = out
End Function

Function GetFieldData(ByVal table,ByVal field,ByVal key,ByVal default)
	if table="" then
		table=strTableName
	end if
	if not tables_data.Exists(table) then
		doAssignmentByRef GetFieldData,default
		Exit Function
	end if
	if not tables_data(table).Exists(field) then
		doAssignmentByRef GetFieldData,default
		Exit Function
	end if
	if not tables_data(table)(field).Exists(key) then
		doAssignmentByRef GetFieldData,default
		Exit Function
	end if
	if IsObject(tables_data(table)(field)(key)) then
		set GetFieldData = tables_data(table)(field)(key)
	else
		GetFieldData = tables_data(table)(field)(key)
	end if
End Function


Function GetTableData(ByVal table,ByVal key,ByVal default)

	if table="" then
		table=strTableName
	end if
	if not tables_data.Exists(table) then
		doAssignmentByRef GetTableData,default
		Exit Function
	end if
	if not tables_data(table).Exists(key) then
		doAssignmentByRef GetTableData,default
		Exit Function
	end if
	if IsObject(tables_data(table)(key)) then
		set GetTableData = tables_data(table)(key)
	else
		GetTableData = tables_data(table)(key)
	end if
End Function

Function xt_getvar(byref this_object,ByVal name)
	Dim i
	i = this_object.xt_stack.count-1
	do while 0<=i
		if isobject(this_object.xt_stack(i)) then 
			if this_object.xt_stack(i).Exists(name) then
				doAssignmentByRef xt_getvar,this_object.xt_stack(i)(name)
				Exit Function
			end if
		end if
		i = i-1
	loop
	xt_getvar = false
End Function

Function xt_process_template(ByRef xt,ByVal str)
	Dim start,literal,length,pos,section,var,message,endpos,section_name,endtag,endpos1,this,begin,var_end,keys,i,varparams,var_name,tag
	start = 0
	literal = false
	str = CSmartStr(str)
	length = len(str)
	do while true
		i_xtempl_exitLoop5=false
		do
			pos = instr(start+1,str,"{")-1
			if pos=-1 then
				ResponseWrite mid(str,start+1,length-start)
				i_xtempl_exitLoop5=true
				exit do
			end if
			section = false
			var = false
			message = false
			if mid(str,pos+2,6)="BEGIN " then
				section = true
			elseif mid(str,pos+2,1)="$" then
					var = true
			elseif mid(str,pos+2,14)="mlang_message " then
				message = true
			else
				ResponseWrite mid(str,start+1,pos-start+1)
				start = pos+1
				exit do
			end if
			ResponseWrite mid(str,start+1,pos-start)
			if section then
				endpos = instr(pos+1,str,"}")-1
				if endpos=-1 then
					xt.report_error_p1 "Page is broken"
					Exit Function
				end if
				section_name = trim(mid(str,pos+8,endpos-pos-7))
				endtag = "{END " & section_name & "}"
				endpos1 = instr(endpos+1,str,endtag)-1
				if endpos1=-1 then
					ResponseWrite "End tag not found:" & htmlspecialchars(endtag)
					xt.report_error_p1 "Page is broken"
					Exit Function
				end if
				start = endpos1+len(endtag)
				doAssignmentByRef var,xt_getvar(xt,section_name)
				if IsFalse(var) then
					exit do
				end if
				begin = ""
				var_end = ""
				if IsObject(var) then
					doAssignment begin,ArrayElement(var,"begin")
					doAssignment var_end,ArrayElement(var,"end")
					doAssignment var,ArrayElement(var,"data")
				end if
				if not IsObject(var) then
					ResponseWrite begin
					xt_process_template xt, mid(str,endpos+2,endpos1-endpos-1)
					xt.processVar_p2 var_end,varparams
				else
				
					ResponseWrite begin
					keys = var.keys
					for each i in keys
						if isObject(var(i)) then
							if var(i).Exists("begin") then 
								responsewrite var(i)("begin")
							end if
						end if
						setArrElementByRef xt.xt_stack,asp_count(xt.xt_stack),var(i)
						xt_process_template xt, mid(str,endpos+2,endpos1-endpos-1)
						array_pop xt.xt_stack
						if isObject(var(i)) then
							if var(i).Exists("end") then 
								responsewrite var(i)("end")
							end if
						end if
					next
					xt.processVar_p2 var_end,varparams
				end if
			else
				if bValue(var) then
					endpos = instr(pos+1,str,"}")-1
					if endpos=-1 then
						xt.report_error_p1 "Page is broken"
						Exit Function
					end if
					Set varparams = (CreateDictionary())
					var_name=trim(mid(str,pos+2+1,endpos-pos-2))
					if instr(var_name," ")>0 then
						doAssignmentByRef varparams,explode(" ",var_name)
						var_name = varparams(0)
						varparams.Remove(0)
					end if
					start = endpos+1
					doAssignmentByRef var,xt_getvar(xt,var_name)
					if IsFalse(var) then
						exit do
					end if
					xt.processVar_p2 var,varparams
				else
					if bValue(message) then
						doAssignmentByRef endpos,asp_strpos(str,"}",pos)
						if IsFalse(endpos) then
							xt.report_error_p1 "Page is broken"
							Exit Function
						end if
						doAssignmentByRef tag,trim(asp_substr(str,CSmartDbl(pos)+15,(CSmartDbl(endpos)-CSmartDbl(pos))-15))
						start = CSmartDbl(endpos)+1
						ResponseWrite htmlspecialchars(mlang_message(tag))
					end if
				end if
			end if
		loop while false
		if i_xtempl_exitLoop5 then _
			exit do
	loop
End Function
function isabspath(path)
	if mid(path,2,1)=":" or mid(path,1,2)="\\" or instr(path,"://")<>0 then
		isabspath=true
	else
		isabspath=false
	end if
end function

function bin2hex(byref val)
	dim str,i,c,a,s
	str = ""
	for i=1 to lenb(val)
		c = midb(val,i,1)
		a = ascb(c)
		if a<16 then
			str = str & "0"
		end if
		str = str & hex(a)
	next
	bin2hex=lcase(str)
end function


function db_getfieldslist(table)
	dim res,strSQL,rs,i
	set res = CreateDictionary()
	table_tmp=table
	if IsEqual(GetDatabaseType(),2) then
		doAssignmentByRef pos,asp_strrpos(asp_strtoupper(table_tmp),"ORDER BY",empty)
		if bValue(pos) then
			doAssignmentByRef table_tmp,asp_substr(table_tmp,0,pos)
		end if
	end if	
	if not BValue(IsStoredProcedure(table_tmp)) then
		if not IsEqual(GetDatabaseType(),1) then
			strSQL="select * from (" & table_tmp & ") as t where 1=0"
		else
			strSQL="select * from (" & table_tmp & ") where 1=0"
		end if
	else
		strSQL=table_tmp
	end if
	Set rs = server.CreateObject("ADODB.Recordset")
	rs.Open strSQL,conn
	for i=0 to db_numfields(rs)-1
		j=res.count
		set res(j) = CreateDictionary()
		res(j)("fieldname")=db_fieldname(rs,i)
		res(j)("type")=rs(i).Type
		res(j)("not_null")=false
	next
	doAssignment db_getfieldslist,res
end function

function db_gettablelist()
	dim rstSchema, res
	set res = CreateDictionary()
	set rstSchema = server.CreateObject("ADODB.Recordset")
	Set rstSchema = conn.OpenSchema(20)
	while not rstSchema.EOF
		if Trim(rstSchema("TABLE_TYPE")) = "TABLE" or Trim(rstSchema("TABLE_TYPE")) = "VIEW" then _
			res(res.count)=rstSchema("TABLE_NAME")
		rstSchema.MoveNext
	wend
	rstSchema.Close
	doAssignment db_gettablelist,res
 end function
 
 function getFileNameFromURL()
	dim scriptname,pos
	scriptname=Request.ServerVariables("SCRIPT_NAME")
	pos=instrrev(scriptname,"/")
	if pos<>0 then scriptname=mid(scriptname,pos+1)
	getFileNameFromURL=scriptname
end function

function db2_escape_string(str)
	db2_escape_string = asp_str_replace("'","''",str)
end function

function mysql_real_escape_string(str)
	mysql_real_escape_string = asp_str_replace("'","''",str)
end function

function pg_escape_string(str)
	pg_escape_string = asp_str_replace("'","''",str)
end function

function pg_escape_bytea(str)
	if lenb(val)=0 then
		pg_escape_bytea ="''"
	else
		pg_escape_bytea = "0x" & bin2hex(val)
	end if
end function

function pg_unescape_bytea(str)
	if isnull(str) or isempty(str) then
		pg_unescape_bytea=""
		exit function
	end if
	pg_unescape_bytea = str
end function

function strlen_bin(byref str)
	if isnull(str) or isempty(str) then
 		strlen_bin=0
		exit function
	end if
	strlen_bin=lenb(str)
end function

function db_stripslashesbinaryAccess(str)
	if isnull(str) or isempty(str) then
		db_stripslashesbinaryAccess=""
		exit function
	end if
'	try to remove ole header for BMP pictures
	pos = instrb(str,unicode2bytes(".Picture"))
	if pos=0 or pos>300 then 
		db_stripslashesbinaryAccess = str
		exit function
	end if
	pos1=instrb(pos,str,unicode2bytes("BM"))
	if pos1=0 or pos1>300 then
		db_stripslashesbinaryAccess = str
		exit function
	end if
	db_stripslashesbinaryAccess = midb(str,pos1)
end function

function SendContentLength(len)
end function

function runner_move_uploaded_file(source, dest)
end function

function add_mysql_binaryslashes(val)
	if lenb(val)=0 then
		add_mysql_binaryslashes ="''"
	else
		add_mysql_binaryslashes = "0x" & bin2hex(val)
	end if
end function

function escapeEntities(str)
  escapeEntities = str
end function

function DecodeUTF8(str)
	DecodeUTF8 = ConvertUtf8BytesToString(ConvertStringToUtf8Bytes(str))
end function

Public Function ConvertStringToUtf8Bytes(ByRef strText)
    Dim objStream
    ' init stream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Charset = cCharset
    objStream.Mode = 3
    objStream.Type = 2
    objStream.Open
    
    ' write bytes into stream
    objStream.WriteText strText
    objStream.Flush
    
    ' rewind stream and read text
    objStream.Position = 0
    objStream.Type = 1
    objStream.Read 0

	ConvertStringToUtf8Bytes = objStream.Read
End Function

Public Function ConvertUtf8BytesToString(data)

    Set objStream = CreateObject("ADODB.Stream")
    Dim strTmp
    ' init stream
    objStream.Charset = "utf-8"
    objStream.Mode = 3
    objStream.Type = 1
    objStream.Open
    
    ' write bytes into stream
    objStream.Write data
    objStream.Flush
    
    ' rewind stream and read text
    objStream.Position = 0
    objStream.Type = 2
    strTmp = objStream.ReadText
    
    ' close up and return
    objStream.Close
    ConvertUtf8BytesToString = strTmp

End Function
%>