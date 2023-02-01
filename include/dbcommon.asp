<%@codepage=1252%>
<%

cCharset = "Windows-1252"

response.Charset=cCharset

useAJAX=true
suggestAllContent=true
gLoadSearchControls = 10

SQLUpdateMode=false

Session.LCID = 7177
session.codepage=1252


nDatabaseType = 2


dim dbConnection
set gGoodFieldNameCache = CreateDictionary()
dim conn

Set myRequest = CreateObject("Scripting.Dictionary")
Set myRequestFiles = CreateObject("Scripting.Dictionary")
formParsed=0

%>
<!--#include file="locale.asp"-->
<!--#include file="events.asp"-->
<!--#include file="commonfunctions.asp"-->
<!--#include file="dbconnection.asp"-->
<!--#include file="dbfunctions.asp"-->
<!--#include file="aspfunctions.asp"-->
<!--#include file="dal.asp"-->
<!--#include file="appsettings.asp"-->
<%



function db_connect()
	set dbConnection = server.CreateObject("ADODB.Connection")
   	dbConnection.ConnectionString = strConnection
   	dbConnection.Open
    set db_connect=dbConnection	
end function

function db_closequery(qhandle)
	qhandle.close
end function


function AddTableWrappers(strName)
	if mid(strName,1,1)=strLeftWrapper then
		AddTableWrappers = strName
		exit function
	end if
	dim arr
	arr=split(strName,".")
	ret=strLeftWrapper & arr(0) & strRightWrapper
	if ubound(arr)>0 then ret=ret & "." & strLeftWrapper & arr(1) & strRightWrapper
	AddTableWrappers = ret
end function

function db_upper(dbval)
	db_upper = "upper(" & dbval & ")"
end function

function AddFieldWrappers(strName)
	if mid(strName,1,1)=strLeftWrapper then
		AddFieldWrappers = strName
	else
		AddFieldWrappers = strLeftWrapper & strName & strRightWrapper
	end if
end function
function FieldNeedQuotes(rs,field)
	ttype=db_fieldtype(rs,field)
	if ttype=20 or ttype=128 or ttype=11 or ttype=6 or ttype=14 or ttype=5 or ttype=3 or ttype=131 _
	or ttype=4	or ttype=2	or ttype=16 or ttype=21 or ttype=19 or ttype=18 or ttype=17 or ttype=139 then
		FieldNeedQuotes = false
	else
		FieldNeedQuotes = true
	end if
end function

function db_datequotes(val)
	db_datequotes = "convert(datetime,'" & val & "',120)"
end function

function db_fieldtype(lhandle,fname)
	Dim i
	for i=0 to db_numfields(lhandle)-1
		if db_fieldname(lhandle,i)=fname then
			ttype=db_fieldtypen(lhandle,i)
			db_fieldtype = ttype
			exit function
		end if
	next
	db_fieldtype = ""
end function
function db_numfields(lhandle)
	db_numfields = lhandle.Fields.Count
end function

function db_fieldname(lhandle,fnumber)
	db_fieldname = lhandle.Fields(fnumber).Name
end function

function db_fieldtypen(lhandle,fnumber)
	db_fieldtypen = lhandle.Fields(fnumber).Type
end function

function date2str(val)
	if isnull(val) then
		date2str=""
		exit function
	end if
	if isdate(val) then
		date2str = CStr(year(val)) & "-" & CStr(month(val)) & "-" & CStr(day(val)) & _
				" " & CStr(hour(val)) & ":" & CStr(minute(val)) & ":" & CStr(second(val))
		exit function
	end if
	date2str=""
end function


%>
