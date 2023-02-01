<%
%>
<!--#include file="include/dbcommon.asp"-->
<%
asp_include "classes/searchclause.asp",false
asp_include "include/payment_variables.asp",false
gQuery.ReplaceFieldsWithDummies_p1 GetBinaryFieldsIndices("")
if bValue(eventObj.exists_p1("BeforeProcessExport")) then
	eventObj.BeforeProcessExport_p1 conn
end if
strWhereClause = ""
strHavingClause = ""
Set selected_recs = (CreateDictionary())
options = "1"
asp_header "Expires: Thu, 01 Jan 1970 00:00:01 GMT"
asp_include "include/xtempl.asp",false
asp_include "classes/runnerpage.asp",false
Set xt = (CreateClass("Xtempl",0,Empty,Empty,Empty,Empty,Empty,Empty,Empty))
doAssignment id,IIF(not IsEqual(postvalue("id"),""),postvalue("id"),1)
Set params = (CreateDictionary3("pageType",PAGE_EXPORT,"id",id,"tName",strTableName))
setArrElementByRef params,"xt",xt
if not bValue(eventObj.exists_p1("ListGetRowCount")) and not bValue(eventObj.exists_p1("ListQuery")) then
	setArrElement params,"needSearchClauseObj",false
end if
Set pageObject = (CreateClass("RunnerPage",1,params,Empty,Empty,Empty,Empty,Empty,Empty))
if not IsEqual(GetRequestValue(Request,"a"),"") then
	options = ""
	sWhere = "1=0"
	Set selected_recs = (CreateDictionary())
	if bValue(GetRequestValue(Request,"mdelete")) then
		GetCollectionBounds GetRequestValue(Request,"mdelete"),export_loopIdx2,export_loopMax2
		do while export_loopIdx2<=export_loopMax2
			export_arrKey2 = GetCollectionKey(GetRequestValue(Request,"mdelete"),export_loopIdx2)
			doAssignment ind,GetRequestValue(GetRequestValue(Request,"mdelete"),export_arrKey2)
			Set keys = (CreateDictionary())
			setArrElement selected_recs,asp_count(selected_recs),keys
			export_loopIdx2=export_loopIdx2+1
		loop
	else
		if bValue(GetRequestValue(Request,"selection")) then
			GetCollectionBounds GetRequestValue(Request,"selection"),export_loopIdx3,export_loopMax3
			export_exitLoop3=false
			do while export_loopIdx3<=export_loopMax3
				export_exitLoop3=false
				do
					export_arrKey3 = GetCollectionKey(GetRequestValue(Request,"selection"),export_loopIdx3)
					doAssignment keyblock,GetRequestValue(GetRequestValue(Request,"selection"),export_arrKey3)
					doAssignmentByRef arr,explode("&",keyblock)
					if IsLess(asp_count(arr),0) then
						exit do
					end if
					Set keys = (CreateDictionary())
					setArrElement selected_recs,asp_count(selected_recs),keys
				loop while false
				if export_exitLoop3 then _
					exit do
				export_loopIdx3=export_loopIdx3+1
			loop
		end if
	end if
	GetCollectionBounds selected_recs,export_loopIdx4,export_loopMax4
	do while export_loopIdx4<=export_loopMax4
		export_arrKey4 = GetCollectionKey(selected_recs,export_loopIdx4)
		doAssignment keys,ArrayElement(selected_recs,export_arrKey4)
		sWhere = CSmartStr(sWhere) & " or "
		sWhere = CSmartStr(sWhere) & CSmartStr(KeyWhere(keys,""))
		export_loopIdx4=export_loopIdx4+1
	loop
	doAssignmentByRef strSQL,gSQLWhere(sWhere,"")
	doAssignment strWhereClause,sWhere
	setArrElement Session,CSmartStr(strTableName) & "_SelectedSQL",strSQL
	setArrElement Session,CSmartStr(strTableName) & "_SelectedWhere",sWhere
	setArrElement Session,CSmartStr(strTableName) & "_SelectedRecords",selected_recs
end if
if not IsEqual(Session(CSmartStr(strTableName) & "_SelectedSQL"),"") and IsEqual(GetRequestValue(Request,"records"),"") then
	doAssignment strSQL,Session(CSmartStr(strTableName) & "_SelectedSQL")
	doAssignment strWhereClause,Session(CSmartStr(strTableName) & "_SelectedWhere")
	doAssignment selected_recs,Session(CSmartStr(strTableName) & "_SelectedRecords")
else
	doAssignment strWhereClause,Session(CSmartStr(strTableName) & "_where")
	doAssignment strHavingClause,Session(CSmartStr(strTableName) & "_having")
	doAssignmentByRef strSQL,gSQLWhere(strWhereClause,strHavingClause)
end if
mypage = 1
if bValue(GetRequestValue(Request,"type")) then
	doAssignment strOrderBy,Session(CSmartStr(strTableName) & "_order")
	if not bValue(strOrderBy) then
		doAssignment strOrderBy,gstrOrderBy
	end if
	strSQL = CSmartStr(strSQL) & (" " & CSmartStr(trim(strOrderBy)))
	doAssignment strSQLbak,strSQL
	if bValue(eventObj.exists_p1("BeforeQueryExport")) then
		eventObj.BeforeQueryExport_p3 strSQL,strWhereClause,strOrderBy
	end if
	if not IsEqual(strSQL,strSQLbak) then
		doAssignmentByRef numrows,GetRowCount(strSQL)
	else
		doAssignmentByRef strSQL,gSQLWhere(strWhereClause,strHavingClause)
		strSQL = CSmartStr(strSQL) & (" " & CSmartStr(trim(strOrderBy)))
		rowcount = false
		if bValue(eventObj.exists_p1("ListGetRowCount")) then
			Set masterKeysReq = (CreateDictionary())
			i = 0
			do while IsLess(i,asp_count(pageObject.detailKeysByM))
				setArrElement masterKeysReq,asp_count(masterKeysReq),Session((CSmartStr(strTableName) & "_masterkey") & CSmartStr(CSmartDbl(i)+1))
				i = CSmartDbl(i)+1
			loop
			doAssignmentByRef rowcount,eventObj.ListGetRowCount_p4(pageObject.searchClauseObj,Session(CSmartStr(strTableName) & "_mastertable"),masterKeysReq,selected_recs)
		end if
		if not IsFalse(rowcount) then
			doAssignment numrows,rowcount
		else
			doAssignmentByRef numrows,gSQLRowCount(strWhereClause,strHavingClause)
		end if
	end if
	LogInfo strSQL
	nPageSize = 0
	if IsEqual(GetRequestValue(Request,"records"),"page") and bValue(numrows) then
		mypage = CSmartLng(Session(CSmartStr(strTableName) & "_pagenumber"))
		nPageSize = CSmartLng(Session(CSmartStr(strTableName) & "_pagesize"))
		if not bValue(nPageSize) then
			doAssignmentByRef nPageSize,GetTableData(strTableName,".pageSize",0)
		end if
		if IsLess(nPageSize,0) then
			nPageSize = 0
		end if
		if IsLess(0,nPageSize) then
			if IsLessOrEqual(numrows,(CSmartDbl(mypage)-1)*CSmartDbl(nPageSize)) then
				doAssignmentByRef mypage,asp_ceil(CSmartDbl(numrows)/CSmartDbl(nPageSize))
			end if
			if not bValue(mypage) then
				mypage = 1
			end if
			doAssignmentByRef strSQL,AddTop(strSQL,CSmartDbl(mypage)*CSmartDbl(nPageSize))
		end if
	end if
	listarray = false
	if bValue(eventObj.exists_p1("ListQuery")) then
		doAssignmentByRef listarray,eventObj.ListQuery_p8(pageObject.searchClauseObj,Session(CSmartStr(strTableName) & "_arrFieldForSort"),Session(CSmartStr(strTableName) & "_arrHowFieldSort"),Session(CSmartStr(strTableName) & "_mastertable"),masterKeysReq,selected_recs,nPageSize,mypage)
	end if
	if not IsFalse(listarray) then
		doAssignment rs,listarray
	else
		if IsLess(0,nPageSize) then
			doAssignmentByRef rs,db_query(strSQL,conn)
			db_pageseek rs,nPageSize,mypage
		else
			doAssignmentByRef rs,db_query(strSQL,conn)
		end if
	end if
	if not bValue(false) then
		Server.ScriptTimeOut=300
	end if
	if IsEqual(GetRequestValue(Request,"type"),"excel") then
		setArrElement locale_info,"LOCALE_SGROUPING","0"
		setArrElement locale_info,"LOCALE_SMONGROUPING","0"
		ExportToExcel 
	else
		if IsEqual(GetRequestValue(Request,"type"),"word") then
			ExportToWord 
		else
			if IsEqual(GetRequestValue(Request,"type"),"xml") then
				ExportToXML 
			else
				if IsEqual(GetRequestValue(Request,"type"),"csv") then
					setArrElement locale_info,"LOCALE_SGROUPING","0"
					setArrElement locale_info,"LOCALE_SDECIMAL","."
					setArrElement locale_info,"LOCALE_SMONGROUPING","0"
					setArrElement locale_info,"LOCALE_SMONDECIMALSEP","."
					ExportToCSV 
				end if
			end if
		end if
	end if
	db_close conn
	response.end
end if
pageObject.addButtonHandlers 
doAssignmentByRef onLoadJsCode,GetTableData(pageObject.tName,".jsOnloadExport","")
pageObject.addOnLoadJsEvent_p1 onLoadJsCode
if bValue(options) then
	xt.assign_p2 "rangeheader_block",true
	xt.assign_p2 "range_block",true
end if
xt.assign_p2 "exportlink_attrs",("id=""saveButton" & CSmartStr(pageObject.id)) & """"
setArrElement pageObject.body,"begin",CSmartStr(ArrayElement(pageObject.body,"begin")) & "<script type=""text/javascript"" src=""include/jquery.js""></script>" & vbcrlf
setArrElement pageObject.body,"begin",CSmartStr(ArrayElement(pageObject.body,"begin")) & "<script language=""JavaScript"" src=""include/jsfunctions.js""></script>" & vbcrlf
if IsIdentical(pageObject.debugJSMode,true) then
	setArrElement pageObject.body,"begin",CSmartStr(ArrayElement(pageObject.body,"begin")) & "<script language=""JavaScript"" src=""include/runnerJS/RunnerBase.js""></script>" & vbcrlf
else
	setArrElement pageObject.body,"begin",CSmartStr(ArrayElement(pageObject.body,"begin")) & "<script language=""JavaScript"" src=""include/runnerJS/RunnerBase.js""></script>" & vbcrlf
end if
pageObject.fillSetCntrlMaps 
setArrElement pageObject.body,"end",CSmartStr(ArrayElement(pageObject.body,"end")) & "<script>"
setArrElement pageObject.body,"end",CSmartStr(ArrayElement(pageObject.body,"end")) & (("window.controlsMap = '" & CSmartStr(jsreplace(my_json_encode(pageObject.controlsHTMLMap)))) & "';")
setArrElement pageObject.body,"end",CSmartStr(ArrayElement(pageObject.body,"end")) & (("window.settings = '" & CSmartStr(jsreplace(my_json_encode(pageObject.jsSettings)))) & "';")
setArrElement pageObject.body,"end",CSmartStr(ArrayElement(pageObject.body,"end")) & "</script>"
pageObject.addCommonJs 
setArrElement pageObject.body,"end",CSmartStr(ArrayElement(pageObject.body,"end")) & (("<script>" & CSmartStr(pageObject.PrepareJS())) & "</script>")
xt.assignbyref_p2 "body",pageObject.body
xt.display_p1 "payment_export.htm"
Function ExportToExcel()
	asp_header "Content-Type: application/vnd.ms-excel"
	asp_header "Content-Disposition: attachment;Filename=payment.xls"
	ResponseWrite "<html>"
	ResponseWrite "<html xmlns:o=""urn:schemas-microsoft-com:office:office"" xmlns:x=""urn:schemas-microsoft-com:office:excel"" xmlns=""http://www.w3.org/TR/REC-html40"">"
	ResponseWrite ("<meta http-equiv=""Content-Type"" content=""text/html; charset=" & CSmartStr(cCharset)) & """>"
	ResponseWrite "<body>"
	ResponseWrite "<table border=1>"
	WriteTableData 
	ResponseWrite "</table>"
	ResponseWrite "</body>"
	ResponseWrite "</html>"
End Function
Function ExportToWord()
	asp_header "Content-Type: application/vnd.ms-word"
	asp_header "Content-Disposition: attachment;Filename=payment.doc"
	ResponseWrite "<html>"
	ResponseWrite ("<meta http-equiv=""Content-Type"" content=""text/html; charset=" & CSmartStr(cCharset)) & """>"
	ResponseWrite "<body>"
	ResponseWrite "<table border=1>"
	WriteTableData 
	ResponseWrite "</table>"
	ResponseWrite "</body>"
	ResponseWrite "</html>"
End Function
Function ExportToXML()
	Dim row,i,values,eventRes,field,fName
	asp_header "Content-Type: text/xml"
	asp_header "Content-Disposition: attachment;Filename=payment.xml"
	if bValue(eventObj.exists_p1("ListFetchArray")) then
		doAssignmentByRef row,eventObj.ListFetchArray_p1(rs)
	else
		doAssignmentByRef row,db_fetch_array(rs)
	end if
	if not bValue(row) then
		Exit Function
	end if
	ResponseWrite ("<?xml version=""1.0"" encoding=""" & CSmartStr(cCharset)) & """ standalone=""yes""?>" & vbcrlf
	ResponseWrite "<table>" & vbcrlf
	i = 0
	do while (not bValue(nPageSize) or IsLess(i,nPageSize)) and bValue(row)
		Set values = (CreateDictionary())
		setArrElement values,"station",GetData(row,"station","")
		setArrElement values,"name",GetData(row,"name","")
		setArrElement values,"payment_date",GetData(row,"payment_date","")
		setArrElement values,"fop",GetData(row,"fop","")
		setArrElement values,"prefix",GetData(row,"prefix","")
		setArrElement values,"serial",GetData(row,"serial","")
		setArrElement values,"amount_received",GetData(row,"amount_received","")
		eventRes = true
		if bValue(eventObj.exists_p1("BeforeOut")) then
			doAssignmentByRef eventRes,eventObj.BeforeOut_p2(row,values)
		end if
		if bValue(eventRes) then
			i = CSmartDbl(i)+1
			ResponseWrite "<row>" & vbcrlf
			GetCollectionBounds values,export_loopIdx7,export_loopMax7
			do while export_loopIdx7<=export_loopMax7
				fName = GetCollectionKey(values,export_loopIdx7)
				doAssignment val,ArrayElement(values,fName)
				doAssignmentByRef field,htmlspecialchars(XMLNameEncode(fName))
				ResponseWrite ("<" & CSmartStr(field)) & ">"
				ResponseWrite htmlspecialchars(ArrayElement(values,fName))
				ResponseWrite ("</" & CSmartStr(field)) & ">" & vbcrlf
				export_loopIdx7=export_loopIdx7+1
			loop
			ResponseWrite "</row>" & vbcrlf
		end if
		if bValue(eventObj.exists_p1("ListFetchArray")) then
			doAssignmentByRef row,eventObj.ListFetchArray_p1(rs)
		else
			doAssignmentByRef row,db_fetch_array(rs)
		end if
	loop
	ResponseWrite "</table>" & vbcrlf
End Function
Function ExportToCSV()
	Dim row,totals,outstr,iNumberOfRows,values,format,eventRes
	asp_header "Content-Type: application/csv"
	asp_header "Content-Disposition: attachment;Filename=payment.csv"
	if bValue(eventObj.exists_p1("ListFetchArray")) then
		doAssignmentByRef row,eventObj.ListFetchArray_p1(rs)
	else
		doAssignmentByRef row,db_fetch_array(rs)
	end if
	if not bValue(row) then
		Exit Function
	end if
	Set totals = (CreateDictionary())
	outstr = ""
	if not IsEqual(outstr,"") then
		outstr = CSmartStr(outstr) & ","
	end if
	outstr = CSmartStr(outstr) & """station"""
	if not IsEqual(outstr,"") then
		outstr = CSmartStr(outstr) & ","
	end if
	outstr = CSmartStr(outstr) & """name"""
	if not IsEqual(outstr,"") then
		outstr = CSmartStr(outstr) & ","
	end if
	outstr = CSmartStr(outstr) & """payment_date"""
	if not IsEqual(outstr,"") then
		outstr = CSmartStr(outstr) & ","
	end if
	outstr = CSmartStr(outstr) & """fop"""
	if not IsEqual(outstr,"") then
		outstr = CSmartStr(outstr) & ","
	end if
	outstr = CSmartStr(outstr) & """prefix"""
	if not IsEqual(outstr,"") then
		outstr = CSmartStr(outstr) & ","
	end if
	outstr = CSmartStr(outstr) & """serial"""
	if not IsEqual(outstr,"") then
		outstr = CSmartStr(outstr) & ","
	end if
	outstr = CSmartStr(outstr) & """amount_received"""
	ResponseWrite outstr
	ResponseWrite vbcrlf
	iNumberOfRows = 0
	do while (not bValue(nPageSize) or IsLess(iNumberOfRows,nPageSize)) and bValue(row)
		Set values = (CreateDictionary())
		format = ""
		setArrElement values,"station",GetData(row,"station",format)
		format = ""
		setArrElement values,"name",GetData(row,"name",format)
		format = "Short Date"
		setArrElement values,"payment_date",GetData(row,"payment_date",format)
		format = ""
		setArrElement values,"fop",GetData(row,"fop",format)
		format = ""
		setArrElement values,"prefix",GetData(row,"prefix",format)
		format = ""
		setArrElement values,"serial",GetData(row,"serial",format)
		format = "Number"
		setArrElement values,"amount_received",ArrayElement(row,"amount_received")
		eventRes = true
		if bValue(eventObj.exists_p1("BeforeOut")) then
			doAssignmentByRef eventRes,eventObj.BeforeOut_p2(row,values)
		end if
		if bValue(eventRes) then
			outstr = ""
			if not IsEqual(outstr,"") then
				outstr = CSmartStr(outstr) & ","
			end if
			outstr = CSmartStr(outstr) & (("""" & CSmartStr(asp_str_replace("""","""""",ArrayElement(values,"station")))) & """")
			if not IsEqual(outstr,"") then
				outstr = CSmartStr(outstr) & ","
			end if
			outstr = CSmartStr(outstr) & (("""" & CSmartStr(asp_str_replace("""","""""",ArrayElement(values,"name")))) & """")
			if not IsEqual(outstr,"") then
				outstr = CSmartStr(outstr) & ","
			end if
			outstr = CSmartStr(outstr) & (("""" & CSmartStr(asp_str_replace("""","""""",ArrayElement(values,"payment_date")))) & """")
			if not IsEqual(outstr,"") then
				outstr = CSmartStr(outstr) & ","
			end if
			outstr = CSmartStr(outstr) & (("""" & CSmartStr(asp_str_replace("""","""""",ArrayElement(values,"fop")))) & """")
			if not IsEqual(outstr,"") then
				outstr = CSmartStr(outstr) & ","
			end if
			outstr = CSmartStr(outstr) & (("""" & CSmartStr(asp_str_replace("""","""""",ArrayElement(values,"prefix")))) & """")
			if not IsEqual(outstr,"") then
				outstr = CSmartStr(outstr) & ","
			end if
			outstr = CSmartStr(outstr) & (("""" & CSmartStr(asp_str_replace("""","""""",ArrayElement(values,"serial")))) & """")
			if not IsEqual(outstr,"") then
				outstr = CSmartStr(outstr) & ","
			end if
			outstr = CSmartStr(outstr) & (("""" & CSmartStr(asp_str_replace("""","""""",ArrayElement(values,"amount_received")))) & """")
			ResponseWrite outstr
		end if
		iNumberOfRows = CSmartDbl(iNumberOfRows)+1
		if bValue(eventObj.exists_p1("ListFetchArray")) then
			doAssignmentByRef row,eventObj.ListFetchArray_p1(rs)
		else
			doAssignmentByRef row,db_fetch_array(rs)
		end if
		if ((not bValue(nPageSize) or IsLess(iNumberOfRows,nPageSize)) and bValue(row)) and bValue(eventRes) then
			ResponseWrite vbcrlf
		end if
	loop
End Function
Function WriteTableData()
	Dim row,var_REQUEST,totals,iNumberOfRows,values,format,eventRes
	if bValue(eventObj.exists_p1("ListFetchArray")) then
		doAssignmentByRef row,eventObj.ListFetchArray_p1(rs)
	else
		doAssignmentByRef row,db_fetch_array(rs)
	end if
	if not bValue(row) then
		Exit Function
	end if
	ResponseWrite "<tr>"
	if IsEqual(GetRequestValue(Request,"type"),"excel") then
		ResponseWrite ("<td style=""width: 100"" x:str>" & CSmartStr(PrepareForExcel("Station"))) & "</td>"
		ResponseWrite ("<td style=""width: 100"" x:str>" & CSmartStr(PrepareForExcel("Name"))) & "</td>"
		ResponseWrite ("<td style=""width: 100"" x:str>" & CSmartStr(PrepareForExcel("Payment Date"))) & "</td>"
		ResponseWrite ("<td style=""width: 100"" x:str>" & CSmartStr(PrepareForExcel("Fop"))) & "</td>"
		ResponseWrite ("<td style=""width: 100"" x:str>" & CSmartStr(PrepareForExcel("Prefix"))) & "</td>"
		ResponseWrite ("<td style=""width: 100"" x:str>" & CSmartStr(PrepareForExcel("Serial"))) & "</td>"
		ResponseWrite ("<td style=""width: 100"" x:str>" & CSmartStr(PrepareForExcel("Amount Received"))) & "</td>"
	else
		ResponseWrite ("<td>" & CSmartStr("Station")) & "</td>"
		ResponseWrite ("<td>" & CSmartStr("Name")) & "</td>"
		ResponseWrite ("<td>" & CSmartStr("Payment Date")) & "</td>"
		ResponseWrite ("<td>" & CSmartStr("Fop")) & "</td>"
		ResponseWrite ("<td>" & CSmartStr("Prefix")) & "</td>"
		ResponseWrite ("<td>" & CSmartStr("Serial")) & "</td>"
		ResponseWrite ("<td>" & CSmartStr("Amount Received")) & "</td>"
	end if
	ResponseWrite "</tr>"
	Set totals = (CreateDictionary())
	iNumberOfRows = 0
	do while (not bValue(nPageSize) or IsLess(iNumberOfRows,nPageSize)) and bValue(row)
		Set values = (CreateDictionary())
		format = ""
		setArrElement values,"station",GetData(row,"station",format)
		format = ""
		setArrElement values,"name",GetData(row,"name",format)
		format = "Short Date"
		setArrElement values,"payment_date",GetData(row,"payment_date",format)
		format = ""
		setArrElement values,"fop",GetData(row,"fop",format)
		format = ""
		setArrElement values,"prefix",GetData(row,"prefix",format)
		format = ""
		setArrElement values,"serial",GetData(row,"serial",format)
		format = "Number"
		setArrElement values,"amount_received",GetData(row,"amount_received",format)
		eventRes = true
		if bValue(eventObj.exists_p1("BeforeOut")) then
			doAssignmentByRef eventRes,eventObj.BeforeOut_p2(row,values)
		end if
		if bValue(eventRes) then
			iNumberOfRows = CSmartDbl(iNumberOfRows)+1
			ResponseWrite "<tr>"
			if IsEqual(GetRequestValue(Request,"type"),"excel") then
				ResponseWrite "<td x:str>"
			else
				ResponseWrite "<td>"
			end if
			format = ""
			if IsEqual(GetRequestValue(Request,"type"),"excel") then
				ResponseWrite PrepareForExcel(ArrayElement(values,"station"))
			else
				ResponseWrite htmlspecialchars(ArrayElement(values,"station"))
			end if
			ResponseWrite "</td>"
			if IsEqual(GetRequestValue(Request,"type"),"excel") then
				ResponseWrite "<td x:str>"
			else
				ResponseWrite "<td>"
			end if
			format = ""
			if IsEqual(GetRequestValue(Request,"type"),"excel") then
				ResponseWrite PrepareForExcel(ArrayElement(values,"name"))
			else
				ResponseWrite htmlspecialchars(ArrayElement(values,"name"))
			end if
			ResponseWrite "</td>"
			ResponseWrite "<td>"
			format = "Short Date"
			if IsEqual(GetRequestValue(Request,"type"),"excel") then
				ResponseWrite PrepareForExcel(ArrayElement(values,"payment_date"))
			else
				ResponseWrite htmlspecialchars(ArrayElement(values,"payment_date"))
			end if
			ResponseWrite "</td>"
			if IsEqual(GetRequestValue(Request,"type"),"excel") then
				ResponseWrite "<td x:str>"
			else
				ResponseWrite "<td>"
			end if
			format = ""
			if IsEqual(GetRequestValue(Request,"type"),"excel") then
				ResponseWrite PrepareForExcel(ArrayElement(values,"fop"))
			else
				ResponseWrite htmlspecialchars(ArrayElement(values,"fop"))
			end if
			ResponseWrite "</td>"
			if IsEqual(GetRequestValue(Request,"type"),"excel") then
				ResponseWrite "<td x:str>"
			else
				ResponseWrite "<td>"
			end if
			format = ""
			if IsEqual(GetRequestValue(Request,"type"),"excel") then
				ResponseWrite PrepareForExcel(ArrayElement(values,"prefix"))
			else
				ResponseWrite htmlspecialchars(ArrayElement(values,"prefix"))
			end if
			ResponseWrite "</td>"
			if IsEqual(GetRequestValue(Request,"type"),"excel") then
				ResponseWrite "<td x:str>"
			else
				ResponseWrite "<td>"
			end if
			format = ""
			if IsEqual(GetRequestValue(Request,"type"),"excel") then
				ResponseWrite PrepareForExcel(ArrayElement(values,"serial"))
			else
				ResponseWrite htmlspecialchars(ArrayElement(values,"serial"))
			end if
			ResponseWrite "</td>"
			ResponseWrite "<td>"
			format = "Number"
			ResponseWrite htmlspecialchars(ArrayElement(values,"amount_received"))
			ResponseWrite "</td>"
			ResponseWrite "</tr>"
		end if
		if bValue(eventObj.exists_p1("ListFetchArray")) then
			doAssignmentByRef row,eventObj.ListFetchArray_p1(rs)
		else
			doAssignmentByRef row,db_fetch_array(rs)
		end if
	loop
End Function
Function XMLNameEncode(ByVal strValue)
	Dim search,ret
	Set search = (CreateDictionary9(Empty," ",Empty,"#",Empty,"'",Empty,"/",Empty,"\",Empty,"(",Empty,")",Empty,",",Empty,"["))
	doAssignmentByRef ret,asp_str_replace(search,"",strValue)
	Set search = (CreateDictionary9(Empty,"]",Empty,"+",Empty,"""",Empty,"-",Empty,"_",Empty,"|",Empty,"}",Empty,"{",Empty,"="))
	doAssignmentByRef ret,asp_str_replace(search,"",ret)
	doAssignmentByRef XMLNameEncode,ret
	Exit Function
End Function
Function PrepareForExcel(ByVal str)
	Dim ret
	doAssignmentByRef ret,htmlspecialchars(str)
	if IsEqual(asp_substr(ret,0,1),"=") then
		ret = "&#61;" & CSmartStr(asp_substr(ret,1,empty))
	end if
	doAssignmentByRef PrepareForExcel,ret
	Exit Function
End Function
%>
