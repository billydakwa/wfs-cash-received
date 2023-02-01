<%
%>
<!--#include file="include/dbcommon.asp"-->
<%
asp_include "include/xtempl.asp",false
asp_include "classes/runnerpage.asp",false
Set xt = (CreateClass("Xtempl",0,Empty,Empty,Empty,Empty,Empty,Empty,Empty))
doAssignment id,IIF(not IsIdentical(postvalue("id"),""),postvalue("id"),1)
Set params = (CreateDictionary4("pageType",PAGE_MENU,"id",id,"menuTablesArr",menuTablesArr,"isGroupSecurity",isGroupSecurity))
setArrElementByRef params,"xt",xt
setArrElement params,"tName","global"
setArrElement params,"needSearchClauseObj",false
Set pageObject = (CreateClass("RunnerPage",1,params,Empty,Empty,Empty,Empty,Empty,Empty))
if bValue(globalEvents.exists_p1("BeforeProcessMenu")) then
	globalEvents.BeforeProcessMenu_p1 conn
end if
setArrElement pageObject.body,"begin",CSmartStr(ArrayElement(pageObject.body,"begin")) & ("<script type=""text/javascript"" src=""include/jquery.js""></script>" & "<script type=""text/javascript"" src=""include/jsfunctions.js""></script>")
if IsIdentical(pageObject.debugJSMode,true) then
	setArrElement pageObject.body,"begin",CSmartStr(ArrayElement(pageObject.body,"begin")) & "<script language=""JavaScript"" src=""include/runnerJS/RunnerBase.js""></script>" & vbcrlf
else
	setArrElement pageObject.body,"begin",CSmartStr(ArrayElement(pageObject.body,"begin")) & "<script type=""text/javascript"" src=""include/runnerJS/RunnerBase.js""></script>"
end if
pageObject.addCommonJs 
pageObject.fillSetCntrlMaps 
setArrElement pageObject.body,"end",CSmartStr(ArrayElement(pageObject.body,"end")) & "<script>"
setArrElement pageObject.body,"end",CSmartStr(ArrayElement(pageObject.body,"end")) & (("window.controlsMap = '" & CSmartStr(jsreplace(my_json_encode(pageObject.controlsHTMLMap)))) & "';")
setArrElement pageObject.body,"end",CSmartStr(ArrayElement(pageObject.body,"end")) & (("window.settings = '" & CSmartStr(jsreplace(my_json_encode(pageObject.jsSettings)))) & "';")
setArrElement pageObject.body,"end",CSmartStr(ArrayElement(pageObject.body,"end")) & (CSmartStr(pageObject.PrepareJS()) & "</script>")
xt.assignbyref_p2 "body",pageObject.body
doAssignmentByRef menuInfo,pageObject.createOldMenu()
if bValue(pageObject.isCreateMenu()) then
	xt.assign_p2 "menustyle_block",true
end if
if IsLess(ArrayElement(menuInfo,"menuTablesCount"),2) and bValue(asp_strlen(ArrayElement(menuInfo,"urlForRedirect"))) then
	asp_header "Location: " & CSmartStr(ArrayElement(menuInfo,"urlForRedirect"))
	Response.End
end if
templatefile = "menu.htm"
if bValue(globalEvents.exists_p1("BeforeShowMenu")) then
	globalEvents.BeforeShowMenu_p2 xt,templatefile
end if
xt.display_p1 templatefile
%>
