<%
%>
<!--#include file="include/dbcommon.asp"-->
<%
add_nocache_headers 
asp_include "include/xtempl.asp",false
asp_include "include/payment_variables.asp",false
asp_include "classes/runnerpage.asp",false
asp_include "classes/listpage.asp",false
asp_include "classes/searchpanel.asp",false
asp_include "classes/searchcontrol.asp",false
asp_include "classes/searchclause.asp",false
asp_include "classes/panelsearchcontrol.asp",false
if IsEqual(postvalue("mode"),"") then
	mode = LIST_SIMPLE
	asp_include "classes/listpage_simple.asp",false
	asp_include "classes/searchpanelsimple.asp",false
else
	if IsEqual(postvalue("mode"),"ajax") then
		mode = LIST_AJAX
		asp_include "classes/listpage_simple.asp",false
		asp_include "classes/listpage_ajax.asp",false
		asp_include "classes/searchpanelsimple.asp",false
	else
		if IsEqual(postvalue("mode"),"lookup") then
			asp_include "classes/listpage_embed.asp",false
			asp_include "classes/listpage_lookup.asp",false
			asp_include "classes/searchpanellookup.asp",false
			mode = LIST_LOOKUP
			setArrElement params,"lookupSelectField","prefix"
		else
			if IsEqual(postvalue("mode"),"listdetails") then
				asp_include "classes/listpage_embed.asp",false
				asp_include "classes/listpage_dpinline.asp",false
				mode = LIST_DETAILS
			end if
		end if
	end if
end if
Set xt = (CreateClass("Xtempl",0,Empty,Empty,Empty,Empty,Empty,Empty,Empty))
noBlobReplace = false
if not bValue(noBlobReplace) then
	gQuery.ReplaceFieldsWithDummies_p1 GetBinaryFieldsIndices("")
end if
Set options = (CreateDictionary())
setArrElement options,"pageType",PAGE_LIST
setArrElement options,"id",IIF(postvalue("id"),postvalue("id"),1)
setArrElement options,"mode",mode
setArrElementByRef options,"xt",xt
setArrElement options,"mainMasterPageType",postvalue("mainmasterpagetype")
setArrElement options,"masterPageType",postvalue("masterpagetype")
setArrElement options,"masterTable",postvalue("mastertable")
setArrElement options,"masterId",postvalue("masterid")
setArrElement options,"firstTime",postvalue("firsttime")
i = 1
do while not IsEmpty(GetRequestValue(Request,"masterkey" & CSmartStr(i)))
	setArrElementN options,CreateArray2("masterKeysReq",i),GetRequestValue(Request,"masterkey" & CSmartStr(i))
	i = CSmartDbl(i)+1
loop
doAssignmentByRef pageObject,method_ListPage_createListPage(this_object,strTableName,options)
pageObject.prepareForBuildPage 
Set includesArr = (CreateDictionary())
i = 0
do while IsLess(i,asp_count(includesArr))
	asp_include ArrayElement(includesArr,i),false
	i = CSmartDbl(i)+1
loop
pageObject.showPage 
%>
