<%
pageType = ""
if not IsEmpty(pageObject) then
	doAssignment pageType,pageObject.pageType
end if
Set xt = (CreateClass("Xtempl",0,Empty,Empty,Empty,Empty,Empty,Empty,Empty))
if bValue(asp_array_key_exists("custom1",menuparams)) and IsEqual(ArrayElement(menuparams,"custom1"),"horizontal") then
	xt.assign_p2 "horizontal_typemenu",true
	xt.assign_p2 "tophorizontal_item",true
	xt.assign_p2 "simplehorizontal_item",true
	horiz = true
else
	xt.assign_p2 "vertical_typemenu",true
	xt.assign_p2 "topvertical_item",true
	xt.assign_p2 "simplevertical_item",true
	horiz = false
end if
Set menuNodes = (CreateDictionary())
Set menuNode = (CreateDictionary())
setArrElement menuNode,"name",""
setArrElement menuNode,"nameType","Text"
setArrElement menuNode,"id","1"
setArrElement menuNode,"parent","0"
setArrElement menuNode,"style",""
setArrElement menuNode,"href","mypage.htm"
setArrElement menuNode,"params",""
setArrElement menuNode,"table","payment"
setArrElement menuNode,"type","Leaf"
setArrElement menuNode,"linkType","Internal"
setArrElement menuNode,"pageType","List"
setArrElement menuNode,"openType","None"
setArrElement menuNode,"title","Payment"
setArrElement menuNodes,asp_count(menuNodes),menuNode
nullParent = null
Set rootInfoArr = (CreateDictionary2("id",0,"href",""))
Set menuRoot = (CreateClass("MenuItem",3,rootInfoArr,menuNodes,nullParent,Empty,Empty,Empty,Empty))
menuRoot.setMenuSession 
menuRoot.assignMenuAttrsToTempl_p1 xt
menuRoot.setCurrMenuElem_p1 xt
xt.assign_p2 "mainmenu_block",true
doAssignmentByRef rOrder,xt.getReadingOrder()
Set mainmenu = (CreateDictionary())
if bValue(isEnableSection508()) then
	setArrElement mainmenu,"begin","<a name=""skipmenu""></a>"
end if
setArrElement mainmenu,"end",vbcrlf & _
	"	<script type=""text/javascript"" language=""javascript"" src=""include/jquery.dropshadow.js""></script>"
countLinks = 0
countGroups = 0
GetCollectionBounds menuRoot.children,i_displaymenu_loopIdx2,i_displaymenu_loopMax2
do while i_displaymenu_loopIdx2<=i_displaymenu_loopMax2
	ind = GetCollectionKey(menuRoot.children,i_displaymenu_loopIdx2)
	doAssignment val,ArrayElement(menuRoot.children,ind)
	if bValue(val.showAsLink) then
		countLinks = CSmartDbl(countLinks)+1
	end if
	if bValue(val.showAsGroup) then
		countGroups = CSmartDbl(countGroups)+1
	end if
	i_displaymenu_loopIdx2=i_displaymenu_loopIdx2+1
loop
if (IsEqual(pageType,PAGE_MENU) or IsLess(1,countLinks)) or IsLess(0,countGroups) then
	xt.assignbyref_p2 "mainmenu_block",mainmenu
	xt.assign_p2 "mainmenustyle_block",true
	xt.assign_p2 "mainmenuiestyle_block",true
	if not IsEmpty(ArrayElement(menuparams,"quickjump")) then
		xt.display_p1 "mainmenu_quickjump.htm"
	else
		if bValue(asp_array_key_exists("custom1",menuparams)) and IsEqual(ArrayElement(menuparams,"custom1"),"horizontal") then
			xt.display_p1 "mainmenu_horiz.htm"
		else
			xt.display_p1 "mainmenu.htm"
		end if
	end if
end if
%>
