<%
set daltable_dbo_acc_awb = server.CreateObject("Scripting.Dictionary")
daltable_dbo_acc_awb.Add "code",server.CreateObject("Scripting.Dictionary")
daltable_dbo_acc_awb("code")("nType")=129
daltable_dbo_acc_awb("code")("varname")="code"
daltable_dbo_acc_awb.Add "name",server.CreateObject("Scripting.Dictionary")
daltable_dbo_acc_awb("name")("nType")=200
daltable_dbo_acc_awb("name")("varname")="name"
	dalTable_dbo_acc_awb("code")("key")=true
dal_info.Add "acc_awb",dalTable_dbo_acc_awb
%>