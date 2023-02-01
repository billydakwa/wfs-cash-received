<%
set daltable_dbo_part = server.CreateObject("Scripting.Dictionary")
daltable_dbo_part.Add "code",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("code")("nType")=200
daltable_dbo_part("code")("varname")="code"
daltable_dbo_part.Add "type_list",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("type_list")("nType")=200
daltable_dbo_part("type_list")("varname")="type_list"
daltable_dbo_part.Add "part_id",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("part_id")("nType")=129
daltable_dbo_part("part_id")("varname")="part_id"
daltable_dbo_part.Add "name",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("name")("nType")=200
daltable_dbo_part("name")("varname")="name"
daltable_dbo_part.Add "short_name",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("short_name")("nType")=200
daltable_dbo_part("short_name")("varname")="short_name"
daltable_dbo_part.Add "branch",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("branch")("nType")=129
daltable_dbo_part("branch")("varname")="branch"
daltable_dbo_part.Add "sales_region",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("sales_region")("nType")=200
daltable_dbo_part("sales_region")("varname")="sales_region"
daltable_dbo_part.Add "tel1",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("tel1")("nType")=200
daltable_dbo_part("tel1")("varname")="tel1"
daltable_dbo_part.Add "fax",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("fax")("nType")=200
daltable_dbo_part("fax")("varname")="fax"
daltable_dbo_part.Add "address1",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("address1")("nType")=200
daltable_dbo_part("address1")("varname")="address1"
daltable_dbo_part.Add "address2",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("address2")("nType")=200
daltable_dbo_part("address2")("varname")="address2"
daltable_dbo_part.Add "state",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("state")("nType")=200
daltable_dbo_part("state")("varname")="state"
daltable_dbo_part.Add "zip",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("zip")("nType")=200
daltable_dbo_part("zip")("varname")="zip"
daltable_dbo_part.Add "city",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("city")("nType")=200
daltable_dbo_part("city")("varname")="city"
daltable_dbo_part.Add "country",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("country")("nType")=129
daltable_dbo_part("country")("varname")="country"
daltable_dbo_part.Add "account",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("account")("nType")=200
daltable_dbo_part("account")("varname")="account"
daltable_dbo_part.Add "vat_number",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("vat_number")("nType")=200
daltable_dbo_part("vat_number")("varname")="vat_number"
daltable_dbo_part.Add "customs_number",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("customs_number")("nType")=200
daltable_dbo_part("customs_number")("varname")="customs_number"
daltable_dbo_part.Add "iata_comm_perc",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("iata_comm_perc")("nType")=131
daltable_dbo_part("iata_comm_perc")("varname")="iata_comm_perc"
daltable_dbo_part.Add "is_cass",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("is_cass")("nType")=129
daltable_dbo_part("is_cass")("varname")="is_cass"
daltable_dbo_part.Add "is_import",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("is_import")("nType")=129
daltable_dbo_part("is_import")("varname")="is_import"
daltable_dbo_part.Add "cass_import",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("cass_import")("nType")=200
daltable_dbo_part("cass_import")("varname")="cass_import"
daltable_dbo_part.Add "deal_optimize",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("deal_optimize")("nType")=129
daltable_dbo_part("deal_optimize")("varname")="deal_optimize"
daltable_dbo_part.Add "currency",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("currency")("nType")=129
daltable_dbo_part("currency")("varname")="fldcurrency"
daltable_dbo_part.Add "my_agent",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("my_agent")("nType")=200
daltable_dbo_part("my_agent")("varname")="my_agent"
daltable_dbo_part.Add "awb_seq",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("awb_seq")("nType")=131
daltable_dbo_part("awb_seq")("varname")="awb_seq"
daltable_dbo_part.Add "remarks",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("remarks")("nType")=200
daltable_dbo_part("remarks")("varname")="remarks"
daltable_dbo_part.Add "status",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("status")("nType")=129
daltable_dbo_part("status")("varname")="status"
daltable_dbo_part.Add "address_cont",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("address_cont")("nType")=200
daltable_dbo_part("address_cont")("varname")="address_cont"
daltable_dbo_part.Add "invoice_due_days",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("invoice_due_days")("nType")=131
daltable_dbo_part("invoice_due_days")("varname")="invoice_due_days"
daltable_dbo_part.Add "no_dunning",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("no_dunning")("nType")=129
daltable_dbo_part("no_dunning")("varname")="no_dunning"
daltable_dbo_part.Add "exclude_invoice",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("exclude_invoice")("nType")=129
daltable_dbo_part("exclude_invoice")("varname")="exclude_invoice"
daltable_dbo_part.Add "sales_office",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("sales_office")("nType")=200
daltable_dbo_part("sales_office")("varname")="sales_office"
daltable_dbo_part.Add "sales_area",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("sales_area")("nType")=200
daltable_dbo_part("sales_area")("varname")="sales_area"
daltable_dbo_part.Add "sales_oper",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("sales_oper")("nType")=200
daltable_dbo_part("sales_oper")("varname")="sales_oper"
daltable_dbo_part.Add "name2",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("name2")("nType")=200
daltable_dbo_part("name2")("varname")="name2"
daltable_dbo_part.Add "department",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("department")("nType")=200
daltable_dbo_part("department")("varname")="department"
daltable_dbo_part.Add "security",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("security")("nType")=129
daltable_dbo_part("security")("varname")="security"
daltable_dbo_part.Add "billing_agent",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("billing_agent")("nType")=200
daltable_dbo_part("billing_agent")("varname")="billing_agent"
daltable_dbo_part.Add "no_cc_fee",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("no_cc_fee")("nType")=129
daltable_dbo_part("no_cc_fee")("varname")="no_cc_fee"
daltable_dbo_part.Add "run_flag",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("run_flag")("nType")=129
daltable_dbo_part("run_flag")("varname")="run_flag"
daltable_dbo_part.Add "last_use_date",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("last_use_date")("nType")=135
daltable_dbo_part("last_use_date")("varname")="last_use_date"
daltable_dbo_part.Add "language",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("language")("nType")=200
daltable_dbo_part("language")("varname")="language"
daltable_dbo_part.Add "no_vat",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("no_vat")("nType")=129
daltable_dbo_part("no_vat")("varname")="no_vat"
daltable_dbo_part.Add "flags",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("flags")("nType")=200
daltable_dbo_part("flags")("varname")="flags"
daltable_dbo_part.Add "all_branch_code",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("all_branch_code")("nType")=200
daltable_dbo_part("all_branch_code")("varname")="all_branch_code"
daltable_dbo_part.Add "tax_number",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("tax_number")("nType")=200
daltable_dbo_part("tax_number")("varname")="tax_number"
daltable_dbo_part.Add "oper",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("oper")("nType")=200
daltable_dbo_part("oper")("varname")="oper"
daltable_dbo_part.Add "time_stamp",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("time_stamp")("nType")=135
daltable_dbo_part("time_stamp")("varname")="time_stamp"
daltable_dbo_part.Add "cha_discount",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("cha_discount")("nType")=129
daltable_dbo_part("cha_discount")("varname")="cha_discount"
daltable_dbo_part.Add "consignor",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("consignor")("nType")=200
daltable_dbo_part("consignor")("varname")="consignor"
daltable_dbo_part.Add "valuation_perc",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("valuation_perc")("nType")=131
daltable_dbo_part("valuation_perc")("varname")="valuation_perc"
daltable_dbo_part.Add "client_id",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("client_id")("nType")=129
daltable_dbo_part("client_id")("varname")="client_id"
daltable_dbo_part.Add "firms_code",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("firms_code")("nType")=200
daltable_dbo_part("firms_code")("varname")="firms_code"
daltable_dbo_part.Add "color",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("color")("nType")=129
daltable_dbo_part("color")("varname")="color"
daltable_dbo_part.Add "product_code",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("product_code")("nType")=129
daltable_dbo_part("product_code")("varname")="product_code"
daltable_dbo_part.Add "invoice_limit",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("invoice_limit")("nType")=131
daltable_dbo_part("invoice_limit")("varname")="invoice_limit"
daltable_dbo_part.Add "stock_holder",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("stock_holder")("nType")=129
daltable_dbo_part("stock_holder")("varname")="stock_holder"
daltable_dbo_part.Add "no_arr_print",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("no_arr_print")("nType")=129
daltable_dbo_part("no_arr_print")("varname")="no_arr_print"
daltable_dbo_part.Add "iac_code",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("iac_code")("nType")=200
daltable_dbo_part("iac_code")("varname")="iac_code"
daltable_dbo_part.Add "iac_until_date",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("iac_until_date")("nType")=135
daltable_dbo_part("iac_until_date")("varname")="iac_until_date"
daltable_dbo_part.Add "handle_agent",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("handle_agent")("nType")=200
daltable_dbo_part("handle_agent")("varname")="handle_agent"
daltable_dbo_part.Add "account2",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("account2")("nType")=200
daltable_dbo_part("account2")("varname")="account2"
daltable_dbo_part.Add "po_mand",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("po_mand")("nType")=129
daltable_dbo_part("po_mand")("varname")="po_mand"
daltable_dbo_part.Add "pima_address",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("pima_address")("nType")=200
daltable_dbo_part("pima_address")("varname")="pima_address"
daltable_dbo_part.Add "ccs_address",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("ccs_address")("nType")=200
daltable_dbo_part("ccs_address")("varname")="ccs_address"
daltable_dbo_part.Add "dsip",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("dsip")("nType")=200
daltable_dbo_part("dsip")("varname")="dsip"
daltable_dbo_part.Add "tsa",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("tsa")("nType")=200
daltable_dbo_part("tsa")("varname")="tsa"
daltable_dbo_part.Add "account_oper",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("account_oper")("nType")=200
daltable_dbo_part("account_oper")("varname")="account_oper"
daltable_dbo_part.Add "billing_cycle",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("billing_cycle")("nType")=131
daltable_dbo_part("billing_cycle")("varname")="billing_cycle"
daltable_dbo_part.Add "is_customs_ware",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("is_customs_ware")("nType")=129
daltable_dbo_part("is_customs_ware")("varname")="is_customs_ware"
daltable_dbo_part.Add "invoice_layout",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("invoice_layout")("nType")=129
daltable_dbo_part("invoice_layout")("varname")="invoice_layout"
daltable_dbo_part.Add "is_pay_slip",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("is_pay_slip")("nType")=129
daltable_dbo_part("is_pay_slip")("varname")="is_pay_slip"
daltable_dbo_part.Add "account3",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("account3")("nType")=200
daltable_dbo_part("account3")("varname")="account3"
daltable_dbo_part.Add "is_agent",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("is_agent")("nType")=129
daltable_dbo_part("is_agent")("varname")="is_agent"
daltable_dbo_part.Add "is_shipper",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("is_shipper")("nType")=129
daltable_dbo_part("is_shipper")("varname")="is_shipper"
daltable_dbo_part.Add "is_consignee",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("is_consignee")("nType")=129
daltable_dbo_part("is_consignee")("varname")="is_consignee"
daltable_dbo_part.Add "is_temp",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("is_temp")("nType")=129
daltable_dbo_part("is_temp")("varname")="is_temp"
daltable_dbo_part.Add "my_agent_until_date",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("my_agent_until_date")("nType")=135
daltable_dbo_part("my_agent_until_date")("varname")="my_agent_until_date"
daltable_dbo_part.Add "agent_group",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("agent_group")("nType")=129
daltable_dbo_part("agent_group")("varname")="agent_group"
daltable_dbo_part.Add "invoice_per_item",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("invoice_per_item")("nType")=129
daltable_dbo_part("invoice_per_item")("varname")="invoice_per_item"
daltable_dbo_part.Add "region_location_code",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("region_location_code")("nType")=200
daltable_dbo_part("region_location_code")("varname")="region_location_code"
daltable_dbo_part.Add "cost_center",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("cost_center")("nType")=200
daltable_dbo_part("cost_center")("varname")="cost_center"
daltable_dbo_part.Add "billing_currency",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("billing_currency")("nType")=129
daltable_dbo_part("billing_currency")("varname")="billing_currency"
daltable_dbo_part.Add "part_restrict_flags",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("part_restrict_flags")("nType")=200
daltable_dbo_part("part_restrict_flags")("varname")="part_restrict_flags"
daltable_dbo_part.Add "net_comm_perc",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("net_comm_perc")("nType")=131
daltable_dbo_part("net_comm_perc")("varname")="net_comm_perc"
daltable_dbo_part.Add "cass_branch",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("cass_branch")("nType")=129
daltable_dbo_part("cass_branch")("varname")="cass_branch"
daltable_dbo_part.Add "iata_agent_code",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("iata_agent_code")("nType")=200
daltable_dbo_part("iata_agent_code")("varname")="iata_agent_code"
daltable_dbo_part.Add "agent_type",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("agent_type")("nType")=129
daltable_dbo_part("agent_type")("varname")="agent_type"
daltable_dbo_part.Add "external_id",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("external_id")("nType")=131
daltable_dbo_part("external_id")("varname")="external_id"
daltable_dbo_part.Add "parent_agent",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("parent_agent")("nType")=200
daltable_dbo_part("parent_agent")("varname")="parent_agent"
daltable_dbo_part.Add "efreight_group",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("efreight_group")("nType")=200
daltable_dbo_part("efreight_group")("varname")="efreight_group"
daltable_dbo_part.Add "shipper_invoice",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("shipper_invoice")("nType")=129
daltable_dbo_part("shipper_invoice")("varname")="shipper_invoice"
daltable_dbo_part.Add "mail_part_id",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("mail_part_id")("nType")=200
daltable_dbo_part("mail_part_id")("varname")="mail_part_id"
daltable_dbo_part.Add "mail_oe_group",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("mail_oe_group")("nType")=200
daltable_dbo_part("mail_oe_group")("varname")="mail_oe_group"
daltable_dbo_part.Add "mail_print_inv",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("mail_print_inv")("nType")=129
daltable_dbo_part("mail_print_inv")("varname")="mail_print_inv"
daltable_dbo_part.Add "mail_email_pdf",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("mail_email_pdf")("nType")=129
daltable_dbo_part("mail_email_pdf")("varname")="mail_email_pdf"
daltable_dbo_part.Add "mail_email_xls",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("mail_email_xls")("nType")=129
daltable_dbo_part("mail_email_xls")("varname")="mail_email_xls"
daltable_dbo_part.Add "sap_number",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("sap_number")("nType")=200
daltable_dbo_part("sap_number")("varname")="sap_number"
daltable_dbo_part.Add "gsa_code",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("gsa_code")("nType")=200
daltable_dbo_part("gsa_code")("varname")="gsa_code"
daltable_dbo_part.Add "tax_group",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("tax_group")("nType")=129
daltable_dbo_part("tax_group")("varname")="tax_group"
daltable_dbo_part.Add "creation_date",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("creation_date")("nType")=135
daltable_dbo_part("creation_date")("varname")="creation_date"
daltable_dbo_part.Add "on_invoice",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("on_invoice")("nType")=129
daltable_dbo_part("on_invoice")("varname")="on_invoice"
daltable_dbo_part.Add "formatted_message",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("formatted_message")("nType")=129
daltable_dbo_part("formatted_message")("varname")="formatted_message"
daltable_dbo_part.Add "rakc_id",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("rakc_id")("nType")=200
daltable_dbo_part("rakc_id")("varname")="rakc_id"
daltable_dbo_part.Add "son_flag",server.CreateObject("Scripting.Dictionary")
daltable_dbo_part("son_flag")("nType")=200
daltable_dbo_part("son_flag")("varname")="son_flag"
	dalTable_dbo_part("code")("key")=true
dal_info.Add "part",dalTable_dbo_part
%>