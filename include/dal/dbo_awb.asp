<%
set daltable_dbo_awb = server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb.Add "seq",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("seq")("nType")=131
daltable_dbo_awb("seq")("varname")="seq"
daltable_dbo_awb.Add "prefix",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("prefix")("nType")=129
daltable_dbo_awb("prefix")("varname")="prefix"
daltable_dbo_awb.Add "serial",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("serial")("nType")=200
daltable_dbo_awb("serial")("varname")="serial"
daltable_dbo_awb.Add "awb_exp_seq",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("awb_exp_seq")("nType")=131
daltable_dbo_awb("awb_exp_seq")("varname")="awb_exp_seq"
daltable_dbo_awb.Add "issue_date",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("issue_date")("nType")=135
daltable_dbo_awb("issue_date")("varname")="issue_date"
daltable_dbo_awb.Add "issue_place",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("issue_place")("nType")=129
daltable_dbo_awb("issue_place")("varname")="issue_place"
daltable_dbo_awb.Add "issue_carrier",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("issue_carrier")("nType")=200
daltable_dbo_awb("issue_carrier")("varname")="issue_carrier"
daltable_dbo_awb.Add "agent",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("agent")("nType")=200
daltable_dbo_awb("agent")("varname")="agent"
daltable_dbo_awb.Add "service",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("service")("nType")=129
daltable_dbo_awb("service")("varname")="service"
daltable_dbo_awb.Add "origin",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("origin")("nType")=129
daltable_dbo_awb("origin")("varname")="origin"
daltable_dbo_awb.Add "dest",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("dest")("nType")=129
daltable_dbo_awb("dest")("varname")="dest"
daltable_dbo_awb.Add "value_carriage",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("value_carriage")("nType")=131
daltable_dbo_awb("value_carriage")("varname")="value_carriage"
daltable_dbo_awb.Add "value_customs",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("value_customs")("nType")=131
daltable_dbo_awb("value_customs")("varname")="value_customs"
daltable_dbo_awb.Add "insurance",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("insurance")("nType")=131
daltable_dbo_awb("insurance")("varname")="insurance"
daltable_dbo_awb.Add "tax",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("tax")("nType")=131
daltable_dbo_awb("tax")("varname")="tax"
daltable_dbo_awb.Add "volume_code",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("volume_code")("nType")=129
daltable_dbo_awb("volume_code")("varname")="volume_code"
daltable_dbo_awb.Add "volume",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("volume")("nType")=131
daltable_dbo_awb("volume")("varname")="volume"
daltable_dbo_awb.Add "density",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("density")("nType")=129
daltable_dbo_awb("density")("varname")="density"
daltable_dbo_awb.Add "kg_lb",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("kg_lb")("nType")=129
daltable_dbo_awb("kg_lb")("varname")="kg_lb"
daltable_dbo_awb.Add "pieces",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("pieces")("nType")=131
daltable_dbo_awb("pieces")("varname")="pieces"
daltable_dbo_awb.Add "weight",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("weight")("nType")=131
daltable_dbo_awb("weight")("varname")="weight"
daltable_dbo_awb.Add "dim_msr_code",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("dim_msr_code")("nType")=129
daltable_dbo_awb("dim_msr_code")("varname")="dim_msr_code"
daltable_dbo_awb.Add "master_seq",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("master_seq")("nType")=131
daltable_dbo_awb("master_seq")("varname")="master_seq"
daltable_dbo_awb.Add "profit",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("profit")("nType")=131
daltable_dbo_awb("profit")("varname")="profit"
daltable_dbo_awb.Add "customs_status",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("customs_status")("nType")=200
daltable_dbo_awb("customs_status")("varname")="customs_status"
daltable_dbo_awb.Add "arrival_date",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("arrival_date")("nType")=135
daltable_dbo_awb("arrival_date")("varname")="arrival_date"
daltable_dbo_awb.Add "nature_goods",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("nature_goods")("nType")=200
daltable_dbo_awb("nature_goods")("varname")="nature_goods"
daltable_dbo_awb.Add "shipper_reference",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("shipper_reference")("nType")=200
daltable_dbo_awb("shipper_reference")("varname")="shipper_reference"
daltable_dbo_awb.Add "book_reference",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("book_reference")("nType")=200
daltable_dbo_awb("book_reference")("varname")="book_reference"
daltable_dbo_awb.Add "notifier",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("notifier")("nType")=200
daltable_dbo_awb("notifier")("varname")="notifier"
daltable_dbo_awb.Add "import_agent",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("import_agent")("nType")=200
daltable_dbo_awb("import_agent")("varname")="import_agent"
daltable_dbo_awb.Add "import_service",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("import_service")("nType")=129
daltable_dbo_awb("import_service")("varname")="import_service"
daltable_dbo_awb.Add "awb_imp_seq",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("awb_imp_seq")("nType")=131
daltable_dbo_awb("awb_imp_seq")("varname")="awb_imp_seq"
daltable_dbo_awb.Add "time_stamp",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("time_stamp")("nType")=135
daltable_dbo_awb("time_stamp")("varname")="time_stamp"
daltable_dbo_awb.Add "priority",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("priority")("nType")=129
daltable_dbo_awb("priority")("varname")="priority"
daltable_dbo_awb.Add "price_class",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("price_class")("nType")=200
daltable_dbo_awb("price_class")("varname")="price_class"
daltable_dbo_awb.Add "book_status",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("book_status")("nType")=129
daltable_dbo_awb("book_status")("varname")="book_status"
daltable_dbo_awb.Add "service_level",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("service_level")("nType")=200
daltable_dbo_awb("service_level")("varname")="service_level"
daltable_dbo_awb.Add "sales_region",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("sales_region")("nType")=200
daltable_dbo_awb("sales_region")("varname")="sales_region"
daltable_dbo_awb.Add "oper",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("oper")("nType")=200
daltable_dbo_awb("oper")("varname")="oper"
daltable_dbo_awb.Add "remarks",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("remarks")("nType")=200
daltable_dbo_awb("remarks")("varname")="remarks"
daltable_dbo_awb.Add "son_flag",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("son_flag")("nType")=200
daltable_dbo_awb("son_flag")("varname")="son_flag"
daltable_dbo_awb.Add "house_flag",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("house_flag")("nType")=129
daltable_dbo_awb("house_flag")("varname")="house_flag"
daltable_dbo_awb.Add "cha_code",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("cha_code")("nType")=200
daltable_dbo_awb("cha_code")("varname")="cha_code"
daltable_dbo_awb.Add "osi_remarks",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("osi_remarks")("nType")=200
daltable_dbo_awb("osi_remarks")("varname")="osi_remarks"
daltable_dbo_awb.Add "message_status",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("message_status")("nType")=200
daltable_dbo_awb("message_status")("varname")="message_status"
daltable_dbo_awb.Add "pua_code",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("pua_code")("nType")=200
daltable_dbo_awb("pua_code")("varname")="pua_code"
daltable_dbo_awb.Add "book_price_seq",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("book_price_seq")("nType")=131
daltable_dbo_awb("book_price_seq")("varname")="book_price_seq"
daltable_dbo_awb.Add "customs_reference",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("customs_reference")("nType")=200
daltable_dbo_awb("customs_reference")("varname")="customs_reference"
daltable_dbo_awb.Add "shc_list",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("shc_list")("nType")=200
daltable_dbo_awb("shc_list")("varname")="shc_list"
daltable_dbo_awb.Add "net_prorate",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("net_prorate")("nType")=131
daltable_dbo_awb("net_prorate")("varname")="net_prorate"
daltable_dbo_awb.Add "prorate_date",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("prorate_date")("nType")=135
daltable_dbo_awb("prorate_date")("varname")="prorate_date"
daltable_dbo_awb.Add "client_id",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("client_id")("nType")=129
daltable_dbo_awb("client_id")("varname")="client_id"
daltable_dbo_awb.Add "security_flag",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("security_flag")("nType")=129
daltable_dbo_awb("security_flag")("varname")="security_flag"
daltable_dbo_awb.Add "incl_stats",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("incl_stats")("nType")=129
daltable_dbo_awb("incl_stats")("varname")="incl_stats"
daltable_dbo_awb.Add "slc_pieces",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("slc_pieces")("nType")=131
daltable_dbo_awb("slc_pieces")("varname")="slc_pieces"
daltable_dbo_awb.Add "un_number",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("un_number")("nType")=131
daltable_dbo_awb("un_number")("varname")="un_number"
daltable_dbo_awb.Add "book_weight",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("book_weight")("nType")=131
daltable_dbo_awb("book_weight")("varname")="book_weight"
daltable_dbo_awb.Add "book_volume",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("book_volume")("nType")=131
daltable_dbo_awb("book_volume")("varname")="book_volume"
daltable_dbo_awb.Add "export_station",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("export_station")("nType")=129
daltable_dbo_awb("export_station")("varname")="export_station"
daltable_dbo_awb.Add "book_rate",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("book_rate")("nType")=131
daltable_dbo_awb("book_rate")("varname")="book_rate"
daltable_dbo_awb.Add "all_prorate",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("all_prorate")("nType")=131
daltable_dbo_awb("all_prorate")("varname")="all_prorate"
daltable_dbo_awb.Add "departure_date",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("departure_date")("nType")=135
daltable_dbo_awb("departure_date")("varname")="departure_date"
daltable_dbo_awb.Add "x_ray",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("x_ray")("nType")=129
daltable_dbo_awb("x_ray")("varname")="x_ray"
daltable_dbo_awb.Add "rcv_weight",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("rcv_weight")("nType")=131
daltable_dbo_awb("rcv_weight")("varname")="rcv_weight"
daltable_dbo_awb.Add "book_station",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("book_station")("nType")=129
daltable_dbo_awb("book_station")("varname")="book_station"
daltable_dbo_awb.Add "book_office",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("book_office")("nType")=200
daltable_dbo_awb("book_office")("varname")="book_office"
daltable_dbo_awb.Add "dgr_check",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("dgr_check")("nType")=129
daltable_dbo_awb("dgr_check")("varname")="dgr_check"
daltable_dbo_awb.Add "csr_flt_line",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("csr_flt_line")("nType")=131
daltable_dbo_awb("csr_flt_line")("varname")="csr_flt_line"
daltable_dbo_awb.Add "screen_oper",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("screen_oper")("nType")=200
daltable_dbo_awb("screen_oper")("varname")="screen_oper"
daltable_dbo_awb.Add "customs_code",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("customs_code")("nType")=129
daltable_dbo_awb("customs_code")("varname")="customs_code"
daltable_dbo_awb.Add "screen_pieces",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("screen_pieces")("nType")=131
daltable_dbo_awb("screen_pieces")("varname")="screen_pieces"
daltable_dbo_awb.Add "stage_flag",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("stage_flag")("nType")=129
daltable_dbo_awb("stage_flag")("varname")="stage_flag"
daltable_dbo_awb.Add "gsa_comm_perc",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("gsa_comm_perc")("nType")=131
daltable_dbo_awb("gsa_comm_perc")("varname")="gsa_comm_perc"
daltable_dbo_awb.Add "son_flag2",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("son_flag2")("nType")=200
daltable_dbo_awb("son_flag2")("varname")="son_flag2"
daltable_dbo_awb.Add "direct_weight",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("direct_weight")("nType")=131
daltable_dbo_awb("direct_weight")("varname")="direct_weight"
daltable_dbo_awb.Add "security_profile_seq",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("security_profile_seq")("nType")=131
daltable_dbo_awb("security_profile_seq")("varname")="security_profile_seq"
daltable_dbo_awb.Add "part_secure_seq",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("part_secure_seq")("nType")=131
daltable_dbo_awb("part_secure_seq")("varname")="part_secure_seq"
daltable_dbo_awb.Add "rule_code",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("rule_code")("nType")=200
daltable_dbo_awb("rule_code")("varname")="rule_code"
daltable_dbo_awb.Add "balance_flag",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("balance_flag")("nType")=129
daltable_dbo_awb("balance_flag")("varname")="balance_flag"
daltable_dbo_awb.Add "po_ref",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("po_ref")("nType")=200
daltable_dbo_awb("po_ref")("varname")="po_ref"
daltable_dbo_awb.Add "gcg_status_code",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("gcg_status_code")("nType")=131
daltable_dbo_awb("gcg_status_code")("varname")="gcg_status_code"
daltable_dbo_awb.Add "issue_date_log",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("issue_date_log")("nType")=135
daltable_dbo_awb("issue_date_log")("varname")="issue_date_log"
daltable_dbo_awb.Add "issue_timestamp",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("issue_timestamp")("nType")=200
daltable_dbo_awb("issue_timestamp")("varname")="issue_timestamp"
daltable_dbo_awb.Add "serial_prefix",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("serial_prefix")("nType")=200
daltable_dbo_awb("serial_prefix")("varname")="serial_prefix"
daltable_dbo_awb.Add "compatibility_check",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("compatibility_check")("nType")=129
daltable_dbo_awb("compatibility_check")("varname")="compatibility_check"
daltable_dbo_awb.Add "issue_place_text",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("issue_place_text")("nType")=200
daltable_dbo_awb("issue_place_text")("varname")="issue_place_text"
daltable_dbo_awb.Add "signature",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("signature")("nType")=200
daltable_dbo_awb("signature")("varname")="signature"
daltable_dbo_awb.Add "dim_density",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("dim_density")("nType")=131
daltable_dbo_awb("dim_density")("varname")="dim_density"
daltable_dbo_awb.Add "rating_weight",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("rating_weight")("nType")=131
daltable_dbo_awb("rating_weight")("varname")="rating_weight"
daltable_dbo_awb.Add "biling_type",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("biling_type")("nType")=200
daltable_dbo_awb("biling_type")("varname")="biling_type"
daltable_dbo_awb.Add "biling_account",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("biling_account")("nType")=200
daltable_dbo_awb("biling_account")("varname")="biling_account"
daltable_dbo_awb.Add "cit_reason",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("cit_reason")("nType")=200
daltable_dbo_awb("cit_reason")("varname")="cit_reason"
daltable_dbo_awb.Add "iata_note",server.CreateObject("Scripting.Dictionary")
daltable_dbo_awb("iata_note")("nType")=129
daltable_dbo_awb("iata_note")("varname")="iata_note"
	dalTable_dbo_awb("seq")("key")=true
dal_info.Add "awb",dalTable_dbo_awb
%>