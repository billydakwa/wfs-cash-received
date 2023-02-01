<%
set daltable_dbo_payment = server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment.Add "seq",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("seq")("nType")=131
daltable_dbo_payment("seq")("varname")="seq"
daltable_dbo_payment.Add "station",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("station")("nType")=129
daltable_dbo_payment("station")("varname")="station"
daltable_dbo_payment.Add "currency",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("currency")("nType")=129
daltable_dbo_payment("currency")("varname")="fldcurrency"
daltable_dbo_payment.Add "part",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("part")("nType")=200
daltable_dbo_payment("part")("varname")="part"
daltable_dbo_payment.Add "payment_date",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("payment_date")("nType")=135
daltable_dbo_payment("payment_date")("varname")="payment_date"
daltable_dbo_payment.Add "amount",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("amount")("nType")=131
daltable_dbo_payment("amount")("varname")="amount"
daltable_dbo_payment.Add "fop",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("fop")("nType")=129
daltable_dbo_payment("fop")("varname")="fop"
daltable_dbo_payment.Add "fop_code",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("fop_code")("nType")=200
daltable_dbo_payment("fop_code")("varname")="fop_code"
daltable_dbo_payment.Add "fop_number",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("fop_number")("nType")=200
daltable_dbo_payment("fop_number")("varname")="fop_number"
daltable_dbo_payment.Add "confirm_number",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("confirm_number")("nType")=200
daltable_dbo_payment("confirm_number")("varname")="confirm_number"
daltable_dbo_payment.Add "oper",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("oper")("nType")=200
daltable_dbo_payment("oper")("varname")="oper"
daltable_dbo_payment.Add "remark",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("remark")("nType")=200
daltable_dbo_payment("remark")("varname")="remark"
daltable_dbo_payment.Add "fop_date",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("fop_date")("nType")=129
daltable_dbo_payment("fop_date")("varname")="fop_date"
daltable_dbo_payment.Add "fop_ref",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("fop_ref")("nType")=200
daltable_dbo_payment("fop_ref")("varname")="fop_ref"
daltable_dbo_payment.Add "awb_seq",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("awb_seq")("nType")=131
daltable_dbo_payment("awb_seq")("varname")="awb_seq"
daltable_dbo_payment.Add "awb_inv_number",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("awb_inv_number")("nType")=200
daltable_dbo_payment("awb_inv_number")("varname")="awb_inv_number"
daltable_dbo_payment.Add "time_stamp",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("time_stamp")("nType")=135
daltable_dbo_payment("time_stamp")("varname")="time_stamp"
daltable_dbo_payment.Add "exp_imp",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("exp_imp")("nType")=129
daltable_dbo_payment("exp_imp")("varname")="exp_imp"
daltable_dbo_payment.Add "account_flag",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("account_flag")("nType")=200
daltable_dbo_payment("account_flag")("varname")="account_flag"
daltable_dbo_payment.Add "run_seq",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("run_seq")("nType")=131
daltable_dbo_payment("run_seq")("varname")="run_seq"
daltable_dbo_payment.Add "run_inv_number",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("run_inv_number")("nType")=131
daltable_dbo_payment("run_inv_number")("varname")="run_inv_number"
daltable_dbo_payment.Add "inv_pay_seq",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("inv_pay_seq")("nType")=131
daltable_dbo_payment("inv_pay_seq")("varname")="inv_pay_seq"
daltable_dbo_payment.Add "freight_charge",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("freight_charge")("nType")=129
daltable_dbo_payment("freight_charge")("varname")="freight_charge"
daltable_dbo_payment.Add "other_charge",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("other_charge")("nType")=129
daltable_dbo_payment("other_charge")("varname")="other_charge"
daltable_dbo_payment.Add "cass_file_header",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("cass_file_header")("nType")=200
daltable_dbo_payment("cass_file_header")("varname")="cass_file_header"
daltable_dbo_payment.Add "payment_type",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("payment_type")("nType")=129
daltable_dbo_payment("payment_type")("varname")="payment_type"
daltable_dbo_payment.Add "product_inv_seq",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("product_inv_seq")("nType")=131
daltable_dbo_payment("product_inv_seq")("varname")="product_inv_seq"
daltable_dbo_payment.Add "cash_box",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("cash_box")("nType")=200
daltable_dbo_payment("cash_box")("varname")="cash_box"
daltable_dbo_payment.Add "cass_reason_code",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("cass_reason_code")("nType")=129
daltable_dbo_payment("cass_reason_code")("varname")="cass_reason_code"
daltable_dbo_payment.Add "cass_contract_no",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("cass_contract_no")("nType")=200
daltable_dbo_payment("cass_contract_no")("varname")="cass_contract_no"
daltable_dbo_payment.Add "vat_amount",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("vat_amount")("nType")=131
daltable_dbo_payment("vat_amount")("varname")="vat_amount"
daltable_dbo_payment.Add "csr_no",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("csr_no")("nType")=131
daltable_dbo_payment("csr_no")("varname")="csr_no"
daltable_dbo_payment.Add "awb_dlv_seq",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("awb_dlv_seq")("nType")=131
daltable_dbo_payment("awb_dlv_seq")("varname")="awb_dlv_seq"
daltable_dbo_payment.Add "carrier",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("carrier")("nType")=200
daltable_dbo_payment("carrier")("varname")="carrier"
daltable_dbo_payment.Add "state",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("state")("nType")=200
daltable_dbo_payment("state")("varname")="state"
daltable_dbo_payment.Add "driver_license",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("driver_license")("nType")=200
daltable_dbo_payment("driver_license")("varname")="driver_license"
daltable_dbo_payment.Add "cash_box_run",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("cash_box_run")("nType")=129
daltable_dbo_payment("cash_box_run")("varname")="cash_box_run"
daltable_dbo_payment.Add "original_currency",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("original_currency")("nType")=129
daltable_dbo_payment("original_currency")("varname")="original_currency"
daltable_dbo_payment.Add "original_amount",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("original_amount")("nType")=131
daltable_dbo_payment("original_amount")("varname")="original_amount"
daltable_dbo_payment.Add "original_vat_amt",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("original_vat_amt")("nType")=131
daltable_dbo_payment("original_vat_amt")("varname")="original_vat_amt"
daltable_dbo_payment.Add "exch_rate",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("exch_rate")("nType")=5
daltable_dbo_payment("exch_rate")("varname")="exch_rate"
daltable_dbo_payment.Add "son_flag",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("son_flag")("nType")=200
daltable_dbo_payment("son_flag")("varname")="son_flag"
daltable_dbo_payment.Add "awb_exp_seq",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("awb_exp_seq")("nType")=131
daltable_dbo_payment("awb_exp_seq")("varname")="awb_exp_seq"
daltable_dbo_payment.Add "cca_number",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("cca_number")("nType")=131
daltable_dbo_payment("cca_number")("varname")="cca_number"
daltable_dbo_payment.Add "receipt_number",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("receipt_number")("nType")=131
daltable_dbo_payment("receipt_number")("varname")="receipt_number"
daltable_dbo_payment.Add "withold_tax",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("withold_tax")("nType")=131
daltable_dbo_payment("withold_tax")("varname")="withold_tax"
daltable_dbo_payment.Add "amount_received",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("amount_received")("nType")=131
daltable_dbo_payment("amount_received")("varname")="amount_received"
daltable_dbo_payment.Add "incomplete",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("incomplete")("nType")=129
daltable_dbo_payment("incomplete")("varname")="incomplete"
daltable_dbo_payment.Add "credit_account_seq",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("credit_account_seq")("nType")=131
daltable_dbo_payment("credit_account_seq")("varname")="credit_account_seq"
daltable_dbo_payment.Add "bank_code_mand",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("bank_code_mand")("nType")=200
daltable_dbo_payment("bank_code_mand")("varname")="bank_code_mand"
daltable_dbo_payment.Add "grid_row",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("grid_row")("nType")=131
daltable_dbo_payment("grid_row")("varname")="grid_row"
daltable_dbo_payment.Add "balance_remain",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("balance_remain")("nType")=131
daltable_dbo_payment("balance_remain")("varname")="balance_remain"
daltable_dbo_payment.Add "balance_remain_pay",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("balance_remain_pay")("nType")=131
daltable_dbo_payment("balance_remain_pay")("varname")="balance_remain_pay"
daltable_dbo_payment.Add "token_id",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("token_id")("nType")=129
daltable_dbo_payment("token_id")("varname")="token_id"
daltable_dbo_payment.Add "txn_id",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("txn_id")("nType")=131
daltable_dbo_payment("txn_id")("varname")="txn_id"
daltable_dbo_payment.Add "refund_flag",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("refund_flag")("nType")=129
daltable_dbo_payment("refund_flag")("varname")="refund_flag"
daltable_dbo_payment.Add "hash",server.CreateObject("Scripting.Dictionary")
daltable_dbo_payment("hash")("nType")=200
daltable_dbo_payment("hash")("varname")="hash"
	dalTable_dbo_payment("seq")("key")=true
dal_info.Add "payment",dalTable_dbo_payment
%>