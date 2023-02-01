<%

'------ Class Chart ------
Class Chart
	Public strSQL
	Public label2
	Public numRecordsToShow
	Public totalRecords
	Public header
	Public footer
	Public strLabel
	Public arrDataLabels
	Public arrDataSeries
	Public arrDataColor
	Public arrFormatCurrency
	Public arrFormatDecimal
	Public arrFormatCustomer
	Public arrFormatCustomerStr
	Public arrDataSize
	Public arrAxesColor
	Public arrGaugeColor
	Public arrOHLC_high
	Public arrOHLC_low
	Public arrOHLC_open
	Public arrOHLC_close
	Public arrOHLC_candle
	Public arrOHLC_color
	Public arrOHLC_color_up
	Public arrOHLC_color_down
	Public sleg
	Public scol
	Public chrt_array
	Public webchart
	Public cname
	Public gstrOrderBy
	Public table_type
	Public sessionPrefix
	Public Function init_Chart_p2(ByRef ch_array,ByVal param)
		DoAssignmentByRef init_Chart_p2,method_Chart_Chart(me,ch_array,param)
	End Function
	Public Function setFooter_p1(ByVal name)
		DoAssignmentByRef setFooter_p1,method_Chart_setFooter(me,name)
	End Function
	Public Function getFooter()
		DoAssignmentByRef getFooter,method_Chart_getFooter(me)
	End Function
	Public Function setHeader_p1(ByVal name)
		DoAssignmentByRef setHeader_p1,method_Chart_setHeader(me,name)
	End Function
	Public Function getHeader()
		DoAssignmentByRef getHeader,method_Chart_getHeader(me)
	End Function
	Public Function setLabelField_p1(ByVal name)
		DoAssignmentByRef setLabelField_p1,method_Chart_setLabelField(me,name)
	End Function
	Public Function getLabelField()
		DoAssignmentByRef getLabelField,method_Chart_getLabelField(me)
	End Function
	Public Function setSeriaColor_p2(ByVal color,ByVal index)
		DoAssignmentByRef setSeriaColor_p2,method_Chart_setSeriaColor(me,color,index)
	End Function
	Public Function getSeriaColor_p1(ByVal index)
		DoAssignmentByRef getSeriaColor_p1,method_Chart_getSeriaColor(me,index)
	End Function
	Public Function setScrollingState_p1(ByVal scroll)
		DoAssignmentByRef setScrollingState_p1,method_Chart_setScrollingState(me,scroll)
	End Function
	Public Function getScrollingState()
		DoAssignmentByRef getScrollingState,method_Chart_getScrollingState(me)
	End Function
	Public Function setMaxBarScroll_p1(ByVal num)
		DoAssignmentByRef setMaxBarScroll_p1,method_Chart_setMaxBarScroll(me,num)
	End Function
	Public Function getMaxBarScroll()
		DoAssignmentByRef getMaxBarScroll,method_Chart_getMaxBarScroll(me)
	End Function
	Public Function write()
		DoAssignmentByRef write,method_Chart_write(me)
	End Function
	Public Function write_legend()
		DoAssignmentByRef write_legend,method_Chart_write_legend(me)
	End Function
	Public Function write_format()
		DoAssignmentByRef write_format,method_Chart_write_format(me)
	End Function
	Public Function write_data()
		DoAssignmentByRef write_data,method_Chart_write_data(me)
	End Function
	Public Function write_dps()
		DoAssignmentByRef write_dps,method_Chart_write_dps(me)
	End Function
	Public Function write_chart_settings()
		DoAssignmentByRef write_chart_settings,method_Chart_write_chart_settings(me)
	End Function
	Public Function formatCurrency_p2(ByVal val,ByVal series)
		DoAssignmentByRef formatCurrency_p2,method_Chart_formatCurrency(me,val,series)
	End Function
	Public Function write_axes_custom()
		DoAssignmentByRef write_axes_custom,method_Chart_write_axes_custom(me)
	End Function
	Public Function write_Logarithmic()
		DoAssignmentByRef write_Logarithmic,method_Chart_write_Logarithmic(me)
	End Function
	Public Function write_Grid()
		DoAssignmentByRef write_Grid,method_Chart_write_Grid(me)
	End Function
	Public Function write_extra()
		DoAssignmentByRef write_extra,method_Chart_write_extra(me)
	End Function
	Public Function write_chart_background()
		DoAssignmentByRef write_chart_background,method_Chart_write_chart_background(me)
	End Function
	Public Function color_series_p1(ByVal series)
		DoAssignmentByRef color_series_p1,method_Chart_color_series(me,series)
	End Function
	Public Function labelFormat_p2(ByVal fieldName,ByVal data)
		DoAssignmentByRef labelFormat_p2,method_Chart_labelFormat(me,fieldName,data)
	End Function
	Public Function getDefaultValue_p2(ByVal series,ByVal row)
		DoAssignmentByRef getDefaultValue_p2,method_Chart_getDefaultValue(me,series,row)
	End Function
	Public Function valueFormat_p2(ByVal series,ByVal x_axis)
		DoAssignmentByRef valueFormat_p2,method_Chart_valueFormat(me,series,x_axis)
	End Function
	Public Function valueFormat_p1(ByVal series)
		DoAssignmentByRef valueFormat_p1,method_Chart_valueFormat(me,series,false)
	End Function
	Public Function tooltipFormat_p2(ByVal series,ByVal row)
		DoAssignmentByRef tooltipFormat_p2,method_Chart_tooltipFormat(me,series,row)
	End Function
	Public Function get_data_p1(ByVal refr)
		DoAssignmentByRef get_data_p1,method_Chart_get_data(me,refr)
	End Function
	Public Function chart_xmlencode_p1(ByVal str)
		DoAssignmentByRef chart_xmlencode_p1,method_Chart_chart_xmlencode(me,str)
	End Function
	Public Function write_plot_background()
		DoAssignmentByRef write_plot_background,method_Chart_write_plot_background(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"strSQL", strSQL
		setArrElement out,"label2", label2
		setArrElement out,"numRecordsToShow", numRecordsToShow
		setArrElement out,"totalRecords", totalRecords
		setArrElement out,"header", header
		setArrElement out,"footer", footer
		setArrElement out,"strLabel", strLabel
		setArrElement out,"arrDataLabels", arrDataLabels
		setArrElement out,"arrDataSeries", arrDataSeries
		setArrElement out,"arrDataColor", arrDataColor
		setArrElement out,"arrFormatCurrency", arrFormatCurrency
		setArrElement out,"arrFormatDecimal", arrFormatDecimal
		setArrElement out,"arrFormatCustomer", arrFormatCustomer
		setArrElement out,"arrFormatCustomerStr", arrFormatCustomerStr
		setArrElement out,"arrDataSize", arrDataSize
		setArrElement out,"arrAxesColor", arrAxesColor
		setArrElement out,"arrGaugeColor", arrGaugeColor
		setArrElement out,"arrOHLC_high", arrOHLC_high
		setArrElement out,"arrOHLC_low", arrOHLC_low
		setArrElement out,"arrOHLC_open", arrOHLC_open
		setArrElement out,"arrOHLC_close", arrOHLC_close
		setArrElement out,"arrOHLC_candle", arrOHLC_candle
		setArrElement out,"arrOHLC_color", arrOHLC_color
		setArrElement out,"arrOHLC_color_up", arrOHLC_color_up
		setArrElement out,"arrOHLC_color_down", arrOHLC_color_down
		setArrElement out,"sleg", sleg
		setArrElement out,"scol", scol
		setArrElement out,"chrt_array", chrt_array
		setArrElement out,"webchart", webchart
		setArrElement out,"cname", cname
		setArrElement out,"gstrOrderBy", gstrOrderBy
		setArrElement out,"table_type", table_type
		setArrElement out,"sessionPrefix", sessionPrefix
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment strSQL, dict("strSQL")
		DoAssignment label2, dict("label2")
		DoAssignment numRecordsToShow, dict("numRecordsToShow")
		DoAssignment totalRecords, dict("totalRecords")
		DoAssignment header, dict("header")
		DoAssignment footer, dict("footer")
		DoAssignment strLabel, dict("strLabel")
		DoAssignment arrDataLabels, dict("arrDataLabels")
		DoAssignment arrDataSeries, dict("arrDataSeries")
		DoAssignment arrDataColor, dict("arrDataColor")
		DoAssignment arrFormatCurrency, dict("arrFormatCurrency")
		DoAssignment arrFormatDecimal, dict("arrFormatDecimal")
		DoAssignment arrFormatCustomer, dict("arrFormatCustomer")
		DoAssignment arrFormatCustomerStr, dict("arrFormatCustomerStr")
		DoAssignment arrDataSize, dict("arrDataSize")
		DoAssignment arrAxesColor, dict("arrAxesColor")
		DoAssignment arrGaugeColor, dict("arrGaugeColor")
		DoAssignment arrOHLC_high, dict("arrOHLC_high")
		DoAssignment arrOHLC_low, dict("arrOHLC_low")
		DoAssignment arrOHLC_open, dict("arrOHLC_open")
		DoAssignment arrOHLC_close, dict("arrOHLC_close")
		DoAssignment arrOHLC_candle, dict("arrOHLC_candle")
		DoAssignment arrOHLC_color, dict("arrOHLC_color")
		DoAssignment arrOHLC_color_up, dict("arrOHLC_color_up")
		DoAssignment arrOHLC_color_down, dict("arrOHLC_color_down")
		DoAssignment sleg, dict("sleg")
		DoAssignment scol, dict("scol")
		DoAssignment chrt_array, dict("chrt_array")
		DoAssignment webchart, dict("webchart")
		DoAssignment cname, dict("cname")
		DoAssignment gstrOrderBy, dict("gstrOrderBy")
		DoAssignment table_type, dict("table_type")
		DoAssignment sessionPrefix, dict("sessionPrefix")
	End Sub
' end serialize
End Class
' Chart implementation 
Function method_Chart_Chart(byref this_object,ByRef ch_array,ByVal param)
	doClassAssignment this_object,"arrDataLabels",CreateDictionary()
	doClassAssignment this_object,"arrDataSeries",CreateDictionary()
	doClassAssignment this_object,"arrDataColor",CreateDictionary()
	doClassAssignment this_object,"arrFormatCurrency",CreateDictionary()
	doClassAssignment this_object,"arrFormatDecimal",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomer",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomerStr",CreateDictionary()
	doClassAssignment this_object,"arrDataSize",CreateDictionary()
	doClassAssignment this_object,"arrAxesColor",CreateDictionary()
	doClassAssignment this_object,"arrGaugeColor",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_high",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_low",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_open",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_close",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_candle",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_up",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_down",CreateDictionary()
	doClassAssignment this_object,"chrt_array",CreateDictionary()
	this_object.sessionPrefix = ""
	Dim TableName,i,k,beginColor,endColor,gColor,this,j,ind,gQuery,strWhereClause,searchHavingClause,var_SESSION,searchClauseObj,strOrderBy,strSQLbak,tstrSQL,eventObj,sql_query,pos
	if bValue(this_object.webchart) then
		doClassAssignment this_object,"chrt_array",Convert_Old_Chart(ch_array)
	else
		doClassAssignment this_object,"chrt_array",ch_array
	end if
	doClassAssignment this_object,"numRecordsToShow",ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"maxbarscroll")
	if IsLess(this_object.numRecordsToShow,1) then
		this_object.numRecordsToShow = 1
	end if
	doClassAssignment this_object,"table_type",ArrayElement(this_object.chrt_array,"table_type")
	if not bValue(this_object.table_type) then
		this_object.table_type = "project"
	end if
	doClassAssignment this_object,"webchart",ArrayElement(param,"webchart")
	doClassAssignment this_object,"cname",ArrayElement(param,"cname")
	doClassAssignment this_object,"sessionPrefix",ArrayElement(ArrayElement(this_object.chrt_array,"tables"),0)
	doClassAssignment this_object,"gstrOrderBy",ArrayElement(param,"gstrOrderBy")
	doAssignmentByRef TableName,GoodFieldName(ArrayElement(ArrayElement(this_object.chrt_array,"tables"),0))
	doClassAssignment this_object,"header",ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"head")
	doClassAssignment this_object,"footer",ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"foot")
	i = 0
	do while IsLess(i,CSmartDbl(asp_count(ArrayElement(this_object.chrt_array,"parameters")))-1)
		if not IsEmpty(ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"currencyFormat")) then
			setArrElement this_object.arrFormatCurrency,asp_count(this_object.arrFormatCurrency),ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"currencyFormat")
		else
			if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"scur"),"false") then
				setArrElement this_object.arrFormatCurrency,asp_count(this_object.arrFormatCurrency),""
			else
				setArrElement this_object.arrFormatCurrency,asp_count(this_object.arrFormatCurrency),ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"scur")
			end if
		end if
		if not IsEmpty(ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"decimalFormat")) then
			setArrElement this_object.arrFormatDecimal,asp_count(this_object.arrFormatDecimal),ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"decimalFormat")
		else
			setArrElement this_object.arrFormatDecimal,asp_count(this_object.arrFormatDecimal),ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"dec")
		end if
		setArrElement this_object.arrFormatCustomer,asp_count(this_object.arrFormatCustomer),ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"customFormat")
		setArrElement this_object.arrFormatCustomerStr,asp_count(this_object.arrFormatCustomerStr),ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"customFormatStr")
		if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"ohlc") or IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"candlestick") then
			setArrElement this_object.arrOHLC_open,asp_count(this_object.arrOHLC_open),ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"ohlcOpen")
			setArrElement this_object.arrOHLC_high,asp_count(this_object.arrOHLC_high),ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"ohlcHigh")
			setArrElement this_object.arrOHLC_low,asp_count(this_object.arrOHLC_low),ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"ohlcLow")
			setArrElement this_object.arrOHLC_close,asp_count(this_object.arrOHLC_close),ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"ohlcClose")
			setArrElement this_object.arrOHLC_color,asp_count(this_object.arrOHLC_color),"#" & CSmartStr(ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"ohlcColor"))
			if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"candlestick") then
				setArrElement this_object.arrOHLC_candle,asp_count(this_object.arrOHLC_candle),"#" & CSmartStr(ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"ohlcCandleColor"))
			end if
		else
			if not IsEqual(ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"name"),"") then
				setArrElement this_object.arrDataSeries,asp_count(this_object.arrDataSeries),IIF(ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"agr_func"),ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"label"),ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"name"))
				if not IsEmpty(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),("scolor" & CSmartStr(CSmartDbl(i)+1)) & "1")) then
					setArrElement this_object.arrDataColor,asp_count(this_object.arrDataColor),ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),("scolor" & CSmartStr(CSmartDbl(i)+1)) & "1")
				else
					setArrElement this_object.arrDataColor,asp_count(this_object.arrDataColor),ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"series_color")
				end if
				if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"bubble") then
					setArrElement this_object.arrDataSize,asp_count(this_object.arrDataSize),ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"size")
				end if
				if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"gauge") then
					k = 0
					do while bValue(asp_is_array(ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"gaugeColorZone"))) and IsLess(k,asp_count(ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"gaugeColorZone")))
						beginColor = CSmartDbl(ArrayElement(ArrayElement(ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"gaugeColorZone"),k),"gaugeBeginColor"))
						endColor = CSmartDbl(ArrayElement(ArrayElement(ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"gaugeColorZone"),k),"gaugeEndColor"))
						gColor = "#" & CSmartStr(ArrayElement(ArrayElement(ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"gaugeColorZone"),k),"gaugeColor"))
						setArrElementN this_object.arrGaugeColor,CreateArray2(CSmartDbl(asp_count(this_object.arrDataSeries))-1,empty),CreateDictionary3(Empty,beginColor,Empty,endColor,Empty,gColor)
						k = CSmartDbl(k)+1
					loop
				end if
			end if
		end if
		if IsEqual(this_object.table_type,"project") and not bValue(this_object.webchart) then
			setArrElement this_object.arrDataLabels,asp_count(this_object.arrDataLabels),this_object.chart_xmlencode_p1(GetFieldLabel(GoodFieldName(TableName),GoodFieldName(ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"name"))))
		else
			if not bValue(this_object.chart_xmlencode_p1(ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"label"))) then
				setArrElement this_object.arrDataLabels,asp_count(this_object.arrDataLabels),this_object.chart_xmlencode_p1(ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"name"))
			else
				setArrElement this_object.arrDataLabels,asp_count(this_object.arrDataLabels),this_object.chart_xmlencode_p1(ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),i),"label"))
			end if
		end if
		i = CSmartDbl(i)+1
	loop
	if not IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"gauge") then
		doClassAssignment this_object,"strLabel",ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),CSmartDbl(asp_count(ArrayElement(this_object.chrt_array,"parameters")))-1),"name")
		j = 0
		do while IsLess(j,asp_count(ArrayElement(this_object.chrt_array,"fields")))
			if IsEqual(ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),CSmartDbl(asp_count(ArrayElement(this_object.chrt_array,"parameters")))-1),"name"),ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"fields"),j),"name")) then
				if IsEqual(this_object.table_type,"project") then
					doClassAssignment this_object,"label2",this_object.chart_xmlencode_p1(GetFieldLabel(TableName,GoodFieldName(ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),CSmartDbl(asp_count(ArrayElement(this_object.chrt_array,"parameters")))-1),"name"))))
				else
					doClassAssignment this_object,"label2",this_object.chart_xmlencode_p1(ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),CSmartDbl(asp_count(ArrayElement(this_object.chrt_array,"parameters")))-1),"name"))
				end if
			end if
			j = CSmartDbl(j)+1
		loop
	end if
	if not IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"ohlc") and not IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"candlestick") then
		GetCollectionBounds this_object.arrDataColor,c_charts_loopIdx5,c_charts_loopMax5
		do while c_charts_loopIdx5<=c_charts_loopMax5
			ind = GetCollectionKey(this_object.arrDataColor,c_charts_loopIdx5)
			doAssignment val,ArrayElement(this_object.arrDataColor,ind)
			if IsEqual(ind,0) then
				this_object.arrAxesColor = "#000000"
			else
				this_object.arrAxesColor = "#" & CSmartStr(ArrayElement(this_object.arrDataColor,ind))
			end if
			c_charts_loopIdx5=c_charts_loopIdx5+1
		loop
	else
		GetCollectionBounds this_object.arrOHLC_color,c_charts_loopIdx6,c_charts_loopMax6
		do while c_charts_loopIdx6<=c_charts_loopMax6
			ind = GetCollectionKey(this_object.arrOHLC_color,c_charts_loopIdx6)
			doAssignment val,ArrayElement(this_object.arrOHLC_color,ind)
			if IsEqual(ind,0) then
				this_object.arrAxesColor = "#000000"
			else
				this_object.arrAxesColor = "#" & CSmartStr(ArrayElement(this_object.arrOHLC_color,ind))
			end if
			c_charts_loopIdx6=c_charts_loopIdx6+1
		loop
	end if
	doAssignmentByRef gQuery,GetTableData(this_object.sessionPrefix,".sqlquery",null)
	strWhereClause = ""
	searchHavingClause = ""
	if not bValue(this_object.webchart) then
		if not IsEmpty(Session(CSmartStr(this_object.sessionPrefix) & "_advsearch")) then
			doAssignmentByRef searchClauseObj,unserialize(Session(CSmartStr(this_object.sessionPrefix) & "_advsearch"))
			doAssignmentByRef strWhereClause,searchClauseObj.getWhere_p1(GetListOfFieldsByExprType(false,""))
			doAssignmentByRef searchHavingClause,searchClauseObj.getWhere_p1(GetListOfFieldsByExprType(true,""))
		end if
	else
		if not IsEqual(this_object.table_type,"project") then
			strTableName = "webchart" & CSmartStr(this_object.cname)
		else
			doAssignment strTableName,TableName
		end if
		doAssignmentByRef strWhereClause,CalcSearchParam(not IsEqual(this_object.table_type,"project"))
	end if
	if bValue(strWhereClause) then
		setArrElement this_object.chrt_array,"where",CSmartStr(ArrayElement(this_object.chrt_array,"where")) & CSmartStr(IIF(ArrayElement(this_object.chrt_array,"where"),(" AND (" & CSmartStr(strWhereClause)) & ")",(" WHERE (" & CSmartStr(strWhereClause)) & ")"))
	end if
	if IsEqual(this_object.table_type,"project") then
		if bValue(SecuritySQL("Search","")) then
			doAssignmentByRef strWhereClause,whereAdd(strWhereClause,SecuritySQL("Search",""))
		end if
		doClassAssignment this_object,"strSQL",gSQLWhere(strWhereClause,searchHavingClause)
		doAssignment strOrderBy,this_object.gstrOrderBy
		this_object.strSQL = CSmartStr(this_object.strSQL) & (" " & CSmartStr(strOrderBy))
		doAssignment strSQLbak,this_object.strSQL
		if bValue(tableEventExists("BeforeQueryChart",strTableName)) then
			doAssignment tstrSQL,this_object.strSQL
			doAssignmentByRef eventObj,getEventObject(strTableName)
			eventObj.BeforeQueryChart_p3 tstrSQL,strWhereClause,strOrderBy
			doClassAssignment this_object,"strSQL",tstrSQL
		end if
		if IsEqual(strSQLbak,this_object.strSQL) then
			doClassAssignment this_object,"strSQL",gSQLWhere(strWhereClause,searchHavingClause)
			this_object.strSQL = CSmartStr(this_object.strSQL) & (" " & CSmartStr(strOrderBy))
		end if
	end if
	if bValue(this_object.cname) and IsEqual(this_object.table_type,"db") then
		this_object.strSQL = ((CSmartStr(ArrayElement(this_object.chrt_array,"sql")) & CSmartStr(ArrayElement(this_object.chrt_array,"where"))) & CSmartStr(ArrayElement(this_object.chrt_array,"group_by"))) & CSmartStr(ArrayElement(this_object.chrt_array,"order_by"))
	else
		if bValue(this_object.cname) and IsEqual(this_object.table_type,"custom") then
			if not bValue(IsStoredProcedure(ArrayElement(this_object.chrt_array,"sql"))) then
				doAssignment sql_query,ArrayElement(this_object.chrt_array,"sql")
				if IsEqual(GetDatabaseType(),2) then
					doAssignmentByRef pos,asp_strrpos(asp_strtoupper(sql_query),"ORDER BY",empty)
					if bValue(pos) then
						doAssignmentByRef sql_query,asp_substr(sql_query,0,pos)
					end if
				end if
				if not IsEqual(GetDatabaseType(),1) then
					this_object.strSQL = (("select * from (" & CSmartStr(sql_query)) & ") as custom_query") & CSmartStr(ArrayElement(this_object.chrt_array,"where"))
				else
					this_object.strSQL = (("select * from (" & CSmartStr(sql_query)) & ")") & CSmartStr(ArrayElement(this_object.chrt_array,"where"))
				end if
			else
				doClassAssignment this_object,"strSQL",ArrayElement(this_object.chrt_array,"sql")
			end if
		end if
	end if
	if bValue(tableEventExists("UpdateChartSettings",strTableName)) then
		doAssignmentByRef eventObj,getEventObject(strTableName)
		eventObj.UpdateChartSettings_p1 this_object
	end if
End Function
Function method_Chart_setFooter(byref this_object,ByVal name)
	doClassAssignment this_object,"footer",name
End Function
Function method_Chart_getFooter(byref this_object)
	doAssignmentByRef method_Chart_getFooter,this_object.footer
	Exit Function
End Function
Function method_Chart_setHeader(byref this_object,ByVal name)
	doClassAssignment this_object,"header",name
End Function
Function method_Chart_getHeader(byref this_object)
	doAssignmentByRef method_Chart_getHeader,this_object.header
	Exit Function
End Function
Function method_Chart_setLabelField(byref this_object,ByVal name)
	doClassAssignment this_object,"strLabel",name
End Function
Function method_Chart_getLabelField(byref this_object)
	doAssignmentByRef method_Chart_getLabelField,this_object.strLabel
	Exit Function
End Function
Function method_Chart_setSeriaColor(byref this_object,ByVal color,ByVal index)
	setArrElement this_object.arrDataColor,index,color
End Function
Function method_Chart_getSeriaColor(byref this_object,ByVal index)
	doAssignmentByRef method_Chart_getSeriaColor,ArrayElement(this_object.arrDataColor,index)
	Exit Function
End Function
Function method_Chart_setScrollingState(byref this_object,ByVal scroll)
	setArrElementN this_object.chrt_array,CreateArray2("appearance","cscroll"),scroll
End Function
Function method_Chart_getScrollingState(byref this_object)
	method_Chart_getScrollingState = IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"cscroll"),"true")
	Exit Function
End Function
Function method_Chart_setMaxBarScroll(byref this_object,ByVal num)
	doClassAssignment this_object,"numRecordsToShow",num
End Function
Function method_Chart_getMaxBarScroll(byref this_object)
	doAssignmentByRef method_Chart_getMaxBarScroll,this_object.numRecordsToShow
	Exit Function
End Function
Function method_Chart_write(byref this_object)
	Dim this
	ResponseWrite "<?xml version=""1.0"" standalone=""yes""?>" & vblf
	ResponseWrite "<anychart>" & vblf
	ResponseWrite "<settings>" & vblf
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sanim"),"true") and not IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"autoupdate"),"true") then
		ResponseWrite "<animation enabled=""True"" />" & vblf
	else
		ResponseWrite "<animation enabled=""False"" />" & vblf
	end if
	ResponseWrite "</settings>" & vblf
	ResponseWrite "<charts>" & vblf
	this_object.write_data 
	this_object.write_dps 
	this_object.write_chart_settings 
	ResponseWrite "</chart>" & vblf
	ResponseWrite "</charts>" & vblf
	ResponseWrite "</anychart>" & vblf
End Function
Function method_Chart_write_legend(byref this_object)
	Dim this
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"slegend"),"true") then
		this_object.write_legend_tag 
		this_object.write_format 
		ResponseWrite "<template></template>" & vblf
		ResponseWrite "<title enabled=""true"">" & vblf
		ResponseWrite (("<text>" & CSmartStr(this_object.footer)) & "</text>") & vblf
		ResponseWrite (("<font color=""#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color111"))) & """/>") & vblf
		ResponseWrite "</title>" & vblf
		ResponseWrite "<columns_separator enabled=""false""/>" & vblf
		ResponseWrite "<background>" & vblf
		ResponseWrite "<inside_margin left=""10"" right=""10""/>" & vblf
		ResponseWrite "</background>" & vblf
		ResponseWrite "<items>" & vblf
		ResponseWrite (("<item source=""" & CSmartStr(this_object.sleg)) & """/>") & vblf
		ResponseWrite "</items>" & vblf
		ResponseWrite "</legend>" & vblf
	end if
End Function
Function method_Chart_write_format(byref this_object)
	Dim this
	if IsEqual(this_object.sleg,"Points") then
		ResponseWrite (("<format>{%Icon} {%Name} (" & CSmartStr(this_object.valueFormat_p1(0))) & ")</format>") & vblf
	end if
End Function
Function method_Chart_write_data(byref this_object)
End Function
Function method_Chart_write_dps(byref this_object)
End Function
Function method_Chart_write_chart_settings(byref this_object)
	Dim this
	ResponseWrite "<chart_settings>" & vblf
	ResponseWrite "<title enabled=""true"" padding=""15"">" & vblf
	ResponseWrite (("<text>" & CSmartStr(this_object.header)) & "</text>") & vblf
	ResponseWrite (("<font color=""#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color101"))) & """/>") & vblf
	ResponseWrite "</title>" & vblf
	this_object.write_legend 
	this_object.write_axes 
	ResponseWrite "<chart_background>" & vblf
	this_object.write_chart_background 
	ResponseWrite "</chart_background>" & vblf
	this_object.write_plot_background 
	ResponseWrite "</chart_settings>" & vblf
End Function
Function method_Chart_formatCurrency(byref this_object,ByVal val,ByVal series)
	if bValue(ArrayElement(this_object.arrFormatCurrency,series)) then
		do
			If IsEqual(ArrayElement(locale_info,"LOCALE_ICURRENCY"),0) then
				method_Chart_formatCurrency = CSmartStr(ArrayElement(locale_info,"LOCALE_SCURRENCY")) & CSmartStr(val)
				Exit Function
			End If
			If IsEqual(ArrayElement(locale_info,"LOCALE_ICURRENCY"),0) or IsEqual(ArrayElement(locale_info,"LOCALE_ICURRENCY"),1) then
				method_Chart_formatCurrency = CSmartStr(val) & CSmartStr(ArrayElement(locale_info,"LOCALE_SCURRENCY"))
				Exit Function
			End If
			If IsEqual(ArrayElement(locale_info,"LOCALE_ICURRENCY"),0) or IsEqual(ArrayElement(locale_info,"LOCALE_ICURRENCY"),1) or IsEqual(ArrayElement(locale_info,"LOCALE_ICURRENCY"),2) then
				method_Chart_formatCurrency = (CSmartStr(ArrayElement(locale_info,"LOCALE_SCURRENCY")) & " ") & CSmartStr(val)
				Exit Function
			End If
			If IsEqual(ArrayElement(locale_info,"LOCALE_ICURRENCY"),0) or IsEqual(ArrayElement(locale_info,"LOCALE_ICURRENCY"),1) or IsEqual(ArrayElement(locale_info,"LOCALE_ICURRENCY"),2) or IsEqual(ArrayElement(locale_info,"LOCALE_ICURRENCY"),3) then
				method_Chart_formatCurrency = (CSmartStr(val) & " ") & CSmartStr(ArrayElement(locale_info,"LOCALE_SCURRENCY"))
				Exit Function
			End If
		Loop While false
	end if
	doAssignmentByRef method_Chart_formatCurrency,val
	Exit Function
End Function
Function method_Chart_write_axes_custom(byref this_object)
	Dim this,scroll
	ResponseWrite "<axes>" & vblf
	ResponseWrite "<y_axis>" & vblf
	if not IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"saxes"),"true") then
		ResponseWrite (("<line thickness=""1"" color=""DarkColor(#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color141"))) & ")"" caps=""None""/>") & vblf
		ResponseWrite (("<major_tickmark thickness=""1"" color=""DarkColor(#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color141"))) & ")"" caps=""None"" opacity=""1""/>") & vblf
		ResponseWrite (("<minor_tickmark thickness=""1"" color=""DarkColor(#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color141"))) & ")"" caps=""None"" opacity=""1""/>") & vblf
	else
		ResponseWrite (("<line thickness=""1"" color=""DarkColor(" & CSmartStr(ArrayElement(this_object.arrAxesColor,0))) & ")"" caps=""None""/>") & vblf
		ResponseWrite (("<major_tickmark thickness=""1"" color=""DarkColor(" & CSmartStr(ArrayElement(this_object.arrAxesColor,0))) & ")"" caps=""None"" opacity=""1""/>") & vblf
		ResponseWrite (("<minor_tickmark thickness=""1"" color=""DarkColor(" & CSmartStr(ArrayElement(this_object.arrAxesColor,0))) & ")"" caps=""None"" opacity=""1""/>") & vblf
	end if
	ResponseWrite "<title enabled=""true"">" & vblf
	ResponseWrite (("<text>" & CSmartStr(ArrayElement(this_object.arrDataLabels,0))) & "</text>") & vblf
	if not IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"saxes"),"true") then
		ResponseWrite (("<font color=""DarkColor(#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color141"))) & ")""/>") & vblf
	else
		ResponseWrite (("<font color=""DarkColor(" & CSmartStr(ArrayElement(this_object.arrAxesColor,0))) & ")""/>") & vblf
	end if
	ResponseWrite "</title>" & vblf
	ResponseWrite (("<labels enabled=""" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sval"))) & """ align=""Inside"">") & vblf
	ResponseWrite (("<format>" & CSmartStr(this_object.valueFormat_p2(0,true))) & "</format>") & vblf
	if not IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"saxes"),"true") then
		ResponseWrite (("<font color=""#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color61"))) & """ bold=""false"" italic=""false"" underline=""false"" render_as_html=""false"">") & vblf
	else
		ResponseWrite (("<font color=""DarkColor(" & CSmartStr(ArrayElement(this_object.arrAxesColor,0))) & ")"" bold=""false"" italic=""false"" underline=""false"" render_as_html=""false"">") & vblf
	end if
	ResponseWrite "</font>" & vblf
	ResponseWrite "</labels>" & vblf
	this_object.write_Logarithmic 
	this_object.write_Stack 
	this_object.write_Grid 
	ResponseWrite "</y_axis>" & vblf
	this_object.write_get_x_axis 
	ResponseWrite (("<text>" & CSmartStr(this_object.label2)) & "</text>") & vblf
	ResponseWrite (("<font color=""DarkColor(#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color131"))) & ")""/>") & vblf
	ResponseWrite "</title>" & vblf
	scroll = "false"
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"cscroll"),"true") and IsLess(this_object.numRecordsToShow,this_object.totalRecords) then
		scroll = "true"
	end if
	ResponseWrite (((("<zoom enabled=""" & CSmartStr(scroll)) & """ allow_drag=""false"" visible_range=""") & CSmartStr(this_object.numRecordsToShow)) & """/>") & vblf
	ResponseWrite (("<labels enabled=""" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sname"))) & """ display_mode=""normal"">") & vblf
	ResponseWrite (("<font color=""#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color51"))) & """ bold=""false"" italic=""false"" underline=""false"" render_as_html=""false"">") & vblf
	ResponseWrite "</font>" & vblf
	ResponseWrite "<background enabled=""false"">" & vblf
	ResponseWrite "<fill enabled=""false"" />" & vblf
	ResponseWrite "<border enabled=""true"" />" & vblf
	ResponseWrite "</background>" & vblf
	ResponseWrite "</labels>" & vblf
	ResponseWrite "</x_axis>" & vblf
	this_object.write_extra 
	ResponseWrite "</axes>" & vblf
End Function
Function method_Chart_write_Logarithmic(byref this_object)
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"slog"),"true") then
		ResponseWrite "<scale type=""Logarithmic"" log_base=""10""/>" & vblf
	end if
End Function
Function method_Chart_write_Grid(byref this_object)
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sgrid"),"true") then
		ResponseWrite "<major_grid interlaced=""True"">" & vblf
		ResponseWrite (("<line color=""#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color121"))) & """ opacity=""0.7""/>") & vblf
		ResponseWrite "<interlaced_fills>" & vblf
		ResponseWrite (("<even><fill color=""#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color121"))) & """ opacity=""0.1""/></even>") & vblf
		ResponseWrite (("<odd><fill color=""#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color121"))) & """ opacity=""0""/></odd>") & vblf
		ResponseWrite "</interlaced_fills>" & vblf
		ResponseWrite "</major_grid>" & vblf
		ResponseWrite "<minor_grid enabled=""false""/>" & vblf
	end if
End Function
Function method_Chart_write_extra(byref this_object)
	Dim i,position,this
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"saxes"),"true") then
		ResponseWrite "<extra>" & vblf
		i = 1
		do while IsLess(i,asp_count(this_object.arrDataSeries))
			doAssignment position,IIF(i mod 2=0,"Normal","Opposite")
			ResponseWrite (((("<y_axis name=""" & CSmartStr(this_object.chart_xmlencode_p1(ArrayElement(this_object.arrDataSeries,i)))) & """ position=""") & CSmartStr(position)) & """ enabled=""true"">") & vblf
			ResponseWrite (("<line thickness=""1"" color=""DarkColor(" & CSmartStr(ArrayElement(this_object.arrAxesColor,i))) & ")"" caps=""None""/>") & vblf
			ResponseWrite (("<major_tickmark thickness=""1"" color=""DarkColor(" & CSmartStr(ArrayElement(this_object.arrAxesColor,i))) & ")"" opacity=""1""/>") & vblf
			ResponseWrite (("<minor_tickmark thickness=""1"" color=""DarkColor(" & CSmartStr(ArrayElement(this_object.arrAxesColor,i))) & ")"" opacity=""1""/>") & vblf
			ResponseWrite "<minor_grid enabled=""false""/>" & vblf
			ResponseWrite "<major_grid enabled=""false""/>" & vblf
			ResponseWrite "<title enabled=""true"" align=""Center"">" & vblf
			ResponseWrite (("<text>" & CSmartStr(ArrayElement(this_object.arrDataLabels,i))) & "</text>") & vblf
			ResponseWrite (("<font color=""DarkColor(" & CSmartStr(ArrayElement(this_object.arrAxesColor,i))) & ")""/>") & vblf
			ResponseWrite "</title>" & vblf
			ResponseWrite (("<labels align=""Inside"" enabled=""" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sval"))) & """>") & vblf
			ResponseWrite (("<font color=""DarkColor(" & CSmartStr(ArrayElement(this_object.arrAxesColor,i))) & ")""/>") & vblf
			ResponseWrite (("<format>" & CSmartStr(this_object.valueFormat_p2(i,true))) & "</format>") & vblf
			ResponseWrite "</labels>" & vblf
			ResponseWrite "</y_axis>" & vblf
			i = CSmartDbl(i)+1
		loop
		ResponseWrite "</extra>" & vblf
	end if
End Function
Function method_Chart_write_chart_background(byref this_object)
	ResponseWrite "<fill type=""Gradient"">" & vblf
	ResponseWrite "<gradient angle=""90"">" & vblf
	ResponseWrite (("<key position=""0"" color=""#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color71"))) & """/>") & vblf
	if bValue(this_object.webchart) then
		ResponseWrite (("<key position=""1"" color=""#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color71"))) & """ opacity=""0.5""/>") & vblf
	else
		ResponseWrite (("<key position=""1"" color=""#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color81"))) & """ opacity=""0.5""/>") & vblf
	end if
	ResponseWrite "</gradient>" & vblf
	ResponseWrite "</fill>" & vblf
	ResponseWrite "<corners type=""Square""/>" & vblf
	ResponseWrite "<border enabled=""True"" thickness=""2"" type=""Gradient"">" & vblf
	ResponseWrite "<gradient type=""Linear"">" & vblf
	ResponseWrite (("<key position=""0"" color=""#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color91"))) & """ opacity=""0.5"" />") & vblf
	ResponseWrite (("<key position=""1"" color=""DarkColor(#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color91"))) & ")"" opacity=""1"" />") & vblf
	ResponseWrite "</gradient>" & vblf
	ResponseWrite "</border>" & vblf
End Function
Function method_Chart_color_series(byref this_object,ByVal series)
	if IsLess(1,asp_count(this_object.arrDataSeries)) then
		this_object.scol = ("color=""#" & CSmartStr(ArrayElement(this_object.arrDataColor,series))) & """"
		this_object.sleg = "Series"
	else
		this_object.scol = "palette=""Default"""
		this_object.sleg = "Points"
	end if
End Function
Function method_Chart_labelFormat(byref this_object,ByVal fieldName,ByVal data)
	Dim table,strViewFormat,strEditFormat,value,this
	doAssignment table,this_object.sessionPrefix
	doAssignmentByRef strViewFormat,GetFieldData(table,fieldName,"ViewFormat","")
	doAssignmentByRef strEditFormat,GetFieldData(table,fieldName,"EditFormat","")
	if ((IsEqual(strEditFormat,EDIT_FORMAT_LOOKUP_WIZARD) or IsEqual(strEditFormat,EDIT_FORMAT_RADIO)) and IsEqual(GetLookupType(fieldName,table),LT_LOOKUPTABLE)) and not IsEqual(GetLWLinkField(fieldName,table,true),GetLWDisplayField(fieldName,table,true)) then
		doAssignmentByRef value,DisplayLookupWizard(fieldName,ArrayElement(data,fieldName),data,"",MODE_PRINT)
	else
		doAssignmentByRef value,GetData(data,fieldName,strViewFormat)
	end if
	if IsLess(50,asp_strlen(value)) then
		value = CSmartStr(asp_substr(value,0,47)) & "..."
	end if
	doAssignmentByRef method_Chart_labelFormat,this_object.chart_xmlencode_p1(value)
	Exit Function
End Function
Function method_Chart_getDefaultValue(byref this_object,ByVal series,ByVal row)
	Dim res,this
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"ohlc") or IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"candlestick") then
		res = ("O: " & CSmartStr(this_object.formatCurrency_p2(("{%Open}{numDecimals:" & CSmartStr(ArrayElement(this_object.arrFormatDecimal,series))) & "}",series))) & vbcr
		res = CSmartStr(res) & (("H: " & CSmartStr(this_object.formatCurrency_p2(("{%High}{numDecimals:" & CSmartStr(ArrayElement(this_object.arrFormatDecimal,series))) & "}",series))) & vbcr)
		res = CSmartStr(res) & (("L: " & CSmartStr(this_object.formatCurrency_p2(("{%Low}{numDecimals:" & CSmartStr(ArrayElement(this_object.arrFormatDecimal,series))) & "}",series))) & vbcr)
		res = CSmartStr(res) & (("C: " & CSmartStr(this_object.formatCurrency_p2(("{%Close}{numDecimals:" & CSmartStr(ArrayElement(this_object.arrFormatDecimal,series))) & "}",series))) & vbcr)
	else
		if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"bubble") then
			res = "Series: {%SeriesName}" & vbcr
			res = CSmartStr(res) & ((("Point Name: " & CSmartStr(this_object.labelFormat_p2(this_object.strLabel,row))) & "") & vbcr)
			res = CSmartStr(res) & (("Value: " & CSmartStr(this_object.formatCurrency_p2(("{%Value}{numDecimals:" & CSmartStr(ArrayElement(this_object.arrFormatDecimal,series))) & "}",series))) & vbcr)
			res = CSmartStr(res) & ("Bubble Size: {%BubbleSize}" & vbcr)
		else
			if ((IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"2d_bar") or IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"2d_column")) or IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"area")) or IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"funnel") then
				res = (CSmartStr(this_object.labelFormat_p2(this_object.strLabel,row)) & " - ") & CSmartStr(this_object.formatCurrency_p2(("{%YValue}{numDecimals:" & CSmartStr(ArrayElement(this_object.arrFormatDecimal,series))) & "}",series))
			else
				if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"line") or IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"combined") then
					res = "Series: {%SeriesName}" & vbcr
					res = CSmartStr(res) & (("Point Name: " & CSmartStr(this_object.labelFormat_p2(this_object.strLabel,row))) & vbcr)
					res = CSmartStr(res) & (("Value: " & CSmartStr(this_object.formatCurrency_p2(("{%Value}{numDecimals:" & CSmartStr(ArrayElement(this_object.arrFormatDecimal,series))) & "}",series))) & vbcr)
				else
					if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"2d_pie") or IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"2d_doughnut") then
						res = ((CSmartStr(this_object.labelFormat_p2(this_object.strLabel,row)) & " - ") & CSmartStr(this_object.formatCurrency_p2(("{%YValue}{numDecimals:" & CSmartStr(ArrayElement(this_object.arrFormatDecimal,series))) & "}",series))) & vbcr
						res = CSmartStr(res) & (("Percent: {%YPercentOfSeries}{numDecimals:" & CSmartStr(ArrayElement(this_object.arrFormatDecimal,series))) & "}%")
					else
						doAssignmentByRef res,this_object.formatCurrency_p2(("{%YValue}{numDecimals:" & CSmartStr(ArrayElement(this_object.arrFormatDecimal,series))) & "}",series)
					end if
				end if
			end if
		end if
	end if
	doAssignmentByRef method_Chart_getDefaultValue,res
	Exit Function
End Function
Function method_Chart_valueFormat(byref this_object,ByVal series,ByVal x_axis)
	Dim value,this
	if not bValue(ArrayElement(this_object.arrFormatCustomer,series)) then
		if bValue(x_axis) then
			doAssignmentByRef value,this_object.formatCurrency_p2(("{%Value}{numDecimals:" & CSmartStr(ArrayElement(this_object.arrFormatDecimal,series))) & "}",series)
		else
			doAssignmentByRef value,this_object.formatCurrency_p2(("{%YValue}{numDecimals:" & CSmartStr(ArrayElement(this_object.arrFormatDecimal,series))) & "}",series)
		end if
	else
		doAssignment value,ArrayElement(this_object.arrFormatCustomerStr,series)
	end if
	doAssignmentByRef method_Chart_valueFormat,value
	Exit Function
End Function
Function method_Chart_tooltipFormat(byref this_object,ByVal series,ByVal row)
	Dim value,this
	if not bValue(ArrayElement(this_object.arrFormatCustomer,series)) then
		doAssignmentByRef value,this_object.getDefaultValue_p2(series,row)
	else
		doAssignment value,ArrayElement(this_object.arrFormatCustomerStr,series)
	end if
	doAssignmentByRef method_Chart_tooltipFormat,value
	Exit Function
End Function
Function method_Chart_get_data(byref this_object,ByVal refr)
	Dim arrSer,i,this,rs,j,recPerRow,row
	Set arrSer = (CreateDictionary())
	i = 0
	do while IsLess(i,asp_count(this_object.arrDataSeries))
		this_object.color_series_p1 i
		setArrElement arrSer,"series" & CSmartStr(i),(((((((("<series id= """ & CSmartStr(this_object.chart_xmlencode_p1(ArrayElement(this_object.arrDataSeries,i)))) & """ name=""") & CSmartStr(ArrayElement(this_object.arrDataLabels,i))) & """ ") & CSmartStr(this_object.scol)) & " ") & CSmartStr(IIF(IsEqual(i,0),"",(" y_axis=""" & CSmartStr(this_object.chart_xmlencode_p1(ArrayElement(this_object.arrDataSeries,i)))) & """"))) & ">") & vblf
		if (not IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"2d_pie") and not IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"2d_doughnut")) and not IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"funnel") then
			setArrElement arrSer,"series" & CSmartStr(i),CSmartStr(ArrayElement(arrSer,"series" & CSmartStr(i))) & ((((("<label enabled=""" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sval"))) & """><format>") & CSmartStr(this_object.valueFormat_p1(i))) & "</format></label>") & vblf)
		end if
		i = CSmartDbl(i)+1
	loop
	doAssignmentByRef rs,db_query(this_object.strSQL,conn)
	j = 0
	doAssignment recPerRow,this_object.numRecordsToShow
	do while bValue(doAssignmentByRef(row,db_fetch_array(rs)))
		j = CSmartDbl(j)+1
		if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"cscroll"),"true") then
			recPerRow = CSmartDbl(recPerRow)+1
		end if
		if IsLess(recPerRow,j) and IsLess(0,recPerRow) then
			exit do
		end if
		i = 0
		do while IsLess(i,asp_count(this_object.arrDataSeries))
			setArrElement arrSer,"series" & CSmartStr(i),CSmartStr(ArrayElement(arrSer,"series" & CSmartStr(i))) & (CSmartStr(this_object.get_point_p2(i,row)) & vblf)
			i = CSmartDbl(i)+1
		loop
	loop
	doClassAssignment this_object,"totalRecords",j
	i = 0
	do while IsLess(i,asp_count(this_object.arrDataSeries))
		if bValue(refr) then
			ResponseWrite CSmartStr(ArrayElement(this_object.arrDataSeries,i)) & vblf
			setArrElement arrSer,"series" & CSmartStr(i),asp_str_replace(CreateDictionary2(Empty,"\",Empty,vblf),CreateDictionary2(Empty,"\\",Empty,"\n"),ArrayElement(arrSer,"series" & CSmartStr(i)))
		end if
		if IsLess(0,j) then
			ResponseWrite CSmartStr(ArrayElement(arrSer,"series" & CSmartStr(i))) & "</series>"
		end if
		if not bValue(refr) or IsLess(i,CSmartDbl(asp_count(this_object.arrDataSeries))-1) then
			ResponseWrite vblf
		end if
		i = CSmartDbl(i)+1
	loop
	db_close conn
End Function
Function method_Chart_chart_xmlencode(byref this_object,ByVal str)
	doAssignmentByRef method_Chart_chart_xmlencode,asp_str_replace(CreateDictionary4(Empty,"&",Empty,"<",Empty,">",Empty,""""),CreateDictionary4(Empty,"&amp;",Empty,"&lt;",Empty,"&gt;",Empty,"&quot;"),str)
	Exit Function
End Function
Function method_Chart_write_plot_background(byref this_object)
	if not IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color81"),"") then
		ResponseWrite "<data_plot_background>" & vblf
		if bValue(this_object.webchart) then
			ResponseWrite ("<fill enabled=""true"" type=""Solid"" color=""#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color81"))) & """ opacity=""1""/>"
		else
			ResponseWrite "<fill opacity=""0.3""/>" & vblf
		end if
		ResponseWrite "</data_plot_background>" & vblf
	end if
End Function

'------ Class Chart_Bar extends Chart ------
Class Chart_Bar
	Public stacked
	Public var_2d
	Public bar
	Public strSQL
	Public label2
	Public numRecordsToShow
	Public totalRecords
	Public header
	Public footer
	Public strLabel
	Public arrDataLabels
	Public arrDataSeries
	Public arrDataColor
	Public arrFormatCurrency
	Public arrFormatDecimal
	Public arrFormatCustomer
	Public arrFormatCustomerStr
	Public arrDataSize
	Public arrAxesColor
	Public arrGaugeColor
	Public arrOHLC_high
	Public arrOHLC_low
	Public arrOHLC_open
	Public arrOHLC_close
	Public arrOHLC_candle
	Public arrOHLC_color
	Public arrOHLC_color_up
	Public arrOHLC_color_down
	Public sleg
	Public scol
	Public chrt_array
	Public webchart
	Public cname
	Public gstrOrderBy
	Public table_type
	Public sessionPrefix
	Public Function init_Chart_Bar_p2(ByRef ch_array,ByVal param)
		DoAssignmentByRef init_Chart_Bar_p2,method_Chart_Bar_Chart_Bar(me,ch_array,param)
	End Function
	Public Function write_data()
		DoAssignmentByRef write_data,method_Chart_Bar_write_data(me)
	End Function
	Public Function write_dps()
		DoAssignmentByRef write_dps,method_Chart_Bar_write_dps(me)
	End Function
	Public Function write_get_x_axis()
		DoAssignmentByRef write_get_x_axis,method_Chart_Bar_write_get_x_axis(me)
	End Function
	Public Function plot_type_name()
		DoAssignmentByRef plot_type_name,method_Chart_Bar_plot_type_name(me)
	End Function
	Public Function series_3d_mode()
		DoAssignmentByRef series_3d_mode,method_Chart_Bar_series_3d_mode(me)
	End Function
	Public Function chart_style_type()
		DoAssignmentByRef chart_style_type,method_Chart_Bar_chart_style_type(me)
	End Function
	Public Function write_Stack()
		DoAssignmentByRef write_Stack,method_Chart_Bar_write_Stack(me)
	End Function
	Public Function write_label_settings()
		DoAssignmentByRef write_label_settings,method_Chart_Bar_write_label_settings(me)
	End Function
	Public Function get_point_p2(ByVal series,ByVal row)
		DoAssignmentByRef get_point_p2,method_Chart_Bar_get_point(me,series,row)
	End Function
	Public Function write_axes()
		DoAssignmentByRef write_axes,method_Chart_Bar_write_axes(me)
	End Function
	Public Function write_legend_tag()
		DoAssignmentByRef write_legend_tag,method_Chart_Bar_write_legend_tag(me)
	End Function
	Public Function setFooter_p1(ByVal name)
		DoAssignmentByRef setFooter_p1,method_Chart_setFooter(me,name)
	End Function
	Public Function getFooter()
		DoAssignmentByRef getFooter,method_Chart_getFooter(me)
	End Function
	Public Function setHeader_p1(ByVal name)
		DoAssignmentByRef setHeader_p1,method_Chart_setHeader(me,name)
	End Function
	Public Function getHeader()
		DoAssignmentByRef getHeader,method_Chart_getHeader(me)
	End Function
	Public Function setLabelField_p1(ByVal name)
		DoAssignmentByRef setLabelField_p1,method_Chart_setLabelField(me,name)
	End Function
	Public Function getLabelField()
		DoAssignmentByRef getLabelField,method_Chart_getLabelField(me)
	End Function
	Public Function setSeriaColor_p2(ByVal color,ByVal index)
		DoAssignmentByRef setSeriaColor_p2,method_Chart_setSeriaColor(me,color,index)
	End Function
	Public Function getSeriaColor_p1(ByVal index)
		DoAssignmentByRef getSeriaColor_p1,method_Chart_getSeriaColor(me,index)
	End Function
	Public Function setScrollingState_p1(ByVal scroll)
		DoAssignmentByRef setScrollingState_p1,method_Chart_setScrollingState(me,scroll)
	End Function
	Public Function getScrollingState()
		DoAssignmentByRef getScrollingState,method_Chart_getScrollingState(me)
	End Function
	Public Function setMaxBarScroll_p1(ByVal num)
		DoAssignmentByRef setMaxBarScroll_p1,method_Chart_setMaxBarScroll(me,num)
	End Function
	Public Function getMaxBarScroll()
		DoAssignmentByRef getMaxBarScroll,method_Chart_getMaxBarScroll(me)
	End Function
	Public Function write()
		DoAssignmentByRef write,method_Chart_write(me)
	End Function
	Public Function write_legend()
		DoAssignmentByRef write_legend,method_Chart_write_legend(me)
	End Function
	Public Function write_format()
		DoAssignmentByRef write_format,method_Chart_write_format(me)
	End Function
	Public Function write_chart_settings()
		DoAssignmentByRef write_chart_settings,method_Chart_write_chart_settings(me)
	End Function
	Public Function formatCurrency_p2(ByVal val,ByVal series)
		DoAssignmentByRef formatCurrency_p2,method_Chart_formatCurrency(me,val,series)
	End Function
	Public Function write_axes_custom()
		DoAssignmentByRef write_axes_custom,method_Chart_write_axes_custom(me)
	End Function
	Public Function write_Logarithmic()
		DoAssignmentByRef write_Logarithmic,method_Chart_write_Logarithmic(me)
	End Function
	Public Function write_Grid()
		DoAssignmentByRef write_Grid,method_Chart_write_Grid(me)
	End Function
	Public Function write_extra()
		DoAssignmentByRef write_extra,method_Chart_write_extra(me)
	End Function
	Public Function write_chart_background()
		DoAssignmentByRef write_chart_background,method_Chart_write_chart_background(me)
	End Function
	Public Function color_series_p1(ByVal series)
		DoAssignmentByRef color_series_p1,method_Chart_color_series(me,series)
	End Function
	Public Function labelFormat_p2(ByVal fieldName,ByVal data)
		DoAssignmentByRef labelFormat_p2,method_Chart_labelFormat(me,fieldName,data)
	End Function
	Public Function getDefaultValue_p2(ByVal series,ByVal row)
		DoAssignmentByRef getDefaultValue_p2,method_Chart_getDefaultValue(me,series,row)
	End Function
	Public Function valueFormat_p2(ByVal series,ByVal x_axis)
		DoAssignmentByRef valueFormat_p2,method_Chart_valueFormat(me,series,x_axis)
	End Function
	Public Function valueFormat_p1(ByVal series)
		DoAssignmentByRef valueFormat_p1,method_Chart_valueFormat(me,series,false)
	End Function
	Public Function tooltipFormat_p2(ByVal series,ByVal row)
		DoAssignmentByRef tooltipFormat_p2,method_Chart_tooltipFormat(me,series,row)
	End Function
	Public Function get_data_p1(ByVal refr)
		DoAssignmentByRef get_data_p1,method_Chart_get_data(me,refr)
	End Function
	Public Function chart_xmlencode_p1(ByVal str)
		DoAssignmentByRef chart_xmlencode_p1,method_Chart_chart_xmlencode(me,str)
	End Function
	Public Function write_plot_background()
		DoAssignmentByRef write_plot_background,method_Chart_write_plot_background(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"stacked", stacked
		setArrElement out,"var_2d", var_2d
		setArrElement out,"bar", bar
		setArrElement out,"strSQL", strSQL
		setArrElement out,"label2", label2
		setArrElement out,"numRecordsToShow", numRecordsToShow
		setArrElement out,"totalRecords", totalRecords
		setArrElement out,"header", header
		setArrElement out,"footer", footer
		setArrElement out,"strLabel", strLabel
		setArrElement out,"arrDataLabels", arrDataLabels
		setArrElement out,"arrDataSeries", arrDataSeries
		setArrElement out,"arrDataColor", arrDataColor
		setArrElement out,"arrFormatCurrency", arrFormatCurrency
		setArrElement out,"arrFormatDecimal", arrFormatDecimal
		setArrElement out,"arrFormatCustomer", arrFormatCustomer
		setArrElement out,"arrFormatCustomerStr", arrFormatCustomerStr
		setArrElement out,"arrDataSize", arrDataSize
		setArrElement out,"arrAxesColor", arrAxesColor
		setArrElement out,"arrGaugeColor", arrGaugeColor
		setArrElement out,"arrOHLC_high", arrOHLC_high
		setArrElement out,"arrOHLC_low", arrOHLC_low
		setArrElement out,"arrOHLC_open", arrOHLC_open
		setArrElement out,"arrOHLC_close", arrOHLC_close
		setArrElement out,"arrOHLC_candle", arrOHLC_candle
		setArrElement out,"arrOHLC_color", arrOHLC_color
		setArrElement out,"arrOHLC_color_up", arrOHLC_color_up
		setArrElement out,"arrOHLC_color_down", arrOHLC_color_down
		setArrElement out,"sleg", sleg
		setArrElement out,"scol", scol
		setArrElement out,"chrt_array", chrt_array
		setArrElement out,"webchart", webchart
		setArrElement out,"cname", cname
		setArrElement out,"gstrOrderBy", gstrOrderBy
		setArrElement out,"table_type", table_type
		setArrElement out,"sessionPrefix", sessionPrefix
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment stacked, dict("stacked")
		DoAssignment var_2d, dict("var_2d")
		DoAssignment bar, dict("bar")
		DoAssignment strSQL, dict("strSQL")
		DoAssignment label2, dict("label2")
		DoAssignment numRecordsToShow, dict("numRecordsToShow")
		DoAssignment totalRecords, dict("totalRecords")
		DoAssignment header, dict("header")
		DoAssignment footer, dict("footer")
		DoAssignment strLabel, dict("strLabel")
		DoAssignment arrDataLabels, dict("arrDataLabels")
		DoAssignment arrDataSeries, dict("arrDataSeries")
		DoAssignment arrDataColor, dict("arrDataColor")
		DoAssignment arrFormatCurrency, dict("arrFormatCurrency")
		DoAssignment arrFormatDecimal, dict("arrFormatDecimal")
		DoAssignment arrFormatCustomer, dict("arrFormatCustomer")
		DoAssignment arrFormatCustomerStr, dict("arrFormatCustomerStr")
		DoAssignment arrDataSize, dict("arrDataSize")
		DoAssignment arrAxesColor, dict("arrAxesColor")
		DoAssignment arrGaugeColor, dict("arrGaugeColor")
		DoAssignment arrOHLC_high, dict("arrOHLC_high")
		DoAssignment arrOHLC_low, dict("arrOHLC_low")
		DoAssignment arrOHLC_open, dict("arrOHLC_open")
		DoAssignment arrOHLC_close, dict("arrOHLC_close")
		DoAssignment arrOHLC_candle, dict("arrOHLC_candle")
		DoAssignment arrOHLC_color, dict("arrOHLC_color")
		DoAssignment arrOHLC_color_up, dict("arrOHLC_color_up")
		DoAssignment arrOHLC_color_down, dict("arrOHLC_color_down")
		DoAssignment sleg, dict("sleg")
		DoAssignment scol, dict("scol")
		DoAssignment chrt_array, dict("chrt_array")
		DoAssignment webchart, dict("webchart")
		DoAssignment cname, dict("cname")
		DoAssignment gstrOrderBy, dict("gstrOrderBy")
		DoAssignment table_type, dict("table_type")
		DoAssignment sessionPrefix, dict("sessionPrefix")
	End Sub
' end serialize
End Class
' Chart_Bar implementation 
Function method_Chart_Bar_Chart_Bar(byref this_object,ByRef ch_array,ByVal param)
	doClassAssignment this_object,"arrDataLabels",CreateDictionary()
	doClassAssignment this_object,"arrDataSeries",CreateDictionary()
	doClassAssignment this_object,"arrDataColor",CreateDictionary()
	doClassAssignment this_object,"arrFormatCurrency",CreateDictionary()
	doClassAssignment this_object,"arrFormatDecimal",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomer",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomerStr",CreateDictionary()
	doClassAssignment this_object,"arrDataSize",CreateDictionary()
	doClassAssignment this_object,"arrAxesColor",CreateDictionary()
	doClassAssignment this_object,"arrGaugeColor",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_high",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_low",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_open",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_close",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_candle",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_up",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_down",CreateDictionary()
	doClassAssignment this_object,"chrt_array",CreateDictionary()
	this_object.sessionPrefix = ""
	doClassAssignment this_object,"stacked",ArrayElement(param,"stacked")
	doClassAssignment this_object,"var_2d",ArrayElement(param,"2d")
	doClassAssignment this_object,"bar",ArrayElement(param,"bar")
	method_Chart_Chart this_object,ch_array,param
End Function
Function method_Chart_Bar_write_data(byref this_object)
	Dim this
	ResponseWrite (("<chart plot_type=""" & CSmartStr(this_object.plot_type_name())) & """>") & vblf
	ResponseWrite "<data>" & vblf
	this_object.get_data_p1 false
	ResponseWrite "</data>" & vblf
End Function
Function method_Chart_Bar_write_dps(byref this_object)
	Dim this
	ResponseWrite (("<data_plot_settings default_series_type=""Bar""" & CSmartStr(this_object.series_3d_mode())) & ">") & vblf
	ResponseWrite (("<bar_series group_padding=""0.5"" " & CSmartStr(this_object.chart_style_type())) & ">") & vblf
	ResponseWrite this_object.write_label_settings()
	ResponseWrite "</bar_series>" & vblf
	ResponseWrite "</data_plot_settings>" & vblf
End Function
Function method_Chart_Bar_write_get_x_axis(byref this_object)
	ResponseWrite "<x_axis>" & vblf
	ResponseWrite (("<line thickness=""1"" color=""DarkColor(#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color131"))) & ")"" caps=""None""/>") & vblf
	ResponseWrite (("<major_tickmark thickness=""1"" color=""DarkColor(#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color131"))) & ")"" caps=""None"" opacity=""1""/>") & vblf
	ResponseWrite (("<minor_tickmark thickness=""1"" color=""DarkColor(#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color131"))) & ")"" caps=""None"" opacity=""1""/>") & vblf
	ResponseWrite "<title enabled=""true"" align=""Center"">" & vblf
End Function
Function method_Chart_Bar_plot_type_name(byref this_object)
	if not bValue(this_object.bar) then
		method_Chart_Bar_plot_type_name = "CategorizedVertical"
		Exit Function
	else
		method_Chart_Bar_plot_type_name = "CategorizedHorizontal"
		Exit Function
	end if
End Function
Function method_Chart_Bar_series_3d_mode(byref this_object)
	Dim str
	str = ""
	if not bValue(this_object.var_2d) then
		str = " enable_3d_mode=""True"""
		if bValue(this_object.bar) then
			str = CSmartStr(str) & " z_aspect=""1.1"""
		end if
	end if
	doAssignmentByRef method_Chart_Bar_series_3d_mode,str
	Exit Function
End Function
Function method_Chart_Bar_chart_style_type(byref this_object)
	Dim str
	str = ""
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"aqua"),1) then
		str = " style=""AquaLight"""
	else
		if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"aqua"),2) then
			str = " style=""AquaDark"""
		end if
	end if
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"cview"),1) then
		str = CSmartStr(str) & " shape_type=""Cone"""
	else
		if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"cview"),2) then
			str = CSmartStr(str) & " shape_type=""Cylinder"""
		else
			if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"cview"),3) then
				str = CSmartStr(str) & " shape_type=""Pyramid"""
			end if
		end if
	end if
	doAssignmentByRef method_Chart_Bar_chart_style_type,str
	Exit Function
End Function
Function method_Chart_Bar_write_Stack(byref this_object)
	if bValue(this_object.stacked) then
		if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sstacked"),"true") then
			ResponseWrite "<scale mode=""PercentStacked"" maximum=""100"" major_interval=""10""/>" & vblf
		else
			ResponseWrite "<scale mode=""Stacked""/>" & vblf
		end if
	end if
End Function
Function method_Chart_Bar_write_label_settings(byref this_object)
	Dim rotation,position,effect,str
	rotation = ""
	position = ""
	effect = ""
	if bValue(this_object.stacked) then
		rotation = " rotation=""0"""
		position = "<position  anchor=""Center"" halign=""Center"" valign=""Center"" padding=""0""/>"
		effect = "<font bold=""False"" color=""White"">" & vblf
		effect = CSmartStr(effect) & ("<effects>" & vblf)
		effect = CSmartStr(effect) & ("<drop_shadow enabled=""True"" opacity=""0.5"" distance=""2"" blur_x=""1"" blur_y=""1""/>" & vblf)
		effect = CSmartStr(effect) & ("</effects>" & vblf)
		effect = CSmartStr(effect) & ("</font>" & vblf)
		effect = CSmartStr(effect) & ("<background enabled=""False""/>" & vblf)
	end if
	str = (((("<label_settings enabled=""" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sval"))) & """") & CSmartStr(rotation)) & ">") & vblf
	str = CSmartStr(str) & (CSmartStr(position) & vblf)
	str = CSmartStr(str) & (CSmartStr(effect) & vblf)
	str = CSmartStr(str) & ("</label_settings>" & vblf)
	doAssignmentByRef method_Chart_Bar_write_label_settings,str
	Exit Function
End Function
Function method_Chart_Bar_get_point(byref this_object,ByVal series,ByVal row)
	Dim strLabelFormat,this,str
	doAssignmentByRef strLabelFormat,this_object.labelFormat_p2(this_object.strLabel,row)
	str = ((((("<point id=""" & CSmartStr(strLabelFormat)) & """ name=""") & CSmartStr(strLabelFormat)) & """ y=""") & CSmartStr(this_object.chart_xmlencode_p1(CSmartDbl(asp_str_replace(",",".",ArrayElement(row,ArrayElement(this_object.arrDataSeries,series))))+0))) & """>"
	str = CSmartStr(str) & ((("<tooltip enabled=""True""><format>" & CSmartStr(this_object.tooltipFormat_p2(series,row))) & "</format></tooltip>") & vblf)
	str = CSmartStr(str) & ("</point>" & vblf)
	doAssignmentByRef method_Chart_Bar_get_point,str
	Exit Function
End Function
Function method_Chart_Bar_write_axes(byref this_object)
	Dim this
	this_object.write_axes_custom 
End Function
Function method_Chart_Bar_write_legend_tag(byref this_object)
	Dim posit,padd,hgt,align
	posit = ""
	padd = ""
	hgt = ""
	align = ""
	if (bValue(this_object.var_2d) and not bValue(this_object.bar)) and not bValue(this_object.stacked) then
		posit = "Bottom"
		align = "align=""Spread"""
		padd = "padding=""15"""
		hgt = "height=""20%"""
	else
		posit = "Right"
	end if
	ResponseWrite (((((((("<legend enabled=""true"" position=""" & CSmartStr(posit)) & """ ignore_auto_item=""true"" ") & CSmartStr(align)) & " ") & CSmartStr(padd)) & " ") & CSmartStr(hgt)) & ">") & vblf
End Function

'------ Class Chart_Line extends Chart ------
Class Chart_Line
	Public type_line
	Public strSQL
	Public label2
	Public numRecordsToShow
	Public totalRecords
	Public header
	Public footer
	Public strLabel
	Public arrDataLabels
	Public arrDataSeries
	Public arrDataColor
	Public arrFormatCurrency
	Public arrFormatDecimal
	Public arrFormatCustomer
	Public arrFormatCustomerStr
	Public arrDataSize
	Public arrAxesColor
	Public arrGaugeColor
	Public arrOHLC_high
	Public arrOHLC_low
	Public arrOHLC_open
	Public arrOHLC_close
	Public arrOHLC_candle
	Public arrOHLC_color
	Public arrOHLC_color_up
	Public arrOHLC_color_down
	Public sleg
	Public scol
	Public chrt_array
	Public webchart
	Public cname
	Public gstrOrderBy
	Public table_type
	Public sessionPrefix
	Public Function init_Chart_Line_p2(ByRef ch_array,ByVal param)
		DoAssignmentByRef init_Chart_Line_p2,method_Chart_Line_Chart_Line(me,ch_array,param)
	End Function
	Public Function write_data()
		DoAssignmentByRef write_data,method_Chart_Line_write_data(me)
	End Function
	Public Function write_dps()
		DoAssignmentByRef write_dps,method_Chart_Line_write_dps(me)
	End Function
	Public Function get_point_p2(ByVal series,ByVal row)
		DoAssignmentByRef get_point_p2,method_Chart_Line_get_point(me,series,row)
	End Function
	Public Function color_series_p1(ByVal series)
		DoAssignmentByRef color_series_p1,method_Chart_Line_color_series(me,series)
	End Function
	Public Function write_format()
		DoAssignmentByRef write_format,method_Chart_Line_write_format(me)
	End Function
	Public Function write_axes()
		DoAssignmentByRef write_axes,method_Chart_Line_write_axes(me)
	End Function
	Public Function write_series_type()
		DoAssignmentByRef write_series_type,method_Chart_Line_write_series_type(me)
	End Function
	Public Function write_legend_tag()
		DoAssignmentByRef write_legend_tag,method_Chart_Line_write_legend_tag(me)
	End Function
	Public Function write_get_x_axis()
		DoAssignmentByRef write_get_x_axis,method_Chart_Line_write_get_x_axis(me)
	End Function
	Public Function write_Stack()
		DoAssignmentByRef write_Stack,method_Chart_Line_write_Stack(me)
	End Function
	Public Function setFooter_p1(ByVal name)
		DoAssignmentByRef setFooter_p1,method_Chart_setFooter(me,name)
	End Function
	Public Function getFooter()
		DoAssignmentByRef getFooter,method_Chart_getFooter(me)
	End Function
	Public Function setHeader_p1(ByVal name)
		DoAssignmentByRef setHeader_p1,method_Chart_setHeader(me,name)
	End Function
	Public Function getHeader()
		DoAssignmentByRef getHeader,method_Chart_getHeader(me)
	End Function
	Public Function setLabelField_p1(ByVal name)
		DoAssignmentByRef setLabelField_p1,method_Chart_setLabelField(me,name)
	End Function
	Public Function getLabelField()
		DoAssignmentByRef getLabelField,method_Chart_getLabelField(me)
	End Function
	Public Function setSeriaColor_p2(ByVal color,ByVal index)
		DoAssignmentByRef setSeriaColor_p2,method_Chart_setSeriaColor(me,color,index)
	End Function
	Public Function getSeriaColor_p1(ByVal index)
		DoAssignmentByRef getSeriaColor_p1,method_Chart_getSeriaColor(me,index)
	End Function
	Public Function setScrollingState_p1(ByVal scroll)
		DoAssignmentByRef setScrollingState_p1,method_Chart_setScrollingState(me,scroll)
	End Function
	Public Function getScrollingState()
		DoAssignmentByRef getScrollingState,method_Chart_getScrollingState(me)
	End Function
	Public Function setMaxBarScroll_p1(ByVal num)
		DoAssignmentByRef setMaxBarScroll_p1,method_Chart_setMaxBarScroll(me,num)
	End Function
	Public Function getMaxBarScroll()
		DoAssignmentByRef getMaxBarScroll,method_Chart_getMaxBarScroll(me)
	End Function
	Public Function write()
		DoAssignmentByRef write,method_Chart_write(me)
	End Function
	Public Function write_legend()
		DoAssignmentByRef write_legend,method_Chart_write_legend(me)
	End Function
	Public Function write_chart_settings()
		DoAssignmentByRef write_chart_settings,method_Chart_write_chart_settings(me)
	End Function
	Public Function formatCurrency_p2(ByVal val,ByVal series)
		DoAssignmentByRef formatCurrency_p2,method_Chart_formatCurrency(me,val,series)
	End Function
	Public Function write_axes_custom()
		DoAssignmentByRef write_axes_custom,method_Chart_write_axes_custom(me)
	End Function
	Public Function write_Logarithmic()
		DoAssignmentByRef write_Logarithmic,method_Chart_write_Logarithmic(me)
	End Function
	Public Function write_Grid()
		DoAssignmentByRef write_Grid,method_Chart_write_Grid(me)
	End Function
	Public Function write_extra()
		DoAssignmentByRef write_extra,method_Chart_write_extra(me)
	End Function
	Public Function write_chart_background()
		DoAssignmentByRef write_chart_background,method_Chart_write_chart_background(me)
	End Function
	Public Function labelFormat_p2(ByVal fieldName,ByVal data)
		DoAssignmentByRef labelFormat_p2,method_Chart_labelFormat(me,fieldName,data)
	End Function
	Public Function getDefaultValue_p2(ByVal series,ByVal row)
		DoAssignmentByRef getDefaultValue_p2,method_Chart_getDefaultValue(me,series,row)
	End Function
	Public Function valueFormat_p2(ByVal series,ByVal x_axis)
		DoAssignmentByRef valueFormat_p2,method_Chart_valueFormat(me,series,x_axis)
	End Function
	Public Function valueFormat_p1(ByVal series)
		DoAssignmentByRef valueFormat_p1,method_Chart_valueFormat(me,series,false)
	End Function
	Public Function tooltipFormat_p2(ByVal series,ByVal row)
		DoAssignmentByRef tooltipFormat_p2,method_Chart_tooltipFormat(me,series,row)
	End Function
	Public Function get_data_p1(ByVal refr)
		DoAssignmentByRef get_data_p1,method_Chart_get_data(me,refr)
	End Function
	Public Function chart_xmlencode_p1(ByVal str)
		DoAssignmentByRef chart_xmlencode_p1,method_Chart_chart_xmlencode(me,str)
	End Function
	Public Function write_plot_background()
		DoAssignmentByRef write_plot_background,method_Chart_write_plot_background(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"type_line", type_line
		setArrElement out,"strSQL", strSQL
		setArrElement out,"label2", label2
		setArrElement out,"numRecordsToShow", numRecordsToShow
		setArrElement out,"totalRecords", totalRecords
		setArrElement out,"header", header
		setArrElement out,"footer", footer
		setArrElement out,"strLabel", strLabel
		setArrElement out,"arrDataLabels", arrDataLabels
		setArrElement out,"arrDataSeries", arrDataSeries
		setArrElement out,"arrDataColor", arrDataColor
		setArrElement out,"arrFormatCurrency", arrFormatCurrency
		setArrElement out,"arrFormatDecimal", arrFormatDecimal
		setArrElement out,"arrFormatCustomer", arrFormatCustomer
		setArrElement out,"arrFormatCustomerStr", arrFormatCustomerStr
		setArrElement out,"arrDataSize", arrDataSize
		setArrElement out,"arrAxesColor", arrAxesColor
		setArrElement out,"arrGaugeColor", arrGaugeColor
		setArrElement out,"arrOHLC_high", arrOHLC_high
		setArrElement out,"arrOHLC_low", arrOHLC_low
		setArrElement out,"arrOHLC_open", arrOHLC_open
		setArrElement out,"arrOHLC_close", arrOHLC_close
		setArrElement out,"arrOHLC_candle", arrOHLC_candle
		setArrElement out,"arrOHLC_color", arrOHLC_color
		setArrElement out,"arrOHLC_color_up", arrOHLC_color_up
		setArrElement out,"arrOHLC_color_down", arrOHLC_color_down
		setArrElement out,"sleg", sleg
		setArrElement out,"scol", scol
		setArrElement out,"chrt_array", chrt_array
		setArrElement out,"webchart", webchart
		setArrElement out,"cname", cname
		setArrElement out,"gstrOrderBy", gstrOrderBy
		setArrElement out,"table_type", table_type
		setArrElement out,"sessionPrefix", sessionPrefix
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment type_line, dict("type_line")
		DoAssignment strSQL, dict("strSQL")
		DoAssignment label2, dict("label2")
		DoAssignment numRecordsToShow, dict("numRecordsToShow")
		DoAssignment totalRecords, dict("totalRecords")
		DoAssignment header, dict("header")
		DoAssignment footer, dict("footer")
		DoAssignment strLabel, dict("strLabel")
		DoAssignment arrDataLabels, dict("arrDataLabels")
		DoAssignment arrDataSeries, dict("arrDataSeries")
		DoAssignment arrDataColor, dict("arrDataColor")
		DoAssignment arrFormatCurrency, dict("arrFormatCurrency")
		DoAssignment arrFormatDecimal, dict("arrFormatDecimal")
		DoAssignment arrFormatCustomer, dict("arrFormatCustomer")
		DoAssignment arrFormatCustomerStr, dict("arrFormatCustomerStr")
		DoAssignment arrDataSize, dict("arrDataSize")
		DoAssignment arrAxesColor, dict("arrAxesColor")
		DoAssignment arrGaugeColor, dict("arrGaugeColor")
		DoAssignment arrOHLC_high, dict("arrOHLC_high")
		DoAssignment arrOHLC_low, dict("arrOHLC_low")
		DoAssignment arrOHLC_open, dict("arrOHLC_open")
		DoAssignment arrOHLC_close, dict("arrOHLC_close")
		DoAssignment arrOHLC_candle, dict("arrOHLC_candle")
		DoAssignment arrOHLC_color, dict("arrOHLC_color")
		DoAssignment arrOHLC_color_up, dict("arrOHLC_color_up")
		DoAssignment arrOHLC_color_down, dict("arrOHLC_color_down")
		DoAssignment sleg, dict("sleg")
		DoAssignment scol, dict("scol")
		DoAssignment chrt_array, dict("chrt_array")
		DoAssignment webchart, dict("webchart")
		DoAssignment cname, dict("cname")
		DoAssignment gstrOrderBy, dict("gstrOrderBy")
		DoAssignment table_type, dict("table_type")
		DoAssignment sessionPrefix, dict("sessionPrefix")
	End Sub
' end serialize
End Class
' Chart_Line implementation 
Function method_Chart_Line_Chart_Line(byref this_object,ByRef ch_array,ByVal param)
	doClassAssignment this_object,"arrDataLabels",CreateDictionary()
	doClassAssignment this_object,"arrDataSeries",CreateDictionary()
	doClassAssignment this_object,"arrDataColor",CreateDictionary()
	doClassAssignment this_object,"arrFormatCurrency",CreateDictionary()
	doClassAssignment this_object,"arrFormatDecimal",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomer",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomerStr",CreateDictionary()
	doClassAssignment this_object,"arrDataSize",CreateDictionary()
	doClassAssignment this_object,"arrAxesColor",CreateDictionary()
	doClassAssignment this_object,"arrGaugeColor",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_high",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_low",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_open",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_close",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_candle",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_up",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_down",CreateDictionary()
	doClassAssignment this_object,"chrt_array",CreateDictionary()
	this_object.sessionPrefix = ""
	doClassAssignment this_object,"type_line",ArrayElement(param,"type_line")
	method_Chart_Chart this_object,ch_array,param
End Function
Function method_Chart_Line_write_data(byref this_object)
	Dim this
	ResponseWrite "<chart plot_type=""CategorizedVertical"">" & vblf
	ResponseWrite "<data>" & vblf
	this_object.get_data_p1 false
	ResponseWrite "</data>" & vblf
End Function
Function method_Chart_Line_write_dps(byref this_object)
	Dim this
	ResponseWrite (("<data_plot_settings default_series_type=""" & CSmartStr(this_object.write_series_type())) & """>") & vblf
	ResponseWrite "<line_series point_padding=""0.2"" group_padding=""1"">" & vblf
	ResponseWrite (("<label_settings enabled=""" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sval"))) & """>") & vblf
	ResponseWrite "<background enabled=""false""/>" & vblf
	ResponseWrite "<font color=""Rgb(45,45,45)"" bold=""true"" size=""9"">" & vblf
	ResponseWrite "<effects enabled=""true"">" & vblf
	ResponseWrite "<glow enabled=""true"" color=""White"" opacity=""1"" blur_x=""1.5"" blur_y=""1.5"" strength=""3""/>" & vblf
	ResponseWrite "</effects>" & vblf
	ResponseWrite "</font>" & vblf
	ResponseWrite "</label_settings>" & vblf
	ResponseWrite "<marker_settings enabled=""true""/>" & vblf
	ResponseWrite "<line_style>" & vblf
	ResponseWrite "<line thickness=""3""/>" & vblf
	ResponseWrite "</line_style>" & vblf
	ResponseWrite "</line_series>" & vblf
	ResponseWrite "</data_plot_settings>" & vblf
End Function
Function method_Chart_Line_get_point(byref this_object,ByVal series,ByVal row)
	Dim strLabelFormat,this,str
	doAssignmentByRef strLabelFormat,this_object.labelFormat_p2(this_object.strLabel,row)
	str = ((((("<point id=""" & CSmartStr(strLabelFormat)) & """ name=""") & CSmartStr(strLabelFormat)) & """ y=""") & CSmartStr(this_object.chart_xmlencode_p1(CSmartDbl(asp_str_replace(",",".",ArrayElement(row,ArrayElement(this_object.arrDataSeries,series))))+0))) & """>"
	str = CSmartStr(str) & ((("<tooltip enabled=""True""><format>" & CSmartStr(this_object.tooltipFormat_p2(series,row))) & "</format></tooltip>") & vblf)
	str = CSmartStr(str) & ("</point>" & vblf)
	doAssignmentByRef method_Chart_Line_get_point,str
	Exit Function
End Function
Function method_Chart_Line_color_series(byref this_object,ByVal series)
	this_object.scol = ("color=""#" & CSmartStr(ArrayElement(this_object.arrDataColor,series))) & """"
	if IsLess(1,asp_count(this_object.arrDataSeries)) then
		this_object.sleg = "Series"
	else
		this_object.sleg = "Points"
	end if
End Function
Function method_Chart_Line_write_format(byref this_object)
	Dim this
	if IsEqual(this_object.sleg,"Points") then
		ResponseWrite (("<format>{%Icon} {%Name} (" & CSmartStr(this_object.valueFormat_p1(0))) & ")</format>") & vblf
	end if
End Function
Function method_Chart_Line_write_axes(byref this_object)
	Dim this
	this_object.write_axes_custom 
End Function
Function method_Chart_Line_write_series_type(byref this_object)
	do
		If IsEqual(this_object.type_line,"line") then
			method_Chart_Line_write_series_type = "Line"
			Exit Function
			exit do
		End If
		If IsEqual(this_object.type_line,"line") or IsEqual(this_object.type_line,"spline") then
			method_Chart_Line_write_series_type = "Spline"
			Exit Function
			exit do
		End If
		If IsEqual(this_object.type_line,"line") or IsEqual(this_object.type_line,"spline") or IsEqual(this_object.type_line,"step_line") then
			method_Chart_Line_write_series_type = "StepLineForward"
			Exit Function
			exit do
		End If
	Loop While false
End Function
Function method_Chart_Line_write_legend_tag(byref this_object)
	ResponseWrite "<legend enabled=""true"" position=""Bottom"" ignore_auto_item=""true"" align=""Spread"" padding=""15"" height=""20%"">" & vblf
End Function
Function method_Chart_Line_write_get_x_axis(byref this_object)
	ResponseWrite "<x_axis tickmarks_placement=""Center"">" & vblf
	ResponseWrite "<title enabled=""true"" align=""Center"">" & vblf
End Function
Function method_Chart_Line_write_Stack(byref this_object)
	Exit Function
End Function

'------ Class Chart_Area extends Chart ------
Class Chart_Area
	Public stacked
	Public strSQL
	Public label2
	Public numRecordsToShow
	Public totalRecords
	Public header
	Public footer
	Public strLabel
	Public arrDataLabels
	Public arrDataSeries
	Public arrDataColor
	Public arrFormatCurrency
	Public arrFormatDecimal
	Public arrFormatCustomer
	Public arrFormatCustomerStr
	Public arrDataSize
	Public arrAxesColor
	Public arrGaugeColor
	Public arrOHLC_high
	Public arrOHLC_low
	Public arrOHLC_open
	Public arrOHLC_close
	Public arrOHLC_candle
	Public arrOHLC_color
	Public arrOHLC_color_up
	Public arrOHLC_color_down
	Public sleg
	Public scol
	Public chrt_array
	Public webchart
	Public cname
	Public gstrOrderBy
	Public table_type
	Public sessionPrefix
	Public Function init_Chart_Area_p2(ByRef ch_array,ByVal param)
		DoAssignmentByRef init_Chart_Area_p2,method_Chart_Area_Chart_Area(me,ch_array,param)
	End Function
	Public Function write_data()
		DoAssignmentByRef write_data,method_Chart_Area_write_data(me)
	End Function
	Public Function write_dps()
		DoAssignmentByRef write_dps,method_Chart_Area_write_dps(me)
	End Function
	Public Function get_point_p2(ByVal series,ByVal row)
		DoAssignmentByRef get_point_p2,method_Chart_Area_get_point(me,series,row)
	End Function
	Public Function color_series_p1(ByVal series)
		DoAssignmentByRef color_series_p1,method_Chart_Area_color_series(me,series)
	End Function
	Public Function write_axes()
		DoAssignmentByRef write_axes,method_Chart_Area_write_axes(me)
	End Function
	Public Function write_label_settings()
		DoAssignmentByRef write_label_settings,method_Chart_Area_write_label_settings(me)
	End Function
	Public Function write_legend_tag()
		DoAssignmentByRef write_legend_tag,method_Chart_Area_write_legend_tag(me)
	End Function
	Public Function write_Stack()
		DoAssignmentByRef write_Stack,method_Chart_Area_write_Stack(me)
	End Function
	Public Function write_get_x_axis()
		DoAssignmentByRef write_get_x_axis,method_Chart_Area_write_get_x_axis(me)
	End Function
	Public Function setFooter_p1(ByVal name)
		DoAssignmentByRef setFooter_p1,method_Chart_setFooter(me,name)
	End Function
	Public Function getFooter()
		DoAssignmentByRef getFooter,method_Chart_getFooter(me)
	End Function
	Public Function setHeader_p1(ByVal name)
		DoAssignmentByRef setHeader_p1,method_Chart_setHeader(me,name)
	End Function
	Public Function getHeader()
		DoAssignmentByRef getHeader,method_Chart_getHeader(me)
	End Function
	Public Function setLabelField_p1(ByVal name)
		DoAssignmentByRef setLabelField_p1,method_Chart_setLabelField(me,name)
	End Function
	Public Function getLabelField()
		DoAssignmentByRef getLabelField,method_Chart_getLabelField(me)
	End Function
	Public Function setSeriaColor_p2(ByVal color,ByVal index)
		DoAssignmentByRef setSeriaColor_p2,method_Chart_setSeriaColor(me,color,index)
	End Function
	Public Function getSeriaColor_p1(ByVal index)
		DoAssignmentByRef getSeriaColor_p1,method_Chart_getSeriaColor(me,index)
	End Function
	Public Function setScrollingState_p1(ByVal scroll)
		DoAssignmentByRef setScrollingState_p1,method_Chart_setScrollingState(me,scroll)
	End Function
	Public Function getScrollingState()
		DoAssignmentByRef getScrollingState,method_Chart_getScrollingState(me)
	End Function
	Public Function setMaxBarScroll_p1(ByVal num)
		DoAssignmentByRef setMaxBarScroll_p1,method_Chart_setMaxBarScroll(me,num)
	End Function
	Public Function getMaxBarScroll()
		DoAssignmentByRef getMaxBarScroll,method_Chart_getMaxBarScroll(me)
	End Function
	Public Function write()
		DoAssignmentByRef write,method_Chart_write(me)
	End Function
	Public Function write_legend()
		DoAssignmentByRef write_legend,method_Chart_write_legend(me)
	End Function
	Public Function write_format()
		DoAssignmentByRef write_format,method_Chart_write_format(me)
	End Function
	Public Function write_chart_settings()
		DoAssignmentByRef write_chart_settings,method_Chart_write_chart_settings(me)
	End Function
	Public Function formatCurrency_p2(ByVal val,ByVal series)
		DoAssignmentByRef formatCurrency_p2,method_Chart_formatCurrency(me,val,series)
	End Function
	Public Function write_axes_custom()
		DoAssignmentByRef write_axes_custom,method_Chart_write_axes_custom(me)
	End Function
	Public Function write_Logarithmic()
		DoAssignmentByRef write_Logarithmic,method_Chart_write_Logarithmic(me)
	End Function
	Public Function write_Grid()
		DoAssignmentByRef write_Grid,method_Chart_write_Grid(me)
	End Function
	Public Function write_extra()
		DoAssignmentByRef write_extra,method_Chart_write_extra(me)
	End Function
	Public Function write_chart_background()
		DoAssignmentByRef write_chart_background,method_Chart_write_chart_background(me)
	End Function
	Public Function labelFormat_p2(ByVal fieldName,ByVal data)
		DoAssignmentByRef labelFormat_p2,method_Chart_labelFormat(me,fieldName,data)
	End Function
	Public Function getDefaultValue_p2(ByVal series,ByVal row)
		DoAssignmentByRef getDefaultValue_p2,method_Chart_getDefaultValue(me,series,row)
	End Function
	Public Function valueFormat_p2(ByVal series,ByVal x_axis)
		DoAssignmentByRef valueFormat_p2,method_Chart_valueFormat(me,series,x_axis)
	End Function
	Public Function valueFormat_p1(ByVal series)
		DoAssignmentByRef valueFormat_p1,method_Chart_valueFormat(me,series,false)
	End Function
	Public Function tooltipFormat_p2(ByVal series,ByVal row)
		DoAssignmentByRef tooltipFormat_p2,method_Chart_tooltipFormat(me,series,row)
	End Function
	Public Function get_data_p1(ByVal refr)
		DoAssignmentByRef get_data_p1,method_Chart_get_data(me,refr)
	End Function
	Public Function chart_xmlencode_p1(ByVal str)
		DoAssignmentByRef chart_xmlencode_p1,method_Chart_chart_xmlencode(me,str)
	End Function
	Public Function write_plot_background()
		DoAssignmentByRef write_plot_background,method_Chart_write_plot_background(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"stacked", stacked
		setArrElement out,"strSQL", strSQL
		setArrElement out,"label2", label2
		setArrElement out,"numRecordsToShow", numRecordsToShow
		setArrElement out,"totalRecords", totalRecords
		setArrElement out,"header", header
		setArrElement out,"footer", footer
		setArrElement out,"strLabel", strLabel
		setArrElement out,"arrDataLabels", arrDataLabels
		setArrElement out,"arrDataSeries", arrDataSeries
		setArrElement out,"arrDataColor", arrDataColor
		setArrElement out,"arrFormatCurrency", arrFormatCurrency
		setArrElement out,"arrFormatDecimal", arrFormatDecimal
		setArrElement out,"arrFormatCustomer", arrFormatCustomer
		setArrElement out,"arrFormatCustomerStr", arrFormatCustomerStr
		setArrElement out,"arrDataSize", arrDataSize
		setArrElement out,"arrAxesColor", arrAxesColor
		setArrElement out,"arrGaugeColor", arrGaugeColor
		setArrElement out,"arrOHLC_high", arrOHLC_high
		setArrElement out,"arrOHLC_low", arrOHLC_low
		setArrElement out,"arrOHLC_open", arrOHLC_open
		setArrElement out,"arrOHLC_close", arrOHLC_close
		setArrElement out,"arrOHLC_candle", arrOHLC_candle
		setArrElement out,"arrOHLC_color", arrOHLC_color
		setArrElement out,"arrOHLC_color_up", arrOHLC_color_up
		setArrElement out,"arrOHLC_color_down", arrOHLC_color_down
		setArrElement out,"sleg", sleg
		setArrElement out,"scol", scol
		setArrElement out,"chrt_array", chrt_array
		setArrElement out,"webchart", webchart
		setArrElement out,"cname", cname
		setArrElement out,"gstrOrderBy", gstrOrderBy
		setArrElement out,"table_type", table_type
		setArrElement out,"sessionPrefix", sessionPrefix
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment stacked, dict("stacked")
		DoAssignment strSQL, dict("strSQL")
		DoAssignment label2, dict("label2")
		DoAssignment numRecordsToShow, dict("numRecordsToShow")
		DoAssignment totalRecords, dict("totalRecords")
		DoAssignment header, dict("header")
		DoAssignment footer, dict("footer")
		DoAssignment strLabel, dict("strLabel")
		DoAssignment arrDataLabels, dict("arrDataLabels")
		DoAssignment arrDataSeries, dict("arrDataSeries")
		DoAssignment arrDataColor, dict("arrDataColor")
		DoAssignment arrFormatCurrency, dict("arrFormatCurrency")
		DoAssignment arrFormatDecimal, dict("arrFormatDecimal")
		DoAssignment arrFormatCustomer, dict("arrFormatCustomer")
		DoAssignment arrFormatCustomerStr, dict("arrFormatCustomerStr")
		DoAssignment arrDataSize, dict("arrDataSize")
		DoAssignment arrAxesColor, dict("arrAxesColor")
		DoAssignment arrGaugeColor, dict("arrGaugeColor")
		DoAssignment arrOHLC_high, dict("arrOHLC_high")
		DoAssignment arrOHLC_low, dict("arrOHLC_low")
		DoAssignment arrOHLC_open, dict("arrOHLC_open")
		DoAssignment arrOHLC_close, dict("arrOHLC_close")
		DoAssignment arrOHLC_candle, dict("arrOHLC_candle")
		DoAssignment arrOHLC_color, dict("arrOHLC_color")
		DoAssignment arrOHLC_color_up, dict("arrOHLC_color_up")
		DoAssignment arrOHLC_color_down, dict("arrOHLC_color_down")
		DoAssignment sleg, dict("sleg")
		DoAssignment scol, dict("scol")
		DoAssignment chrt_array, dict("chrt_array")
		DoAssignment webchart, dict("webchart")
		DoAssignment cname, dict("cname")
		DoAssignment gstrOrderBy, dict("gstrOrderBy")
		DoAssignment table_type, dict("table_type")
		DoAssignment sessionPrefix, dict("sessionPrefix")
	End Sub
' end serialize
End Class
' Chart_Area implementation 
Function method_Chart_Area_Chart_Area(byref this_object,ByRef ch_array,ByVal param)
	doClassAssignment this_object,"arrDataLabels",CreateDictionary()
	doClassAssignment this_object,"arrDataSeries",CreateDictionary()
	doClassAssignment this_object,"arrDataColor",CreateDictionary()
	doClassAssignment this_object,"arrFormatCurrency",CreateDictionary()
	doClassAssignment this_object,"arrFormatDecimal",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomer",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomerStr",CreateDictionary()
	doClassAssignment this_object,"arrDataSize",CreateDictionary()
	doClassAssignment this_object,"arrAxesColor",CreateDictionary()
	doClassAssignment this_object,"arrGaugeColor",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_high",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_low",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_open",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_close",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_candle",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_up",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_down",CreateDictionary()
	doClassAssignment this_object,"chrt_array",CreateDictionary()
	this_object.sessionPrefix = ""
	doClassAssignment this_object,"stacked",ArrayElement(param,"stacked")
	method_Chart_Chart this_object,ch_array,param
End Function
Function method_Chart_Area_write_data(byref this_object)
	Dim this
	ResponseWrite "<chart plot_type=""CategorizedVertical"">" & vblf
	ResponseWrite "<data>" & vblf
	this_object.get_data_p1 false
	ResponseWrite "</data>" & vblf
End Function
Function method_Chart_Area_write_dps(byref this_object)
	Dim this
	ResponseWrite "<data_plot_settings default_series_type=""Area"">" & vblf
	ResponseWrite "<area_series point_padding=""0.2"" group_padding=""1"">" & vblf
	this_object.write_label_settings 
	ResponseWrite "<area_style>" & vblf
	ResponseWrite "<line enabled=""true"" thickness=""2"" color=""%Color""/>" & vblf
	ResponseWrite "<fill color=""%Color"" opacity=""0.5""/>" & vblf
	ResponseWrite "<states>" & vblf
	ResponseWrite "<hover>" & vblf
	ResponseWrite "<line enabled=""true"" thickness=""2"" color=""LightColor(%Color)""/>" & vblf
	ResponseWrite "<fill color=""LightColor(%Color)"" opacity=""1.0""/>" & vblf
	ResponseWrite "</hover>" & vblf
	ResponseWrite "</states>" & vblf
	ResponseWrite "</area_style>" & vblf
	ResponseWrite "<marker_settings enabled=""True"">" & vblf
	ResponseWrite "<marker type=""Circle"" size=""6""/>" & vblf
	ResponseWrite "</marker_settings>" & vblf
	ResponseWrite "<tooltip_settings enabled=""True"">" & vblf
	ResponseWrite "<background>" & vblf
	ResponseWrite "<border color=""DarkColor(%Color)""/>" & vblf
	ResponseWrite "</background>" & vblf
	ResponseWrite "<font color=""DarkColor(%Color)""/>" & vblf
	ResponseWrite "</tooltip_settings>" & vblf
	ResponseWrite "</area_series>" & vblf
	ResponseWrite "</data_plot_settings>" & vblf
End Function
Function method_Chart_Area_get_point(byref this_object,ByVal series,ByVal row)
	Dim strLabelFormat,this,str
	doAssignmentByRef strLabelFormat,this_object.labelFormat_p2(this_object.strLabel,row)
	str = ((((("<point id=""" & CSmartStr(strLabelFormat)) & """ name=""") & CSmartStr(strLabelFormat)) & """ y=""") & CSmartStr(this_object.chart_xmlencode_p1(CSmartDbl(asp_str_replace(",",".",ArrayElement(row,ArrayElement(this_object.arrDataSeries,series))))+0))) & """>"
	str = CSmartStr(str) & ((("<tooltip enabled=""True""><format>" & CSmartStr(this_object.tooltipFormat_p2(series,row))) & "</format></tooltip>") & vblf)
	str = CSmartStr(str) & ("</point>" & vblf)
	doAssignmentByRef method_Chart_Area_get_point,str
	Exit Function
End Function
Function method_Chart_Area_color_series(byref this_object,ByVal series)
	this_object.scol = ("color=""#" & CSmartStr(ArrayElement(this_object.arrDataColor,series))) & """"
	this_object.sleg = "Series"
End Function
Function method_Chart_Area_write_axes(byref this_object)
	Dim this
	this_object.write_axes_custom 
End Function
Function method_Chart_Area_write_label_settings(byref this_object)
	ResponseWrite "<label_settings enabled=""true"">" & vblf
	ResponseWrite "<position anchor=""CenterBottom""/>" & vblf
	ResponseWrite "<background enabled=""true"">" & vblf
	ResponseWrite "<border enabled=""false""/>" & vblf
	ResponseWrite "<fill enabled=""true"" type=""Solid"" color=""DarkColor(%Color)"" opacity=""0.8""/>" & vblf
	ResponseWrite "<effects enabled=""false""/>" & vblf
	ResponseWrite "<inside_margin all=""0""/>" & vblf
	ResponseWrite "<corners type=""Rounded"" all=""3""/>" & vblf
	ResponseWrite "</background>" & vblf
	ResponseWrite "<font color=""White"" bold=""false""/>" & vblf
	ResponseWrite "</label_settings>" & vblf
End Function
Function method_Chart_Area_write_legend_tag(byref this_object)
	ResponseWrite "<legend enabled=""true"" position=""Bottom"" ignore_auto_item=""true"" align=""Spread"" padding=""15"" height=""20%"">" & vblf
End Function
Function method_Chart_Area_write_Stack(byref this_object)
	if bValue(this_object.stacked) then
		if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sstacked"),"true") then
			ResponseWrite "<scale mode=""PercentStacked"" maximum=""100"" major_interval=""10""/>" & vblf
		else
			ResponseWrite "<scale mode=""Stacked""/>" & vblf
		end if
	end if
End Function
Function method_Chart_Area_write_get_x_axis(byref this_object)
	ResponseWrite "<x_axis>" & vblf
	ResponseWrite (("<line thickness=""1"" color=""DarkColor(#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color131"))) & ")"" caps=""None""/>") & vblf
	ResponseWrite (("<major_tickmark thickness=""1"" color=""DarkColor(#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color131"))) & ")"" caps=""None"" opacity=""1""/>") & vblf
	ResponseWrite (("<minor_tickmark thickness=""1"" color=""DarkColor(#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color131"))) & ")"" caps=""None"" opacity=""1""/>") & vblf
	ResponseWrite "<title enabled=""true"" align=""Center"">" & vblf
End Function

'------ Class Chart_Pie extends Chart ------
Class Chart_Pie
	Public pie
	Public var_2d
	Public strSQL
	Public label2
	Public numRecordsToShow
	Public totalRecords
	Public header
	Public footer
	Public strLabel
	Public arrDataLabels
	Public arrDataSeries
	Public arrDataColor
	Public arrFormatCurrency
	Public arrFormatDecimal
	Public arrFormatCustomer
	Public arrFormatCustomerStr
	Public arrDataSize
	Public arrAxesColor
	Public arrGaugeColor
	Public arrOHLC_high
	Public arrOHLC_low
	Public arrOHLC_open
	Public arrOHLC_close
	Public arrOHLC_candle
	Public arrOHLC_color
	Public arrOHLC_color_up
	Public arrOHLC_color_down
	Public sleg
	Public scol
	Public chrt_array
	Public webchart
	Public cname
	Public gstrOrderBy
	Public table_type
	Public sessionPrefix
	Public Function init_Chart_Pie_p2(ByRef ch_array,ByVal param)
		DoAssignmentByRef init_Chart_Pie_p2,method_Chart_Pie_Chart_Pie(me,ch_array,param)
	End Function
	Public Function write_data()
		DoAssignmentByRef write_data,method_Chart_Pie_write_data(me)
	End Function
	Public Function write_dps()
		DoAssignmentByRef write_dps,method_Chart_Pie_write_dps(me)
	End Function
	Public Function get_point_p2(ByVal series,ByVal row)
		DoAssignmentByRef get_point_p2,method_Chart_Pie_get_point(me,series,row)
	End Function
	Public Function write_axes()
		DoAssignmentByRef write_axes,method_Chart_Pie_write_axes(me)
	End Function
	Public Function plot_type_name()
		DoAssignmentByRef plot_type_name,method_Chart_Pie_plot_type_name(me)
	End Function
	Public Function write_label_settings()
		DoAssignmentByRef write_label_settings,method_Chart_Pie_write_label_settings(me)
	End Function
	Public Function write_legend_tag()
		DoAssignmentByRef write_legend_tag,method_Chart_Pie_write_legend_tag(me)
	End Function
	Public Function color_series_p1(ByVal series)
		DoAssignmentByRef color_series_p1,method_Chart_Pie_color_series(me,series)
	End Function
	Public Function setFooter_p1(ByVal name)
		DoAssignmentByRef setFooter_p1,method_Chart_setFooter(me,name)
	End Function
	Public Function getFooter()
		DoAssignmentByRef getFooter,method_Chart_getFooter(me)
	End Function
	Public Function setHeader_p1(ByVal name)
		DoAssignmentByRef setHeader_p1,method_Chart_setHeader(me,name)
	End Function
	Public Function getHeader()
		DoAssignmentByRef getHeader,method_Chart_getHeader(me)
	End Function
	Public Function setLabelField_p1(ByVal name)
		DoAssignmentByRef setLabelField_p1,method_Chart_setLabelField(me,name)
	End Function
	Public Function getLabelField()
		DoAssignmentByRef getLabelField,method_Chart_getLabelField(me)
	End Function
	Public Function setSeriaColor_p2(ByVal color,ByVal index)
		DoAssignmentByRef setSeriaColor_p2,method_Chart_setSeriaColor(me,color,index)
	End Function
	Public Function getSeriaColor_p1(ByVal index)
		DoAssignmentByRef getSeriaColor_p1,method_Chart_getSeriaColor(me,index)
	End Function
	Public Function setScrollingState_p1(ByVal scroll)
		DoAssignmentByRef setScrollingState_p1,method_Chart_setScrollingState(me,scroll)
	End Function
	Public Function getScrollingState()
		DoAssignmentByRef getScrollingState,method_Chart_getScrollingState(me)
	End Function
	Public Function setMaxBarScroll_p1(ByVal num)
		DoAssignmentByRef setMaxBarScroll_p1,method_Chart_setMaxBarScroll(me,num)
	End Function
	Public Function getMaxBarScroll()
		DoAssignmentByRef getMaxBarScroll,method_Chart_getMaxBarScroll(me)
	End Function
	Public Function write()
		DoAssignmentByRef write,method_Chart_write(me)
	End Function
	Public Function write_legend()
		DoAssignmentByRef write_legend,method_Chart_write_legend(me)
	End Function
	Public Function write_format()
		DoAssignmentByRef write_format,method_Chart_write_format(me)
	End Function
	Public Function write_chart_settings()
		DoAssignmentByRef write_chart_settings,method_Chart_write_chart_settings(me)
	End Function
	Public Function formatCurrency_p2(ByVal val,ByVal series)
		DoAssignmentByRef formatCurrency_p2,method_Chart_formatCurrency(me,val,series)
	End Function
	Public Function write_axes_custom()
		DoAssignmentByRef write_axes_custom,method_Chart_write_axes_custom(me)
	End Function
	Public Function write_Logarithmic()
		DoAssignmentByRef write_Logarithmic,method_Chart_write_Logarithmic(me)
	End Function
	Public Function write_Grid()
		DoAssignmentByRef write_Grid,method_Chart_write_Grid(me)
	End Function
	Public Function write_extra()
		DoAssignmentByRef write_extra,method_Chart_write_extra(me)
	End Function
	Public Function write_chart_background()
		DoAssignmentByRef write_chart_background,method_Chart_write_chart_background(me)
	End Function
	Public Function labelFormat_p2(ByVal fieldName,ByVal data)
		DoAssignmentByRef labelFormat_p2,method_Chart_labelFormat(me,fieldName,data)
	End Function
	Public Function getDefaultValue_p2(ByVal series,ByVal row)
		DoAssignmentByRef getDefaultValue_p2,method_Chart_getDefaultValue(me,series,row)
	End Function
	Public Function valueFormat_p2(ByVal series,ByVal x_axis)
		DoAssignmentByRef valueFormat_p2,method_Chart_valueFormat(me,series,x_axis)
	End Function
	Public Function valueFormat_p1(ByVal series)
		DoAssignmentByRef valueFormat_p1,method_Chart_valueFormat(me,series,false)
	End Function
	Public Function tooltipFormat_p2(ByVal series,ByVal row)
		DoAssignmentByRef tooltipFormat_p2,method_Chart_tooltipFormat(me,series,row)
	End Function
	Public Function get_data_p1(ByVal refr)
		DoAssignmentByRef get_data_p1,method_Chart_get_data(me,refr)
	End Function
	Public Function chart_xmlencode_p1(ByVal str)
		DoAssignmentByRef chart_xmlencode_p1,method_Chart_chart_xmlencode(me,str)
	End Function
	Public Function write_plot_background()
		DoAssignmentByRef write_plot_background,method_Chart_write_plot_background(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"pie", pie
		setArrElement out,"var_2d", var_2d
		setArrElement out,"strSQL", strSQL
		setArrElement out,"label2", label2
		setArrElement out,"numRecordsToShow", numRecordsToShow
		setArrElement out,"totalRecords", totalRecords
		setArrElement out,"header", header
		setArrElement out,"footer", footer
		setArrElement out,"strLabel", strLabel
		setArrElement out,"arrDataLabels", arrDataLabels
		setArrElement out,"arrDataSeries", arrDataSeries
		setArrElement out,"arrDataColor", arrDataColor
		setArrElement out,"arrFormatCurrency", arrFormatCurrency
		setArrElement out,"arrFormatDecimal", arrFormatDecimal
		setArrElement out,"arrFormatCustomer", arrFormatCustomer
		setArrElement out,"arrFormatCustomerStr", arrFormatCustomerStr
		setArrElement out,"arrDataSize", arrDataSize
		setArrElement out,"arrAxesColor", arrAxesColor
		setArrElement out,"arrGaugeColor", arrGaugeColor
		setArrElement out,"arrOHLC_high", arrOHLC_high
		setArrElement out,"arrOHLC_low", arrOHLC_low
		setArrElement out,"arrOHLC_open", arrOHLC_open
		setArrElement out,"arrOHLC_close", arrOHLC_close
		setArrElement out,"arrOHLC_candle", arrOHLC_candle
		setArrElement out,"arrOHLC_color", arrOHLC_color
		setArrElement out,"arrOHLC_color_up", arrOHLC_color_up
		setArrElement out,"arrOHLC_color_down", arrOHLC_color_down
		setArrElement out,"sleg", sleg
		setArrElement out,"scol", scol
		setArrElement out,"chrt_array", chrt_array
		setArrElement out,"webchart", webchart
		setArrElement out,"cname", cname
		setArrElement out,"gstrOrderBy", gstrOrderBy
		setArrElement out,"table_type", table_type
		setArrElement out,"sessionPrefix", sessionPrefix
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment pie, dict("pie")
		DoAssignment var_2d, dict("var_2d")
		DoAssignment strSQL, dict("strSQL")
		DoAssignment label2, dict("label2")
		DoAssignment numRecordsToShow, dict("numRecordsToShow")
		DoAssignment totalRecords, dict("totalRecords")
		DoAssignment header, dict("header")
		DoAssignment footer, dict("footer")
		DoAssignment strLabel, dict("strLabel")
		DoAssignment arrDataLabels, dict("arrDataLabels")
		DoAssignment arrDataSeries, dict("arrDataSeries")
		DoAssignment arrDataColor, dict("arrDataColor")
		DoAssignment arrFormatCurrency, dict("arrFormatCurrency")
		DoAssignment arrFormatDecimal, dict("arrFormatDecimal")
		DoAssignment arrFormatCustomer, dict("arrFormatCustomer")
		DoAssignment arrFormatCustomerStr, dict("arrFormatCustomerStr")
		DoAssignment arrDataSize, dict("arrDataSize")
		DoAssignment arrAxesColor, dict("arrAxesColor")
		DoAssignment arrGaugeColor, dict("arrGaugeColor")
		DoAssignment arrOHLC_high, dict("arrOHLC_high")
		DoAssignment arrOHLC_low, dict("arrOHLC_low")
		DoAssignment arrOHLC_open, dict("arrOHLC_open")
		DoAssignment arrOHLC_close, dict("arrOHLC_close")
		DoAssignment arrOHLC_candle, dict("arrOHLC_candle")
		DoAssignment arrOHLC_color, dict("arrOHLC_color")
		DoAssignment arrOHLC_color_up, dict("arrOHLC_color_up")
		DoAssignment arrOHLC_color_down, dict("arrOHLC_color_down")
		DoAssignment sleg, dict("sleg")
		DoAssignment scol, dict("scol")
		DoAssignment chrt_array, dict("chrt_array")
		DoAssignment webchart, dict("webchart")
		DoAssignment cname, dict("cname")
		DoAssignment gstrOrderBy, dict("gstrOrderBy")
		DoAssignment table_type, dict("table_type")
		DoAssignment sessionPrefix, dict("sessionPrefix")
	End Sub
' end serialize
End Class
' Chart_Pie implementation 
Function method_Chart_Pie_Chart_Pie(byref this_object,ByRef ch_array,ByVal param)
	doClassAssignment this_object,"arrDataLabels",CreateDictionary()
	doClassAssignment this_object,"arrDataSeries",CreateDictionary()
	doClassAssignment this_object,"arrDataColor",CreateDictionary()
	doClassAssignment this_object,"arrFormatCurrency",CreateDictionary()
	doClassAssignment this_object,"arrFormatDecimal",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomer",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomerStr",CreateDictionary()
	doClassAssignment this_object,"arrDataSize",CreateDictionary()
	doClassAssignment this_object,"arrAxesColor",CreateDictionary()
	doClassAssignment this_object,"arrGaugeColor",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_high",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_low",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_open",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_close",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_candle",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_up",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_down",CreateDictionary()
	doClassAssignment this_object,"chrt_array",CreateDictionary()
	this_object.sessionPrefix = ""
	doClassAssignment this_object,"pie",ArrayElement(param,"pie")
	doClassAssignment this_object,"var_2d",ArrayElement(param,"2d")
	method_Chart_Chart this_object,ch_array,param
End Function
Function method_Chart_Pie_write_data(byref this_object)
	Dim this
	ResponseWrite (("<chart plot_type=""" & CSmartStr(this_object.plot_type_name())) & """>") & vblf
	this_object.write_dps 
	ResponseWrite "<data>" & vblf
	this_object.get_data_p1 false
	ResponseWrite "</data>" & vblf
End Function
Function method_Chart_Pie_write_dps(byref this_object)
	Dim this
	if bValue(this_object.var_2d) then
		ResponseWrite "<data_plot_settings enable_3d_mode=""false"">" & vblf
	else
		ResponseWrite "<data_plot_settings enable_3d_mode=""true"">" & vblf
	end if
	ResponseWrite "<pie_series>" & vblf
	this_object.write_label_settings 
	ResponseWrite "</pie_series>" & vblf
	ResponseWrite "</data_plot_settings>" & vblf
End Function
Function method_Chart_Pie_get_point(byref this_object,ByVal series,ByVal row)
	Dim strLabelFormat,this,str,showvalname,formatvalname
	doAssignmentByRef strLabelFormat,this_object.labelFormat_p2(this_object.strLabel,row)
	str = ((((("<point id=""" & CSmartStr(strLabelFormat)) & """ name=""") & CSmartStr(strLabelFormat)) & """ y=""") & CSmartStr(this_object.chart_xmlencode_p1(CSmartDbl(asp_str_replace(",",".",ArrayElement(row,ArrayElement(this_object.arrDataSeries,series))))+0))) & """>"
	str = CSmartStr(str) & ((("<tooltip enabled=""True""><format>" & CSmartStr(this_object.tooltipFormat_p2(series,row))) & "</format></tooltip>") & vblf)
	showvalname = "false"
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sval"),"true") or IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sname"),"true") then
		showvalname = "true"
	end if
	formatvalname = ""
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sval"),"true") then
		doAssignmentByRef formatvalname,this_object.chart_xmlencode_p1(ArrayElement(row,ArrayElement(this_object.arrDataSeries,series)))
	end if
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sval"),"true") and IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sname"),"true") then
		formatvalname = CSmartStr(formatvalname) & vblf
	end if
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sname"),"true") then
		formatvalname = CSmartStr(formatvalname) & CSmartStr(strLabelFormat)
	end if
	str = CSmartStr(str) & ((((("<label enabled=""" & CSmartStr(showvalname)) & """><format>") & CSmartStr(formatvalname)) & "</format></label>") & vblf)
	str = CSmartStr(str) & ("</point>" & vblf)
	doAssignmentByRef method_Chart_Pie_get_point,str
	Exit Function
End Function
Function method_Chart_Pie_write_axes(byref this_object)
	Exit Function
End Function
Function method_Chart_Pie_plot_type_name(byref this_object)
	if bValue(this_object.pie) then
		method_Chart_Pie_plot_type_name = "Pie"
		Exit Function
	else
		method_Chart_Pie_plot_type_name = "Doughnut"
		Exit Function
	end if
End Function
Function method_Chart_Pie_write_label_settings(byref this_object)
	Dim showvalname
	showvalname = "false"
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sval"),"true") or IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sname"),"true") then
		showvalname = "true"
	end if
	ResponseWrite (("<label_settings enabled=""" & CSmartStr(showvalname)) & """ mode=""Outside"" multi_line_align=""Center"">") & vblf
	ResponseWrite (("<font color=""#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color61"))) & """ bold=""false"" italic=""false"" underline=""false"" render_as_html=""false"">") & vblf
	ResponseWrite "</font>" & vblf
	ResponseWrite "<background enabled=""false""/>" & vblf
	ResponseWrite "<position anchor=""Center"" valign=""Center"" halign=""Center"" padding=""20""/>" & vblf
	ResponseWrite "<font bold=""false"" />" & vblf
	ResponseWrite "</label_settings>" & vblf
	ResponseWrite "<connector color=""Black"" opacity=""0.4""/>" & vblf
End Function
Function method_Chart_Pie_write_legend_tag(byref this_object)
	ResponseWrite "<legend enabled=""true"" position=""Bottom"" ignore_auto_item=""true"" align=""Spread"" padding=""15"" height=""20%"">" & vblf
End Function
Function method_Chart_Pie_color_series(byref this_object,ByVal series)
	this_object.scol = "palette=""Default"""
	this_object.sleg = "Points"
End Function

'------ Class Chart_Combined extends Chart ------
Class Chart_Combined
	Public strSQL
	Public label2
	Public numRecordsToShow
	Public totalRecords
	Public header
	Public footer
	Public strLabel
	Public arrDataLabels
	Public arrDataSeries
	Public arrDataColor
	Public arrFormatCurrency
	Public arrFormatDecimal
	Public arrFormatCustomer
	Public arrFormatCustomerStr
	Public arrDataSize
	Public arrAxesColor
	Public arrGaugeColor
	Public arrOHLC_high
	Public arrOHLC_low
	Public arrOHLC_open
	Public arrOHLC_close
	Public arrOHLC_candle
	Public arrOHLC_color
	Public arrOHLC_color_up
	Public arrOHLC_color_down
	Public sleg
	Public scol
	Public chrt_array
	Public webchart
	Public cname
	Public gstrOrderBy
	Public table_type
	Public sessionPrefix
	Public Function init_Chart_Combined_p2(ByRef ch_array,ByVal param)
		DoAssignmentByRef init_Chart_Combined_p2,method_Chart_Combined_Chart_Combined(me,ch_array,param)
	End Function
	Public Function color_series_p1(ByVal series)
		DoAssignmentByRef color_series_p1,method_Chart_Combined_color_series(me,series)
	End Function
	Public Function get_type_series_p1(ByVal num_series)
		DoAssignmentByRef get_type_series_p1,method_Chart_Combined_get_type_series(me,num_series)
	End Function
	Public Function get_data_p1(ByVal refr)
		DoAssignmentByRef get_data_p1,method_Chart_Combined_get_data(me,refr)
	End Function
	Public Function get_point_p2(ByVal series,ByVal row)
		DoAssignmentByRef get_point_p2,method_Chart_Combined_get_point(me,series,row)
	End Function
	Public Function write_data()
		DoAssignmentByRef write_data,method_Chart_Combined_write_data(me)
	End Function
	Public Function write_dps()
		DoAssignmentByRef write_dps,method_Chart_Combined_write_dps(me)
	End Function
	Public Function write_legend_tag()
		DoAssignmentByRef write_legend_tag,method_Chart_Combined_write_legend_tag(me)
	End Function
	Public Function write_get_x_axis()
		DoAssignmentByRef write_get_x_axis,method_Chart_Combined_write_get_x_axis(me)
	End Function
	Public Function write_axes()
		DoAssignmentByRef write_axes,method_Chart_Combined_write_axes(me)
	End Function
	Public Function write_Stack()
		DoAssignmentByRef write_Stack,method_Chart_Combined_write_Stack(me)
	End Function
	Public Function setFooter_p1(ByVal name)
		DoAssignmentByRef setFooter_p1,method_Chart_setFooter(me,name)
	End Function
	Public Function getFooter()
		DoAssignmentByRef getFooter,method_Chart_getFooter(me)
	End Function
	Public Function setHeader_p1(ByVal name)
		DoAssignmentByRef setHeader_p1,method_Chart_setHeader(me,name)
	End Function
	Public Function getHeader()
		DoAssignmentByRef getHeader,method_Chart_getHeader(me)
	End Function
	Public Function setLabelField_p1(ByVal name)
		DoAssignmentByRef setLabelField_p1,method_Chart_setLabelField(me,name)
	End Function
	Public Function getLabelField()
		DoAssignmentByRef getLabelField,method_Chart_getLabelField(me)
	End Function
	Public Function setSeriaColor_p2(ByVal color,ByVal index)
		DoAssignmentByRef setSeriaColor_p2,method_Chart_setSeriaColor(me,color,index)
	End Function
	Public Function getSeriaColor_p1(ByVal index)
		DoAssignmentByRef getSeriaColor_p1,method_Chart_getSeriaColor(me,index)
	End Function
	Public Function setScrollingState_p1(ByVal scroll)
		DoAssignmentByRef setScrollingState_p1,method_Chart_setScrollingState(me,scroll)
	End Function
	Public Function getScrollingState()
		DoAssignmentByRef getScrollingState,method_Chart_getScrollingState(me)
	End Function
	Public Function setMaxBarScroll_p1(ByVal num)
		DoAssignmentByRef setMaxBarScroll_p1,method_Chart_setMaxBarScroll(me,num)
	End Function
	Public Function getMaxBarScroll()
		DoAssignmentByRef getMaxBarScroll,method_Chart_getMaxBarScroll(me)
	End Function
	Public Function write()
		DoAssignmentByRef write,method_Chart_write(me)
	End Function
	Public Function write_legend()
		DoAssignmentByRef write_legend,method_Chart_write_legend(me)
	End Function
	Public Function write_format()
		DoAssignmentByRef write_format,method_Chart_write_format(me)
	End Function
	Public Function write_chart_settings()
		DoAssignmentByRef write_chart_settings,method_Chart_write_chart_settings(me)
	End Function
	Public Function formatCurrency_p2(ByVal val,ByVal series)
		DoAssignmentByRef formatCurrency_p2,method_Chart_formatCurrency(me,val,series)
	End Function
	Public Function write_axes_custom()
		DoAssignmentByRef write_axes_custom,method_Chart_write_axes_custom(me)
	End Function
	Public Function write_Logarithmic()
		DoAssignmentByRef write_Logarithmic,method_Chart_write_Logarithmic(me)
	End Function
	Public Function write_Grid()
		DoAssignmentByRef write_Grid,method_Chart_write_Grid(me)
	End Function
	Public Function write_extra()
		DoAssignmentByRef write_extra,method_Chart_write_extra(me)
	End Function
	Public Function write_chart_background()
		DoAssignmentByRef write_chart_background,method_Chart_write_chart_background(me)
	End Function
	Public Function labelFormat_p2(ByVal fieldName,ByVal data)
		DoAssignmentByRef labelFormat_p2,method_Chart_labelFormat(me,fieldName,data)
	End Function
	Public Function getDefaultValue_p2(ByVal series,ByVal row)
		DoAssignmentByRef getDefaultValue_p2,method_Chart_getDefaultValue(me,series,row)
	End Function
	Public Function valueFormat_p2(ByVal series,ByVal x_axis)
		DoAssignmentByRef valueFormat_p2,method_Chart_valueFormat(me,series,x_axis)
	End Function
	Public Function valueFormat_p1(ByVal series)
		DoAssignmentByRef valueFormat_p1,method_Chart_valueFormat(me,series,false)
	End Function
	Public Function tooltipFormat_p2(ByVal series,ByVal row)
		DoAssignmentByRef tooltipFormat_p2,method_Chart_tooltipFormat(me,series,row)
	End Function
	Public Function chart_xmlencode_p1(ByVal str)
		DoAssignmentByRef chart_xmlencode_p1,method_Chart_chart_xmlencode(me,str)
	End Function
	Public Function write_plot_background()
		DoAssignmentByRef write_plot_background,method_Chart_write_plot_background(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"strSQL", strSQL
		setArrElement out,"label2", label2
		setArrElement out,"numRecordsToShow", numRecordsToShow
		setArrElement out,"totalRecords", totalRecords
		setArrElement out,"header", header
		setArrElement out,"footer", footer
		setArrElement out,"strLabel", strLabel
		setArrElement out,"arrDataLabels", arrDataLabels
		setArrElement out,"arrDataSeries", arrDataSeries
		setArrElement out,"arrDataColor", arrDataColor
		setArrElement out,"arrFormatCurrency", arrFormatCurrency
		setArrElement out,"arrFormatDecimal", arrFormatDecimal
		setArrElement out,"arrFormatCustomer", arrFormatCustomer
		setArrElement out,"arrFormatCustomerStr", arrFormatCustomerStr
		setArrElement out,"arrDataSize", arrDataSize
		setArrElement out,"arrAxesColor", arrAxesColor
		setArrElement out,"arrGaugeColor", arrGaugeColor
		setArrElement out,"arrOHLC_high", arrOHLC_high
		setArrElement out,"arrOHLC_low", arrOHLC_low
		setArrElement out,"arrOHLC_open", arrOHLC_open
		setArrElement out,"arrOHLC_close", arrOHLC_close
		setArrElement out,"arrOHLC_candle", arrOHLC_candle
		setArrElement out,"arrOHLC_color", arrOHLC_color
		setArrElement out,"arrOHLC_color_up", arrOHLC_color_up
		setArrElement out,"arrOHLC_color_down", arrOHLC_color_down
		setArrElement out,"sleg", sleg
		setArrElement out,"scol", scol
		setArrElement out,"chrt_array", chrt_array
		setArrElement out,"webchart", webchart
		setArrElement out,"cname", cname
		setArrElement out,"gstrOrderBy", gstrOrderBy
		setArrElement out,"table_type", table_type
		setArrElement out,"sessionPrefix", sessionPrefix
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment strSQL, dict("strSQL")
		DoAssignment label2, dict("label2")
		DoAssignment numRecordsToShow, dict("numRecordsToShow")
		DoAssignment totalRecords, dict("totalRecords")
		DoAssignment header, dict("header")
		DoAssignment footer, dict("footer")
		DoAssignment strLabel, dict("strLabel")
		DoAssignment arrDataLabels, dict("arrDataLabels")
		DoAssignment arrDataSeries, dict("arrDataSeries")
		DoAssignment arrDataColor, dict("arrDataColor")
		DoAssignment arrFormatCurrency, dict("arrFormatCurrency")
		DoAssignment arrFormatDecimal, dict("arrFormatDecimal")
		DoAssignment arrFormatCustomer, dict("arrFormatCustomer")
		DoAssignment arrFormatCustomerStr, dict("arrFormatCustomerStr")
		DoAssignment arrDataSize, dict("arrDataSize")
		DoAssignment arrAxesColor, dict("arrAxesColor")
		DoAssignment arrGaugeColor, dict("arrGaugeColor")
		DoAssignment arrOHLC_high, dict("arrOHLC_high")
		DoAssignment arrOHLC_low, dict("arrOHLC_low")
		DoAssignment arrOHLC_open, dict("arrOHLC_open")
		DoAssignment arrOHLC_close, dict("arrOHLC_close")
		DoAssignment arrOHLC_candle, dict("arrOHLC_candle")
		DoAssignment arrOHLC_color, dict("arrOHLC_color")
		DoAssignment arrOHLC_color_up, dict("arrOHLC_color_up")
		DoAssignment arrOHLC_color_down, dict("arrOHLC_color_down")
		DoAssignment sleg, dict("sleg")
		DoAssignment scol, dict("scol")
		DoAssignment chrt_array, dict("chrt_array")
		DoAssignment webchart, dict("webchart")
		DoAssignment cname, dict("cname")
		DoAssignment gstrOrderBy, dict("gstrOrderBy")
		DoAssignment table_type, dict("table_type")
		DoAssignment sessionPrefix, dict("sessionPrefix")
	End Sub
' end serialize
End Class
' Chart_Combined implementation 
Function method_Chart_Combined_Chart_Combined(byref this_object,ByRef ch_array,ByVal param)
	doClassAssignment this_object,"arrDataLabels",CreateDictionary()
	doClassAssignment this_object,"arrDataSeries",CreateDictionary()
	doClassAssignment this_object,"arrDataColor",CreateDictionary()
	doClassAssignment this_object,"arrFormatCurrency",CreateDictionary()
	doClassAssignment this_object,"arrFormatDecimal",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomer",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomerStr",CreateDictionary()
	doClassAssignment this_object,"arrDataSize",CreateDictionary()
	doClassAssignment this_object,"arrAxesColor",CreateDictionary()
	doClassAssignment this_object,"arrGaugeColor",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_high",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_low",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_open",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_close",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_candle",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_up",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_down",CreateDictionary()
	doClassAssignment this_object,"chrt_array",CreateDictionary()
	this_object.sessionPrefix = ""
	method_Chart_Chart this_object,ch_array,param
End Function
Function method_Chart_Combined_color_series(byref this_object,ByVal series)
	if IsLess(1,asp_count(this_object.arrDataSeries)) then
		this_object.scol = ("color=""#" & CSmartStr(ArrayElement(this_object.arrDataColor,series))) & """"
		this_object.sleg = "Series"
	else
		this_object.scol = "palette=""Default"""
		this_object.sleg = "Points"
	end if
End Function
Function method_Chart_Combined_get_type_series(byref this_object,ByVal num_series)
	if IsEqual(num_series,0) then
		method_Chart_Combined_get_type_series = "Spline"
		Exit Function
	else
		if IsEqual(num_series,1) then
			method_Chart_Combined_get_type_series = "SplineArea"
			Exit Function
		else
			method_Chart_Combined_get_type_series = "Bar"
			Exit Function
		end if
	end if
End Function
Function method_Chart_Combined_get_data(byref this_object,ByVal refr)
	Dim arrSer,i,this,rs,j,recPerRow,row
	Set arrSer = (CreateDictionary())
	i = 0
	do while IsLess(i,asp_count(this_object.arrDataSeries))
		this_object.color_series_p1 i
		setArrElement arrSer,"series" & CSmartStr(i),(((((((((("<series id= """ & CSmartStr(this_object.chart_xmlencode_p1(ArrayElement(this_object.arrDataSeries,i)))) & """ name=""") & CSmartStr(ArrayElement(this_object.arrDataLabels,i))) & """ ") & CSmartStr(this_object.scol)) & " ") & CSmartStr(IIF(IsEqual(i,0),"",(" y_axis=""" & CSmartStr(this_object.chart_xmlencode_p1(ArrayElement(this_object.arrDataSeries,i)))) & """"))) & " type=""") & CSmartStr(this_object.get_type_series_p1(i))) & """>") & vblf
		setArrElement arrSer,"series" & CSmartStr(i),CSmartStr(ArrayElement(arrSer,"series" & CSmartStr(i))) & ((((("<label enabled=""" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sval"))) & """><format>") & CSmartStr(this_object.valueFormat_p1(i))) & "</format></label>") & vblf)
		i = CSmartDbl(i)+1
	loop
	doAssignmentByRef rs,db_query(this_object.strSQL,conn)
	j = 0
	doAssignment recPerRow,this_object.numRecordsToShow
	do while bValue(doAssignmentByRef(row,db_fetch_array(rs)))
		j = CSmartDbl(j)+1
		if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"cscroll"),"true") then
			recPerRow = CSmartDbl(recPerRow)+1
		end if
		if IsLess(recPerRow,j) and IsLess(0,recPerRow) then
			exit do
		end if
		i = 0
		do while IsLess(i,asp_count(this_object.arrDataSeries))
			setArrElement arrSer,"series" & CSmartStr(i),CSmartStr(ArrayElement(arrSer,"series" & CSmartStr(i))) & (CSmartStr(this_object.get_point_p2(i,row)) & vblf)
			i = CSmartDbl(i)+1
		loop
	loop
	doClassAssignment this_object,"totalRecords",j
	i = 0
	do while IsLess(i,asp_count(this_object.arrDataSeries))
		if bValue(refr) then
			ResponseWrite CSmartStr(ArrayElement(this_object.arrDataSeries,i)) & vblf
			setArrElement arrSer,"series" & CSmartStr(i),asp_str_replace(CreateDictionary2(Empty,"\",Empty,vblf),CreateDictionary2(Empty,"\\",Empty,"\n"),ArrayElement(arrSer,"series" & CSmartStr(i)))
		end if
		if IsLess(0,j) then
			ResponseWrite CSmartStr(ArrayElement(arrSer,"series" & CSmartStr(i))) & "</series>"
		end if
		if not bValue(refr) or IsLess(i,CSmartDbl(asp_count(this_object.arrDataSeries))-1) then
			ResponseWrite vblf
		end if
		i = CSmartDbl(i)+1
	loop
End Function
Function method_Chart_Combined_get_point(byref this_object,ByVal series,ByVal row)
	Dim strLabelFormat,this,str
	doAssignmentByRef strLabelFormat,this_object.labelFormat_p2(this_object.strLabel,row)
	str = ((((("<point id=""" & CSmartStr(strLabelFormat)) & """ name=""") & CSmartStr(strLabelFormat)) & """ y=""") & CSmartStr(this_object.chart_xmlencode_p1(CSmartDbl(ArrayElement(row,ArrayElement(this_object.arrDataSeries,series)))+0))) & """>"
	str = CSmartStr(str) & ((("<tooltip enabled=""True""><format>" & CSmartStr(this_object.tooltipFormat_p2(series,row))) & "</format></tooltip>") & vblf)
	str = CSmartStr(str) & ("</point>" & vblf)
	doAssignmentByRef method_Chart_Combined_get_point,str
	Exit Function
End Function
Function method_Chart_Combined_write_data(byref this_object)
	Dim this
	ResponseWrite "<chart plot_type=""CategorizedVertical"">" & vblf
	ResponseWrite "<data>" & vblf
	this_object.get_data_p1 false
	ResponseWrite "</data>" & vblf
End Function
Function method_Chart_Combined_write_dps(byref this_object)
	ResponseWrite "<data_plot_settings default_series_type=""Bar"">" & vblf
	ResponseWrite "<bar_series group_padding=""0.3"">" & vblf
	ResponseWrite "</bar_series>" & vblf
	ResponseWrite "<line_series>" & vblf
	ResponseWrite "<line_style>" & vblf
	ResponseWrite "<line thickness=""3""/>" & vblf
	ResponseWrite "</line_style>" & vblf
	ResponseWrite "</line_series>" & vblf
	ResponseWrite "<area_series>" & vblf
	ResponseWrite "<area_style>" & vblf
	ResponseWrite "<line enabled=""true"" thickness=""1"" color=""DarkColor(%Color)""/>" & vblf
	ResponseWrite "<fill opacity=""0.7""/>" & vblf
	ResponseWrite "<states>" & vblf
	ResponseWrite "<hover>" & vblf
	ResponseWrite "<fill opacity=""0.9""/>" & vblf
	ResponseWrite "<hatch_fill enabled=""true"" type=""Checkerboard"" opacity=""0.2""/>" & vblf
	ResponseWrite "</hover>" & vblf
	ResponseWrite "</states>" & vblf
	ResponseWrite "</area_style>" & vblf
	ResponseWrite "</area_series>" & vblf
	ResponseWrite "</data_plot_settings>" & vblf
End Function
Function method_Chart_Combined_write_legend_tag(byref this_object)
	ResponseWrite "<legend enabled=""true"" position=""Bottom"" ignore_auto_item=""true"" align=""Spread"" padding=""15"" height=""20%"">" & vblf
End Function
Function method_Chart_Combined_write_get_x_axis(byref this_object)
	ResponseWrite "<x_axis tickmarks_placement=""Center"">" & vblf
	ResponseWrite "<title enabled=""true"" align=""Center"">" & vblf
End Function
Function method_Chart_Combined_write_axes(byref this_object)
	Dim this
	this_object.write_axes_custom 
End Function
Function method_Chart_Combined_write_Stack(byref this_object)
	Exit Function
End Function

'------ Class Chart_Funnel extends Chart ------
Class Chart_Funnel
	Public ftype
	Public inver
	Public strSQL
	Public label2
	Public numRecordsToShow
	Public totalRecords
	Public header
	Public footer
	Public strLabel
	Public arrDataLabels
	Public arrDataSeries
	Public arrDataColor
	Public arrFormatCurrency
	Public arrFormatDecimal
	Public arrFormatCustomer
	Public arrFormatCustomerStr
	Public arrDataSize
	Public arrAxesColor
	Public arrGaugeColor
	Public arrOHLC_high
	Public arrOHLC_low
	Public arrOHLC_open
	Public arrOHLC_close
	Public arrOHLC_candle
	Public arrOHLC_color
	Public arrOHLC_color_up
	Public arrOHLC_color_down
	Public sleg
	Public scol
	Public chrt_array
	Public webchart
	Public cname
	Public gstrOrderBy
	Public table_type
	Public sessionPrefix
	Public Function init_Chart_Funnel_p2(ByRef ch_array,ByVal param)
		DoAssignmentByRef init_Chart_Funnel_p2,method_Chart_Funnel_Chart_Funnel(me,ch_array,param)
	End Function
	Public Function write_data()
		DoAssignmentByRef write_data,method_Chart_Funnel_write_data(me)
	End Function
	Public Function write_dps()
		DoAssignmentByRef write_dps,method_Chart_Funnel_write_dps(me)
	End Function
	Public Function series_3d_mode()
		DoAssignmentByRef series_3d_mode,method_Chart_Funnel_series_3d_mode(me)
	End Function
	Public Function get_point_p2(ByVal series,ByVal row)
		DoAssignmentByRef get_point_p2,method_Chart_Funnel_get_point(me,series,row)
	End Function
	Public Function write_axes()
		DoAssignmentByRef write_axes,method_Chart_Funnel_write_axes(me)
	End Function
	Public Function write_legend_tag()
		DoAssignmentByRef write_legend_tag,method_Chart_Funnel_write_legend_tag(me)
	End Function
	Public Function color_series_p1(ByVal series)
		DoAssignmentByRef color_series_p1,method_Chart_Funnel_color_series(me,series)
	End Function
	Public Function funnel_series()
		DoAssignmentByRef funnel_series,method_Chart_Funnel_funnel_series(me)
	End Function
	Public Function setFooter_p1(ByVal name)
		DoAssignmentByRef setFooter_p1,method_Chart_setFooter(me,name)
	End Function
	Public Function getFooter()
		DoAssignmentByRef getFooter,method_Chart_getFooter(me)
	End Function
	Public Function setHeader_p1(ByVal name)
		DoAssignmentByRef setHeader_p1,method_Chart_setHeader(me,name)
	End Function
	Public Function getHeader()
		DoAssignmentByRef getHeader,method_Chart_getHeader(me)
	End Function
	Public Function setLabelField_p1(ByVal name)
		DoAssignmentByRef setLabelField_p1,method_Chart_setLabelField(me,name)
	End Function
	Public Function getLabelField()
		DoAssignmentByRef getLabelField,method_Chart_getLabelField(me)
	End Function
	Public Function setSeriaColor_p2(ByVal color,ByVal index)
		DoAssignmentByRef setSeriaColor_p2,method_Chart_setSeriaColor(me,color,index)
	End Function
	Public Function getSeriaColor_p1(ByVal index)
		DoAssignmentByRef getSeriaColor_p1,method_Chart_getSeriaColor(me,index)
	End Function
	Public Function setScrollingState_p1(ByVal scroll)
		DoAssignmentByRef setScrollingState_p1,method_Chart_setScrollingState(me,scroll)
	End Function
	Public Function getScrollingState()
		DoAssignmentByRef getScrollingState,method_Chart_getScrollingState(me)
	End Function
	Public Function setMaxBarScroll_p1(ByVal num)
		DoAssignmentByRef setMaxBarScroll_p1,method_Chart_setMaxBarScroll(me,num)
	End Function
	Public Function getMaxBarScroll()
		DoAssignmentByRef getMaxBarScroll,method_Chart_getMaxBarScroll(me)
	End Function
	Public Function write()
		DoAssignmentByRef write,method_Chart_write(me)
	End Function
	Public Function write_legend()
		DoAssignmentByRef write_legend,method_Chart_write_legend(me)
	End Function
	Public Function write_format()
		DoAssignmentByRef write_format,method_Chart_write_format(me)
	End Function
	Public Function write_chart_settings()
		DoAssignmentByRef write_chart_settings,method_Chart_write_chart_settings(me)
	End Function
	Public Function formatCurrency_p2(ByVal val,ByVal series)
		DoAssignmentByRef formatCurrency_p2,method_Chart_formatCurrency(me,val,series)
	End Function
	Public Function write_axes_custom()
		DoAssignmentByRef write_axes_custom,method_Chart_write_axes_custom(me)
	End Function
	Public Function write_Logarithmic()
		DoAssignmentByRef write_Logarithmic,method_Chart_write_Logarithmic(me)
	End Function
	Public Function write_Grid()
		DoAssignmentByRef write_Grid,method_Chart_write_Grid(me)
	End Function
	Public Function write_extra()
		DoAssignmentByRef write_extra,method_Chart_write_extra(me)
	End Function
	Public Function write_chart_background()
		DoAssignmentByRef write_chart_background,method_Chart_write_chart_background(me)
	End Function
	Public Function labelFormat_p2(ByVal fieldName,ByVal data)
		DoAssignmentByRef labelFormat_p2,method_Chart_labelFormat(me,fieldName,data)
	End Function
	Public Function getDefaultValue_p2(ByVal series,ByVal row)
		DoAssignmentByRef getDefaultValue_p2,method_Chart_getDefaultValue(me,series,row)
	End Function
	Public Function valueFormat_p2(ByVal series,ByVal x_axis)
		DoAssignmentByRef valueFormat_p2,method_Chart_valueFormat(me,series,x_axis)
	End Function
	Public Function valueFormat_p1(ByVal series)
		DoAssignmentByRef valueFormat_p1,method_Chart_valueFormat(me,series,false)
	End Function
	Public Function tooltipFormat_p2(ByVal series,ByVal row)
		DoAssignmentByRef tooltipFormat_p2,method_Chart_tooltipFormat(me,series,row)
	End Function
	Public Function get_data_p1(ByVal refr)
		DoAssignmentByRef get_data_p1,method_Chart_get_data(me,refr)
	End Function
	Public Function chart_xmlencode_p1(ByVal str)
		DoAssignmentByRef chart_xmlencode_p1,method_Chart_chart_xmlencode(me,str)
	End Function
	Public Function write_plot_background()
		DoAssignmentByRef write_plot_background,method_Chart_write_plot_background(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"ftype", ftype
		setArrElement out,"inver", inver
		setArrElement out,"strSQL", strSQL
		setArrElement out,"label2", label2
		setArrElement out,"numRecordsToShow", numRecordsToShow
		setArrElement out,"totalRecords", totalRecords
		setArrElement out,"header", header
		setArrElement out,"footer", footer
		setArrElement out,"strLabel", strLabel
		setArrElement out,"arrDataLabels", arrDataLabels
		setArrElement out,"arrDataSeries", arrDataSeries
		setArrElement out,"arrDataColor", arrDataColor
		setArrElement out,"arrFormatCurrency", arrFormatCurrency
		setArrElement out,"arrFormatDecimal", arrFormatDecimal
		setArrElement out,"arrFormatCustomer", arrFormatCustomer
		setArrElement out,"arrFormatCustomerStr", arrFormatCustomerStr
		setArrElement out,"arrDataSize", arrDataSize
		setArrElement out,"arrAxesColor", arrAxesColor
		setArrElement out,"arrGaugeColor", arrGaugeColor
		setArrElement out,"arrOHLC_high", arrOHLC_high
		setArrElement out,"arrOHLC_low", arrOHLC_low
		setArrElement out,"arrOHLC_open", arrOHLC_open
		setArrElement out,"arrOHLC_close", arrOHLC_close
		setArrElement out,"arrOHLC_candle", arrOHLC_candle
		setArrElement out,"arrOHLC_color", arrOHLC_color
		setArrElement out,"arrOHLC_color_up", arrOHLC_color_up
		setArrElement out,"arrOHLC_color_down", arrOHLC_color_down
		setArrElement out,"sleg", sleg
		setArrElement out,"scol", scol
		setArrElement out,"chrt_array", chrt_array
		setArrElement out,"webchart", webchart
		setArrElement out,"cname", cname
		setArrElement out,"gstrOrderBy", gstrOrderBy
		setArrElement out,"table_type", table_type
		setArrElement out,"sessionPrefix", sessionPrefix
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment ftype, dict("ftype")
		DoAssignment inver, dict("inver")
		DoAssignment strSQL, dict("strSQL")
		DoAssignment label2, dict("label2")
		DoAssignment numRecordsToShow, dict("numRecordsToShow")
		DoAssignment totalRecords, dict("totalRecords")
		DoAssignment header, dict("header")
		DoAssignment footer, dict("footer")
		DoAssignment strLabel, dict("strLabel")
		DoAssignment arrDataLabels, dict("arrDataLabels")
		DoAssignment arrDataSeries, dict("arrDataSeries")
		DoAssignment arrDataColor, dict("arrDataColor")
		DoAssignment arrFormatCurrency, dict("arrFormatCurrency")
		DoAssignment arrFormatDecimal, dict("arrFormatDecimal")
		DoAssignment arrFormatCustomer, dict("arrFormatCustomer")
		DoAssignment arrFormatCustomerStr, dict("arrFormatCustomerStr")
		DoAssignment arrDataSize, dict("arrDataSize")
		DoAssignment arrAxesColor, dict("arrAxesColor")
		DoAssignment arrGaugeColor, dict("arrGaugeColor")
		DoAssignment arrOHLC_high, dict("arrOHLC_high")
		DoAssignment arrOHLC_low, dict("arrOHLC_low")
		DoAssignment arrOHLC_open, dict("arrOHLC_open")
		DoAssignment arrOHLC_close, dict("arrOHLC_close")
		DoAssignment arrOHLC_candle, dict("arrOHLC_candle")
		DoAssignment arrOHLC_color, dict("arrOHLC_color")
		DoAssignment arrOHLC_color_up, dict("arrOHLC_color_up")
		DoAssignment arrOHLC_color_down, dict("arrOHLC_color_down")
		DoAssignment sleg, dict("sleg")
		DoAssignment scol, dict("scol")
		DoAssignment chrt_array, dict("chrt_array")
		DoAssignment webchart, dict("webchart")
		DoAssignment cname, dict("cname")
		DoAssignment gstrOrderBy, dict("gstrOrderBy")
		DoAssignment table_type, dict("table_type")
		DoAssignment sessionPrefix, dict("sessionPrefix")
	End Sub
' end serialize
End Class
' Chart_Funnel implementation 
Function method_Chart_Funnel_Chart_Funnel(byref this_object,ByRef ch_array,ByVal param)
	doClassAssignment this_object,"arrDataLabels",CreateDictionary()
	doClassAssignment this_object,"arrDataSeries",CreateDictionary()
	doClassAssignment this_object,"arrDataColor",CreateDictionary()
	doClassAssignment this_object,"arrFormatCurrency",CreateDictionary()
	doClassAssignment this_object,"arrFormatDecimal",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomer",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomerStr",CreateDictionary()
	doClassAssignment this_object,"arrDataSize",CreateDictionary()
	doClassAssignment this_object,"arrAxesColor",CreateDictionary()
	doClassAssignment this_object,"arrGaugeColor",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_high",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_low",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_open",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_close",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_candle",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_up",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_down",CreateDictionary()
	doClassAssignment this_object,"chrt_array",CreateDictionary()
	this_object.sessionPrefix = ""
	doClassAssignment this_object,"ftype",ArrayElement(param,"funnel_type")
	doClassAssignment this_object,"inver",ArrayElement(param,"funnel_inv")
	method_Chart_Chart this_object,ch_array,param
End Function
Function method_Chart_Funnel_write_data(byref this_object)
	Dim this
	ResponseWrite "<chart plot_type=""Funnel"">" & vblf
	ResponseWrite "<data>" & vblf
	this_object.get_data_p1 false
	ResponseWrite "</data>" & vblf
End Function
Function method_Chart_Funnel_write_dps(byref this_object)
	Dim this
	ResponseWrite (("<data_plot_settings " & CSmartStr(this_object.series_3d_mode())) & ">") & vblf
	this_object.funnel_series 
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sval"),"true") or IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sname"),"true") then
		ResponseWrite "<connector enabled=""true"" color=""Black"" opacity=""0.5""/>" & vblf
	end if
	ResponseWrite "<label_settings enabled=""true"">" & vblf
	ResponseWrite "<animation enabled=""true"" type=""SideFromRight"" show_mode=""Smoothed"" start_time=""0.3"" duration=""2"" interpolation_type=""Back""/>" & vblf
	ResponseWrite "<position anchor=""center"" padding=""50""/>" & vblf
	ResponseWrite "<font bold=""true""/>" & vblf
	ResponseWrite "</label_settings>" & vblf
	ResponseWrite "<tooltip_settings enabled=""true"">" & vblf
	ResponseWrite "<background>" & vblf
	ResponseWrite "<corners type=""Rounded"" all=""3""/>" & vblf
	ResponseWrite "</background>" & vblf
	ResponseWrite "<font bold=""false""/>" & vblf
	ResponseWrite "</tooltip_settings>" & vblf
	ResponseWrite "<funnel_style>" & vblf
	ResponseWrite "<states>" & vblf
	ResponseWrite "<hover>" & vblf
	ResponseWrite "<fill color=""%Color""/>" & vblf
	ResponseWrite "<hatch_fill enabled=""true"" type=""Percent50"" color=""White"" opacity=""0.3""/>" & vblf
	ResponseWrite "</hover>" & vblf
	ResponseWrite "<selected_hover>" & vblf
	ResponseWrite "<fill color=""%Color""/>" & vblf
	ResponseWrite "<hatch_fill type=""Checkerboard"" color=""#404040"" opacity=""0.1""/>" & vblf
	ResponseWrite "</selected_hover>" & vblf
	ResponseWrite "<selected_normal>" & vblf
	ResponseWrite "<fill color=""%Color""/>" & vblf
	ResponseWrite "<hatch_fill type=""Checkerboard"" color=""Black"" opacity=""0.1""/>" & vblf
	ResponseWrite "</selected_normal>" & vblf
	ResponseWrite "</states>" & vblf
	ResponseWrite "</funnel_style>" & vblf
	ResponseWrite "</funnel_series>" & vblf
	ResponseWrite "</data_plot_settings>" & vblf
End Function
Function method_Chart_Funnel_series_3d_mode(byref this_object)
	Dim str
	str = ""
	if IsLess(0,this_object.ftype) then
		str = " enable_3d_mode=""True"""
	end if
	doAssignmentByRef method_Chart_Funnel_series_3d_mode,str
	Exit Function
End Function
Function method_Chart_Funnel_get_point(byref this_object,ByVal series,ByVal row)
	Dim strLabelFormat,this,str,showvalname,formatvalname
	doAssignmentByRef strLabelFormat,this_object.labelFormat_p2(this_object.strLabel,row)
	str = ((((("<point id=""" & CSmartStr(strLabelFormat)) & """ name=""") & CSmartStr(strLabelFormat)) & """ y=""") & CSmartStr(this_object.chart_xmlencode_p1(CSmartDbl(asp_str_replace(",",".",ArrayElement(row,ArrayElement(this_object.arrDataSeries,series))))+0))) & """>"
	str = CSmartStr(str) & ((("<tooltip enabled=""True""><format>" & CSmartStr(this_object.tooltipFormat_p2(series,row))) & "</format></tooltip>") & vblf)
	showvalname = "false"
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sval"),"true") or IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sname"),"true") then
		showvalname = "true"
	end if
	formatvalname = ""
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sval"),"true") then
		doAssignmentByRef formatvalname,this_object.chart_xmlencode_p1(ArrayElement(row,ArrayElement(this_object.arrDataSeries,series)))
	end if
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sval"),"true") and IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sname"),"true") then
		formatvalname = CSmartStr(formatvalname) & " - "
	end if
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sname"),"true") then
		formatvalname = CSmartStr(formatvalname) & CSmartStr(strLabelFormat)
	end if
	str = CSmartStr(str) & ((((("<label enabled=""" & CSmartStr(showvalname)) & """><format>") & CSmartStr(formatvalname)) & "</format></label>") & vblf)
	str = CSmartStr(str) & ("</point>" & vblf)
	doAssignmentByRef method_Chart_Funnel_get_point,str
	Exit Function
End Function
Function method_Chart_Funnel_write_axes(byref this_object)
	Exit Function
End Function
Function method_Chart_Funnel_write_legend_tag(byref this_object)
	ResponseWrite "<legend enabled=""true"" position=""Bottom"" ignore_auto_item=""true"" align=""Spread"" padding=""15"" height=""20%"">" & vblf
End Function
Function method_Chart_Funnel_color_series(byref this_object,ByVal series)
	this_object.scol = "palette=""Default"""
	this_object.sleg = "Points"
End Function
Function method_Chart_Funnel_funnel_series(byref this_object)
	Dim inv
	if bValue(this_object.inver) then
		inv = "inverted=""false"""
	else
		inv = "inverted=""true"""
	end if
	if IsLess(this_object.ftype,2) then
		ResponseWrite (("<funnel_series " & CSmartStr(inv)) & " neck_height=""0"" min_width=""0"" padding=""0"" fit_aspect=""0.9"">") & vblf
		ResponseWrite "<animation enabled=""true"" start_time=""0.3"" duration=""2"" type=""SideFromLeft"" animate_opacity=""false"" interpolation_type=""Elastic"" show_mode=""Smoothed""/>" & vblf
	else
		ResponseWrite (("<funnel_series " & CSmartStr(inv)) & " neck_height=""0"" fit_aspect=""1"" min_width=""0"" padding=""0"" mode=""Square"">") & vblf
		ResponseWrite "<animation enabled=""true"" start_time=""0.3"" duration=""2"" type=""SideFromTop"" animate_opacity=""true"" interpolation_type=""Bounce"" show_mode=""Smoothed"" />" & vblf
	end if
End Function

'------ Class Chart_Bubble extends Chart ------
Class Chart_Bubble
	Public var_2d
	Public oppos
	Public strSQL
	Public label2
	Public numRecordsToShow
	Public totalRecords
	Public header
	Public footer
	Public strLabel
	Public arrDataLabels
	Public arrDataSeries
	Public arrDataColor
	Public arrFormatCurrency
	Public arrFormatDecimal
	Public arrFormatCustomer
	Public arrFormatCustomerStr
	Public arrDataSize
	Public arrAxesColor
	Public arrGaugeColor
	Public arrOHLC_high
	Public arrOHLC_low
	Public arrOHLC_open
	Public arrOHLC_close
	Public arrOHLC_candle
	Public arrOHLC_color
	Public arrOHLC_color_up
	Public arrOHLC_color_down
	Public sleg
	Public scol
	Public chrt_array
	Public webchart
	Public cname
	Public gstrOrderBy
	Public table_type
	Public sessionPrefix
	Public Function init_Chart_Bubble_p2(ByRef ch_array,ByVal param)
		DoAssignmentByRef init_Chart_Bubble_p2,method_Chart_Bubble_Chart_Bubble(me,ch_array,param)
	End Function
	Public Function write_data()
		DoAssignmentByRef write_data,method_Chart_Bubble_write_data(me)
	End Function
	Public Function write_dps()
		DoAssignmentByRef write_dps,method_Chart_Bubble_write_dps(me)
	End Function
	Public Function write_legend_tag()
		DoAssignmentByRef write_legend_tag,method_Chart_Bubble_write_legend_tag(me)
	End Function
	Public Function char_type()
		DoAssignmentByRef char_type,method_Chart_Bubble_char_type(me)
	End Function
	Public Function fill_oppos()
		DoAssignmentByRef fill_oppos,method_Chart_Bubble_fill_oppos(me)
	End Function
	Public Function style_chart()
		DoAssignmentByRef style_chart,method_Chart_Bubble_style_chart(me)
	End Function
	Public Function get_point_p2(ByVal series,ByVal row)
		DoAssignmentByRef get_point_p2,method_Chart_Bubble_get_point(me,series,row)
	End Function
	Public Function write_axes()
		DoAssignmentByRef write_axes,method_Chart_Bubble_write_axes(me)
	End Function
	Public Function setFooter_p1(ByVal name)
		DoAssignmentByRef setFooter_p1,method_Chart_setFooter(me,name)
	End Function
	Public Function getFooter()
		DoAssignmentByRef getFooter,method_Chart_getFooter(me)
	End Function
	Public Function setHeader_p1(ByVal name)
		DoAssignmentByRef setHeader_p1,method_Chart_setHeader(me,name)
	End Function
	Public Function getHeader()
		DoAssignmentByRef getHeader,method_Chart_getHeader(me)
	End Function
	Public Function setLabelField_p1(ByVal name)
		DoAssignmentByRef setLabelField_p1,method_Chart_setLabelField(me,name)
	End Function
	Public Function getLabelField()
		DoAssignmentByRef getLabelField,method_Chart_getLabelField(me)
	End Function
	Public Function setSeriaColor_p2(ByVal color,ByVal index)
		DoAssignmentByRef setSeriaColor_p2,method_Chart_setSeriaColor(me,color,index)
	End Function
	Public Function getSeriaColor_p1(ByVal index)
		DoAssignmentByRef getSeriaColor_p1,method_Chart_getSeriaColor(me,index)
	End Function
	Public Function setScrollingState_p1(ByVal scroll)
		DoAssignmentByRef setScrollingState_p1,method_Chart_setScrollingState(me,scroll)
	End Function
	Public Function getScrollingState()
		DoAssignmentByRef getScrollingState,method_Chart_getScrollingState(me)
	End Function
	Public Function setMaxBarScroll_p1(ByVal num)
		DoAssignmentByRef setMaxBarScroll_p1,method_Chart_setMaxBarScroll(me,num)
	End Function
	Public Function getMaxBarScroll()
		DoAssignmentByRef getMaxBarScroll,method_Chart_getMaxBarScroll(me)
	End Function
	Public Function write()
		DoAssignmentByRef write,method_Chart_write(me)
	End Function
	Public Function write_legend()
		DoAssignmentByRef write_legend,method_Chart_write_legend(me)
	End Function
	Public Function write_format()
		DoAssignmentByRef write_format,method_Chart_write_format(me)
	End Function
	Public Function write_chart_settings()
		DoAssignmentByRef write_chart_settings,method_Chart_write_chart_settings(me)
	End Function
	Public Function formatCurrency_p2(ByVal val,ByVal series)
		DoAssignmentByRef formatCurrency_p2,method_Chart_formatCurrency(me,val,series)
	End Function
	Public Function write_axes_custom()
		DoAssignmentByRef write_axes_custom,method_Chart_write_axes_custom(me)
	End Function
	Public Function write_Logarithmic()
		DoAssignmentByRef write_Logarithmic,method_Chart_write_Logarithmic(me)
	End Function
	Public Function write_Grid()
		DoAssignmentByRef write_Grid,method_Chart_write_Grid(me)
	End Function
	Public Function write_extra()
		DoAssignmentByRef write_extra,method_Chart_write_extra(me)
	End Function
	Public Function write_chart_background()
		DoAssignmentByRef write_chart_background,method_Chart_write_chart_background(me)
	End Function
	Public Function color_series_p1(ByVal series)
		DoAssignmentByRef color_series_p1,method_Chart_color_series(me,series)
	End Function
	Public Function labelFormat_p2(ByVal fieldName,ByVal data)
		DoAssignmentByRef labelFormat_p2,method_Chart_labelFormat(me,fieldName,data)
	End Function
	Public Function getDefaultValue_p2(ByVal series,ByVal row)
		DoAssignmentByRef getDefaultValue_p2,method_Chart_getDefaultValue(me,series,row)
	End Function
	Public Function valueFormat_p2(ByVal series,ByVal x_axis)
		DoAssignmentByRef valueFormat_p2,method_Chart_valueFormat(me,series,x_axis)
	End Function
	Public Function valueFormat_p1(ByVal series)
		DoAssignmentByRef valueFormat_p1,method_Chart_valueFormat(me,series,false)
	End Function
	Public Function tooltipFormat_p2(ByVal series,ByVal row)
		DoAssignmentByRef tooltipFormat_p2,method_Chart_tooltipFormat(me,series,row)
	End Function
	Public Function get_data_p1(ByVal refr)
		DoAssignmentByRef get_data_p1,method_Chart_get_data(me,refr)
	End Function
	Public Function chart_xmlencode_p1(ByVal str)
		DoAssignmentByRef chart_xmlencode_p1,method_Chart_chart_xmlencode(me,str)
	End Function
	Public Function write_plot_background()
		DoAssignmentByRef write_plot_background,method_Chart_write_plot_background(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"var_2d", var_2d
		setArrElement out,"oppos", oppos
		setArrElement out,"strSQL", strSQL
		setArrElement out,"label2", label2
		setArrElement out,"numRecordsToShow", numRecordsToShow
		setArrElement out,"totalRecords", totalRecords
		setArrElement out,"header", header
		setArrElement out,"footer", footer
		setArrElement out,"strLabel", strLabel
		setArrElement out,"arrDataLabels", arrDataLabels
		setArrElement out,"arrDataSeries", arrDataSeries
		setArrElement out,"arrDataColor", arrDataColor
		setArrElement out,"arrFormatCurrency", arrFormatCurrency
		setArrElement out,"arrFormatDecimal", arrFormatDecimal
		setArrElement out,"arrFormatCustomer", arrFormatCustomer
		setArrElement out,"arrFormatCustomerStr", arrFormatCustomerStr
		setArrElement out,"arrDataSize", arrDataSize
		setArrElement out,"arrAxesColor", arrAxesColor
		setArrElement out,"arrGaugeColor", arrGaugeColor
		setArrElement out,"arrOHLC_high", arrOHLC_high
		setArrElement out,"arrOHLC_low", arrOHLC_low
		setArrElement out,"arrOHLC_open", arrOHLC_open
		setArrElement out,"arrOHLC_close", arrOHLC_close
		setArrElement out,"arrOHLC_candle", arrOHLC_candle
		setArrElement out,"arrOHLC_color", arrOHLC_color
		setArrElement out,"arrOHLC_color_up", arrOHLC_color_up
		setArrElement out,"arrOHLC_color_down", arrOHLC_color_down
		setArrElement out,"sleg", sleg
		setArrElement out,"scol", scol
		setArrElement out,"chrt_array", chrt_array
		setArrElement out,"webchart", webchart
		setArrElement out,"cname", cname
		setArrElement out,"gstrOrderBy", gstrOrderBy
		setArrElement out,"table_type", table_type
		setArrElement out,"sessionPrefix", sessionPrefix
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment var_2d, dict("var_2d")
		DoAssignment oppos, dict("oppos")
		DoAssignment strSQL, dict("strSQL")
		DoAssignment label2, dict("label2")
		DoAssignment numRecordsToShow, dict("numRecordsToShow")
		DoAssignment totalRecords, dict("totalRecords")
		DoAssignment header, dict("header")
		DoAssignment footer, dict("footer")
		DoAssignment strLabel, dict("strLabel")
		DoAssignment arrDataLabels, dict("arrDataLabels")
		DoAssignment arrDataSeries, dict("arrDataSeries")
		DoAssignment arrDataColor, dict("arrDataColor")
		DoAssignment arrFormatCurrency, dict("arrFormatCurrency")
		DoAssignment arrFormatDecimal, dict("arrFormatDecimal")
		DoAssignment arrFormatCustomer, dict("arrFormatCustomer")
		DoAssignment arrFormatCustomerStr, dict("arrFormatCustomerStr")
		DoAssignment arrDataSize, dict("arrDataSize")
		DoAssignment arrAxesColor, dict("arrAxesColor")
		DoAssignment arrGaugeColor, dict("arrGaugeColor")
		DoAssignment arrOHLC_high, dict("arrOHLC_high")
		DoAssignment arrOHLC_low, dict("arrOHLC_low")
		DoAssignment arrOHLC_open, dict("arrOHLC_open")
		DoAssignment arrOHLC_close, dict("arrOHLC_close")
		DoAssignment arrOHLC_candle, dict("arrOHLC_candle")
		DoAssignment arrOHLC_color, dict("arrOHLC_color")
		DoAssignment arrOHLC_color_up, dict("arrOHLC_color_up")
		DoAssignment arrOHLC_color_down, dict("arrOHLC_color_down")
		DoAssignment sleg, dict("sleg")
		DoAssignment scol, dict("scol")
		DoAssignment chrt_array, dict("chrt_array")
		DoAssignment webchart, dict("webchart")
		DoAssignment cname, dict("cname")
		DoAssignment gstrOrderBy, dict("gstrOrderBy")
		DoAssignment table_type, dict("table_type")
		DoAssignment sessionPrefix, dict("sessionPrefix")
	End Sub
' end serialize
End Class
' Chart_Bubble implementation 
Function method_Chart_Bubble_Chart_Bubble(byref this_object,ByRef ch_array,ByVal param)
	doClassAssignment this_object,"arrDataLabels",CreateDictionary()
	doClassAssignment this_object,"arrDataSeries",CreateDictionary()
	doClassAssignment this_object,"arrDataColor",CreateDictionary()
	doClassAssignment this_object,"arrFormatCurrency",CreateDictionary()
	doClassAssignment this_object,"arrFormatDecimal",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomer",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomerStr",CreateDictionary()
	doClassAssignment this_object,"arrDataSize",CreateDictionary()
	doClassAssignment this_object,"arrAxesColor",CreateDictionary()
	doClassAssignment this_object,"arrGaugeColor",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_high",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_low",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_open",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_close",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_candle",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_up",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_down",CreateDictionary()
	doClassAssignment this_object,"chrt_array",CreateDictionary()
	this_object.sessionPrefix = ""
	doClassAssignment this_object,"var_2d",ArrayElement(param,"2d")
	doClassAssignment this_object,"oppos",ArrayElement(param,"oppos")
	method_Chart_Chart this_object,ch_array,param
End Function
Function method_Chart_Bubble_write_data(byref this_object)
	Dim this
	ResponseWrite (("<chart " & CSmartStr(this_object.char_type())) & ">") & vblf
	ResponseWrite "<data>" & vblf
	this_object.get_data_p1 false
	ResponseWrite "</data>" & vblf
End Function
Function method_Chart_Bubble_write_dps(byref this_object)
	Dim this
	ResponseWrite "<data_plot_settings default_series_type=""Bubble"">" & vblf
	ResponseWrite (("<bubble_series maximum_bubble_size=""40%"" " & CSmartStr(this_object.style_chart())) & ">") & vblf
	ResponseWrite "<tooltip_settings enabled=""true"">" & vblf
	ResponseWrite "</tooltip_settings>" & vblf
	ResponseWrite "<bubble_style>" & vblf
	this_object.fill_oppos 
	ResponseWrite "<states>" & vblf
	ResponseWrite "<hover>" & vblf
	ResponseWrite "<border thickness=""2""/>" & vblf
	ResponseWrite "<fill color=""LightColor(%Color)""/>" & vblf
	ResponseWrite "</hover>" & vblf
	ResponseWrite "</states>" & vblf
	ResponseWrite "</bubble_style>" & vblf
	ResponseWrite "</bubble_series>" & vblf
	ResponseWrite "</data_plot_settings>" & vblf
End Function
Function method_Chart_Bubble_write_legend_tag(byref this_object)
	ResponseWrite "<legend enabled=""true"" position=""Bottom"" ignore_auto_item=""true"" align=""Spread"" padding=""15"" height=""20%"">" & vblf
End Function
Function method_Chart_Bubble_char_type(byref this_object)
	Dim str
	if IsEqual(this_object.strLabel,"") then
		str = "plot_type=""CategorizedBySeriesHorizontal"""
	else
		if bValue(this_object.var_2d) then
			str = "type=""CategorizedVertical"""
		else
			str = "type=""Categorized"""
		end if
	end if
	doAssignmentByRef method_Chart_Bubble_char_type,str
	Exit Function
End Function
Function method_Chart_Bubble_fill_oppos(byref this_object)
	ResponseWrite (("<fill opacity=""" & CSmartStr(this_object.oppos)) & """/>") & vblf
	ResponseWrite "<border thickness=""2""/>" & vblf
End Function
Function method_Chart_Bubble_style_chart(byref this_object)
	if not bValue(this_object.var_2d) then
		method_Chart_Bubble_style_chart = "style=""Aqua"""
		Exit Function
	end if
End Function
Function method_Chart_Bubble_get_point(byref this_object,ByVal series,ByVal row)
	Dim strLabelFormat,this,id_name,str
	doAssignmentByRef strLabelFormat,this_object.labelFormat_p2(this_object.strLabel,row)
	doAssignment id_name,IIF(not IsEqual(strLabelFormat,""),((("id=""" & CSmartStr(strLabelFormat)) & """ name=""") & CSmartStr(strLabelFormat)) & """","")
	str = ((((("<point " & CSmartStr(id_name)) & " y=""") & CSmartStr(this_object.chart_xmlencode_p1(CSmartDbl(ArrayElement(row,ArrayElement(this_object.arrDataSeries,series)))+0))) & """ size=""") & CSmartStr(this_object.chart_xmlencode_p1(CSmartDbl(asp_str_replace(",",".",ArrayElement(row,ArrayElement(this_object.arrDataSize,series))))+0))) & """>"
	str = CSmartStr(str) & ((("<tooltip enabled=""True""><format>" & CSmartStr(this_object.tooltipFormat_p2(series,row))) & "</format></tooltip>") & vblf)
	str = CSmartStr(str) & ("</point>" & vblf)
	doAssignmentByRef method_Chart_Bubble_get_point,str
	Exit Function
End Function
Function method_Chart_Bubble_write_axes(byref this_object)
	Dim this,scroll
	ResponseWrite "<axes>" & vblf
	ResponseWrite "<y_axis position=""Normal"">" & vblf
	ResponseWrite (("<line thickness=""1"" color=""DarkColor(#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color141"))) & ")"" caps=""None""/>") & vblf
	ResponseWrite (("<major_tickmark thickness=""1"" color=""DarkColor(#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color141"))) & ")"" caps=""None"" opacity=""1""/>") & vblf
	ResponseWrite (("<labels enabled=""" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sval"))) & """ align=""Inside"">") & vblf
	ResponseWrite (("<font color=""#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color61"))) & """ bold=""false"" italic=""false"" underline=""false"" render_as_html=""false"">") & vblf
	ResponseWrite "</font>" & vblf
	ResponseWrite "</labels>" & vblf
	ResponseWrite "<title enabled=""true"">" & vblf
	ResponseWrite (("<text>" & CSmartStr(ArrayElement(this_object.arrDataLabels,0))) & "</text>") & vblf
	ResponseWrite (("<font color=""DarkColor(#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color141"))) & ")""/>") & vblf
	ResponseWrite "</title>" & vblf
	ResponseWrite "<minor_grid enabled=""false""/>" & vblf
	ResponseWrite "<major_grid enabled=""true""/>" & vblf
	ResponseWrite "<minor_tickmark enabled=""false""/>" & vblf
	this_object.write_Grid 
	ResponseWrite "</y_axis>" & vblf
	ResponseWrite "<x_axis tickmarks_placement=""Center"">" & vblf
	scroll = "false"
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"cscroll"),"true") and IsLess(this_object.numRecordsToShow,this_object.totalRecords) then
		scroll = "true"
	end if
	ResponseWrite (((("<zoom enabled=""" & CSmartStr(scroll)) & """ allow_drag=""false"" visible_range=""") & CSmartStr(this_object.numRecordsToShow)) & """/>") & vblf
	ResponseWrite (("<line thickness=""1"" color=""DarkColor(#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color131"))) & ")"" caps=""None""/>") & vblf
	ResponseWrite "<title enabled=""true"" align=""Center"">" & vblf
	ResponseWrite (("<text>" & CSmartStr(this_object.label2)) & "</text>") & vblf
	ResponseWrite (("<font color=""DarkColor(#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color131"))) & ")""/>") & vblf
	ResponseWrite "</title>" & vblf
	ResponseWrite (("<labels enabled=""" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sname"))) & """ display_mode=""normal"">") & vblf
	ResponseWrite (("<font color=""#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color51"))) & """ bold=""false"" italic=""false"" underline=""false"" render_as_html=""false"">") & vblf
	ResponseWrite "</font>" & vblf
	ResponseWrite "<background enabled=""false"">" & vblf
	ResponseWrite "<fill enabled=""false"" />" & vblf
	ResponseWrite "<border enabled=""true"" />" & vblf
	ResponseWrite "</background>" & vblf
	ResponseWrite "</labels>" & vblf
	ResponseWrite "<scale inverted=""True""/>" & vblf
	ResponseWrite "</x_axis>" & vblf
	ResponseWrite "</axes>" & vblf
End Function

'------ Class Chart_Gauge extends Chart ------
Class Chart_Gauge
	Public type_gauge
	Public orientation
	Public start_angle
	Public sweep_angle
	Public scale_min
	Public scale_max
	Public major_interval
	Public minor_interval
	Public strSQL
	Public label2
	Public numRecordsToShow
	Public totalRecords
	Public header
	Public footer
	Public strLabel
	Public arrDataLabels
	Public arrDataSeries
	Public arrDataColor
	Public arrFormatCurrency
	Public arrFormatDecimal
	Public arrFormatCustomer
	Public arrFormatCustomerStr
	Public arrDataSize
	Public arrAxesColor
	Public arrGaugeColor
	Public arrOHLC_high
	Public arrOHLC_low
	Public arrOHLC_open
	Public arrOHLC_close
	Public arrOHLC_candle
	Public arrOHLC_color
	Public arrOHLC_color_up
	Public arrOHLC_color_down
	Public sleg
	Public scol
	Public chrt_array
	Public webchart
	Public cname
	Public gstrOrderBy
	Public table_type
	Public sessionPrefix
	Public Function init_Chart_Gauge_p2(ByRef ch_array,ByVal param)
		DoAssignmentByRef init_Chart_Gauge_p2,method_Chart_Gauge_Chart_Gauge(me,ch_array,param)
	End Function
	Public Function write()
		DoAssignmentByRef write,method_Chart_Gauge_write(me)
	End Function
	Public Function write_templates()
		DoAssignmentByRef write_templates,method_Chart_Gauge_write_templates(me)
	End Function
	Public Function gauge_style()
		DoAssignmentByRef gauge_style,method_Chart_Gauge_gauge_style(me)
	End Function
	Public Function write_data()
		DoAssignmentByRef write_data,method_Chart_Gauge_write_data(me)
	End Function
	Public Function write_chart_settings()
		DoAssignmentByRef write_chart_settings,method_Chart_Gauge_write_chart_settings(me)
	End Function
	Public Function get_data_p1(ByVal refr)
		DoAssignmentByRef get_data_p1,method_Chart_Gauge_get_data(me,refr)
	End Function
	Public Function pointer_type()
		DoAssignmentByRef pointer_type,method_Chart_Gauge_pointer_type(me)
	End Function
	Public Function pointer_label_p1(ByVal series)
		DoAssignmentByRef pointer_label_p1,method_Chart_Gauge_pointer_label(me,series)
	End Function
	Public Function pointer_style()
		DoAssignmentByRef pointer_style,method_Chart_Gauge_pointer_style(me)
	End Function
	Public Function get_frame()
		DoAssignmentByRef get_frame,method_Chart_Gauge_get_frame(me)
	End Function
	Public Function get_axis_p1(ByVal series)
		DoAssignmentByRef get_axis_p1,method_Chart_Gauge_get_axis(me,series)
	End Function
	Public Function get_tickmark()
		DoAssignmentByRef get_tickmark,method_Chart_Gauge_get_tickmark(me)
	End Function
	Public Function get_scale_bar()
		DoAssignmentByRef get_scale_bar,method_Chart_Gauge_get_scale_bar(me)
	End Function
	Public Function get_labels()
		DoAssignmentByRef get_labels,method_Chart_Gauge_get_labels(me)
	End Function
	Public Function setFooter_p1(ByVal name)
		DoAssignmentByRef setFooter_p1,method_Chart_setFooter(me,name)
	End Function
	Public Function getFooter()
		DoAssignmentByRef getFooter,method_Chart_getFooter(me)
	End Function
	Public Function setHeader_p1(ByVal name)
		DoAssignmentByRef setHeader_p1,method_Chart_setHeader(me,name)
	End Function
	Public Function getHeader()
		DoAssignmentByRef getHeader,method_Chart_getHeader(me)
	End Function
	Public Function setLabelField_p1(ByVal name)
		DoAssignmentByRef setLabelField_p1,method_Chart_setLabelField(me,name)
	End Function
	Public Function getLabelField()
		DoAssignmentByRef getLabelField,method_Chart_getLabelField(me)
	End Function
	Public Function setSeriaColor_p2(ByVal color,ByVal index)
		DoAssignmentByRef setSeriaColor_p2,method_Chart_setSeriaColor(me,color,index)
	End Function
	Public Function getSeriaColor_p1(ByVal index)
		DoAssignmentByRef getSeriaColor_p1,method_Chart_getSeriaColor(me,index)
	End Function
	Public Function setScrollingState_p1(ByVal scroll)
		DoAssignmentByRef setScrollingState_p1,method_Chart_setScrollingState(me,scroll)
	End Function
	Public Function getScrollingState()
		DoAssignmentByRef getScrollingState,method_Chart_getScrollingState(me)
	End Function
	Public Function setMaxBarScroll_p1(ByVal num)
		DoAssignmentByRef setMaxBarScroll_p1,method_Chart_setMaxBarScroll(me,num)
	End Function
	Public Function getMaxBarScroll()
		DoAssignmentByRef getMaxBarScroll,method_Chart_getMaxBarScroll(me)
	End Function
	Public Function write_legend()
		DoAssignmentByRef write_legend,method_Chart_write_legend(me)
	End Function
	Public Function write_format()
		DoAssignmentByRef write_format,method_Chart_write_format(me)
	End Function
	Public Function write_dps()
		DoAssignmentByRef write_dps,method_Chart_write_dps(me)
	End Function
	Public Function formatCurrency_p2(ByVal val,ByVal series)
		DoAssignmentByRef formatCurrency_p2,method_Chart_formatCurrency(me,val,series)
	End Function
	Public Function write_axes_custom()
		DoAssignmentByRef write_axes_custom,method_Chart_write_axes_custom(me)
	End Function
	Public Function write_Logarithmic()
		DoAssignmentByRef write_Logarithmic,method_Chart_write_Logarithmic(me)
	End Function
	Public Function write_Grid()
		DoAssignmentByRef write_Grid,method_Chart_write_Grid(me)
	End Function
	Public Function write_extra()
		DoAssignmentByRef write_extra,method_Chart_write_extra(me)
	End Function
	Public Function write_chart_background()
		DoAssignmentByRef write_chart_background,method_Chart_write_chart_background(me)
	End Function
	Public Function color_series_p1(ByVal series)
		DoAssignmentByRef color_series_p1,method_Chart_color_series(me,series)
	End Function
	Public Function labelFormat_p2(ByVal fieldName,ByVal data)
		DoAssignmentByRef labelFormat_p2,method_Chart_labelFormat(me,fieldName,data)
	End Function
	Public Function getDefaultValue_p2(ByVal series,ByVal row)
		DoAssignmentByRef getDefaultValue_p2,method_Chart_getDefaultValue(me,series,row)
	End Function
	Public Function valueFormat_p2(ByVal series,ByVal x_axis)
		DoAssignmentByRef valueFormat_p2,method_Chart_valueFormat(me,series,x_axis)
	End Function
	Public Function valueFormat_p1(ByVal series)
		DoAssignmentByRef valueFormat_p1,method_Chart_valueFormat(me,series,false)
	End Function
	Public Function tooltipFormat_p2(ByVal series,ByVal row)
		DoAssignmentByRef tooltipFormat_p2,method_Chart_tooltipFormat(me,series,row)
	End Function
	Public Function chart_xmlencode_p1(ByVal str)
		DoAssignmentByRef chart_xmlencode_p1,method_Chart_chart_xmlencode(me,str)
	End Function
	Public Function write_plot_background()
		DoAssignmentByRef write_plot_background,method_Chart_write_plot_background(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"type_gauge", type_gauge
		setArrElement out,"orientation", orientation
		setArrElement out,"start_angle", start_angle
		setArrElement out,"sweep_angle", sweep_angle
		setArrElement out,"scale_min", scale_min
		setArrElement out,"scale_max", scale_max
		setArrElement out,"major_interval", major_interval
		setArrElement out,"minor_interval", minor_interval
		setArrElement out,"strSQL", strSQL
		setArrElement out,"label2", label2
		setArrElement out,"numRecordsToShow", numRecordsToShow
		setArrElement out,"totalRecords", totalRecords
		setArrElement out,"header", header
		setArrElement out,"footer", footer
		setArrElement out,"strLabel", strLabel
		setArrElement out,"arrDataLabels", arrDataLabels
		setArrElement out,"arrDataSeries", arrDataSeries
		setArrElement out,"arrDataColor", arrDataColor
		setArrElement out,"arrFormatCurrency", arrFormatCurrency
		setArrElement out,"arrFormatDecimal", arrFormatDecimal
		setArrElement out,"arrFormatCustomer", arrFormatCustomer
		setArrElement out,"arrFormatCustomerStr", arrFormatCustomerStr
		setArrElement out,"arrDataSize", arrDataSize
		setArrElement out,"arrAxesColor", arrAxesColor
		setArrElement out,"arrGaugeColor", arrGaugeColor
		setArrElement out,"arrOHLC_high", arrOHLC_high
		setArrElement out,"arrOHLC_low", arrOHLC_low
		setArrElement out,"arrOHLC_open", arrOHLC_open
		setArrElement out,"arrOHLC_close", arrOHLC_close
		setArrElement out,"arrOHLC_candle", arrOHLC_candle
		setArrElement out,"arrOHLC_color", arrOHLC_color
		setArrElement out,"arrOHLC_color_up", arrOHLC_color_up
		setArrElement out,"arrOHLC_color_down", arrOHLC_color_down
		setArrElement out,"sleg", sleg
		setArrElement out,"scol", scol
		setArrElement out,"chrt_array", chrt_array
		setArrElement out,"webchart", webchart
		setArrElement out,"cname", cname
		setArrElement out,"gstrOrderBy", gstrOrderBy
		setArrElement out,"table_type", table_type
		setArrElement out,"sessionPrefix", sessionPrefix
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment type_gauge, dict("type_gauge")
		DoAssignment orientation, dict("orientation")
		DoAssignment start_angle, dict("start_angle")
		DoAssignment sweep_angle, dict("sweep_angle")
		DoAssignment scale_min, dict("scale_min")
		DoAssignment scale_max, dict("scale_max")
		DoAssignment major_interval, dict("major_interval")
		DoAssignment minor_interval, dict("minor_interval")
		DoAssignment strSQL, dict("strSQL")
		DoAssignment label2, dict("label2")
		DoAssignment numRecordsToShow, dict("numRecordsToShow")
		DoAssignment totalRecords, dict("totalRecords")
		DoAssignment header, dict("header")
		DoAssignment footer, dict("footer")
		DoAssignment strLabel, dict("strLabel")
		DoAssignment arrDataLabels, dict("arrDataLabels")
		DoAssignment arrDataSeries, dict("arrDataSeries")
		DoAssignment arrDataColor, dict("arrDataColor")
		DoAssignment arrFormatCurrency, dict("arrFormatCurrency")
		DoAssignment arrFormatDecimal, dict("arrFormatDecimal")
		DoAssignment arrFormatCustomer, dict("arrFormatCustomer")
		DoAssignment arrFormatCustomerStr, dict("arrFormatCustomerStr")
		DoAssignment arrDataSize, dict("arrDataSize")
		DoAssignment arrAxesColor, dict("arrAxesColor")
		DoAssignment arrGaugeColor, dict("arrGaugeColor")
		DoAssignment arrOHLC_high, dict("arrOHLC_high")
		DoAssignment arrOHLC_low, dict("arrOHLC_low")
		DoAssignment arrOHLC_open, dict("arrOHLC_open")
		DoAssignment arrOHLC_close, dict("arrOHLC_close")
		DoAssignment arrOHLC_candle, dict("arrOHLC_candle")
		DoAssignment arrOHLC_color, dict("arrOHLC_color")
		DoAssignment arrOHLC_color_up, dict("arrOHLC_color_up")
		DoAssignment arrOHLC_color_down, dict("arrOHLC_color_down")
		DoAssignment sleg, dict("sleg")
		DoAssignment scol, dict("scol")
		DoAssignment chrt_array, dict("chrt_array")
		DoAssignment webchart, dict("webchart")
		DoAssignment cname, dict("cname")
		DoAssignment gstrOrderBy, dict("gstrOrderBy")
		DoAssignment table_type, dict("table_type")
		DoAssignment sessionPrefix, dict("sessionPrefix")
	End Sub
' end serialize
End Class
' Chart_Gauge implementation 
Function method_Chart_Gauge_Chart_Gauge(byref this_object,ByRef ch_array,ByVal param)
	doClassAssignment this_object,"arrDataLabels",CreateDictionary()
	doClassAssignment this_object,"arrDataSeries",CreateDictionary()
	doClassAssignment this_object,"arrDataColor",CreateDictionary()
	doClassAssignment this_object,"arrFormatCurrency",CreateDictionary()
	doClassAssignment this_object,"arrFormatDecimal",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomer",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomerStr",CreateDictionary()
	doClassAssignment this_object,"arrDataSize",CreateDictionary()
	doClassAssignment this_object,"arrAxesColor",CreateDictionary()
	doClassAssignment this_object,"arrGaugeColor",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_high",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_low",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_open",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_close",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_candle",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_up",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_down",CreateDictionary()
	doClassAssignment this_object,"chrt_array",CreateDictionary()
	this_object.sessionPrefix = ""
	doClassAssignment this_object,"type_gauge",ArrayElement(param,"type_gauge")
	doClassAssignment this_object,"orientation",ArrayElement(param,"orientation")
	method_Chart_Chart this_object,ch_array,param
End Function
Function method_Chart_Gauge_write(byref this_object)
	Dim this
	ResponseWrite "<?xml version=""1.0"" standalone=""yes""?>" & vblf
	ResponseWrite "<anychart>" & vblf
	ResponseWrite "<settings>" & vblf
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sanim"),"true") then
		ResponseWrite "<animation enabled=""True"" />" & vblf
	else
		ResponseWrite "<animation enabled=""False"" />" & vblf
	end if
	ResponseWrite "</settings>" & vblf
	ResponseWrite "<templates>" & vblf
	ResponseWrite "<template name=""gaugeTemplates"">" & vblf
	ResponseWrite "<gauge>" & vblf
	this_object.write_templates 
	ResponseWrite "</gauge>" & vblf
	ResponseWrite "</template>" & vblf
	ResponseWrite "</templates>" & vblf
	ResponseWrite "<gauges>" & vblf
	ResponseWrite "<gauge template=""gaugeTemplates"">" & vblf
	this_object.write_data 
	ResponseWrite "</gauge>" & vblf
	ResponseWrite "</gauges>" & vblf
	ResponseWrite "</anychart>" & vblf
End Function
Function method_Chart_Gauge_write_templates(byref this_object)
	Dim strwidth,t,this
	strwidth = 100/CSmartDbl(asp_count(this_object.arrDataSeries))
	t = 0
	do while IsLess(t,asp_count(this_object.arrDataSeries))
		ResponseWrite ((((((((("<" & CSmartStr(this_object.type_gauge)) & "_template width=""") & CSmartStr(CSmartDbl(strwidth)-1)) & """ x=""") & CSmartStr(CSmartDbl(t)*CSmartDbl(strwidth)+1)) & """ name=""template_") & CSmartStr(this_object.chart_xmlencode_p1(ArrayElement(this_object.arrDataSeries,t)))) & CSmartStr(t)) & """>") & vblf
		this_object.gauge_style 
		this_object.get_frame 
		this_object.get_axis_p1 t
		ResponseWrite "<pointers>" & vblf
		this_object.pointer_label_p1 t
		ResponseWrite "</pointers>" & vblf
		ResponseWrite (("</" & CSmartStr(this_object.type_gauge)) & "_template>") & vblf
		t = CSmartDbl(t)+1
	loop
End Function
Function method_Chart_Gauge_gauge_style(byref this_object)
	if not IsEqual(this_object.type_gauge,"circular") then
		ResponseWrite "<styles>" & vblf
		ResponseWrite "<color_range_style name=""anychart_default"" align=""Outside"" padding=""3"" start_size=""15"" end_size=""15"">" & vblf
		ResponseWrite "<fill type=""Gradient"">" & vblf
		ResponseWrite "<gradient>" & vblf
		ResponseWrite "<key color=""Blend(%Color,DarkColor(%Color),0.5)""/>" & vblf
		ResponseWrite "<key color=""%Color""/>" & vblf
		ResponseWrite "<key color=""Blend(%Color,DarkColor(%Color),0.5)""/>" & vblf
		ResponseWrite "</gradient>" & vblf
		ResponseWrite "</fill>" & vblf
		ResponseWrite "<border enabled=""true"" color=""DarkColor(%Color)"" opacity=""0.8""/>" & vblf
		ResponseWrite "</color_range_style>" & vblf
		ResponseWrite "</styles>" & vblf
	end if
End Function
Function method_Chart_Gauge_write_data(byref this_object)
	Dim this
	this_object.write_chart_settings 
	this_object.get_data_p1 false
End Function
Function method_Chart_Gauge_write_chart_settings(byref this_object)
	Dim this
	ResponseWrite "<chart_settings>" & vblf
	ResponseWrite "<title enabled=""true"" padding=""15"">" & vblf
	ResponseWrite (("<text>" & CSmartStr(this_object.header)) & "</text>") & vblf
	ResponseWrite (("<font color=""#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color101"))) & """/>") & vblf
	ResponseWrite "</title>" & vblf
	ResponseWrite "<chart_background>" & vblf
	this_object.write_chart_background 
	ResponseWrite "</chart_background>" & vblf
	this_object.write_plot_background 
	ResponseWrite "</chart_settings>" & vblf
End Function
Function method_Chart_Gauge_get_data(byref this_object,ByVal refr)
	Dim i,p,ob,val,ind,rs,row,j,arrSer,this
	if IsEqual(this_object.table_type,"project") then
		i = 0
		doAssignmentByRef p,asp_strpos(asp_strtolower(this_object.strSQL),"order by",empty)
		if IsLess(0,p) then
			ob = "ORDER BY"
			GetCollectionBounds g_orderindexes,c_charts_loopIdx17,c_charts_loopMax17
			do while c_charts_loopIdx17<=c_charts_loopMax17
				ind = GetCollectionKey(g_orderindexes,c_charts_loopIdx17)
				doAssignment val,ArrayElement(g_orderindexes,ind)
				ob = CSmartStr(ob) & ((" " & CSmartStr(ArrayElement(val,0))) & " ")
				if IsEqual(ArrayElement(val,1),"ASC") then
					ob = CSmartStr(ob) & "DESC"
				else
					ob = CSmartStr(ob) & "ASC"
				end if
				if not IsEqual(CSmartDbl(ind)+1,asp_count(g_orderindexes)) then
					ob = CSmartStr(ob) & ","
				end if
				c_charts_loopIdx17=c_charts_loopIdx17+1
			loop
			this_object.strSQL = CSmartStr(asp_substr(this_object.strSQL,0,p)) & CSmartStr(ob)
		end if
	end if
	doAssignmentByRef rs,db_query(this_object.strSQL,conn)
	doAssignmentByRef row,db_fetch_array(rs)
	i = 0
	do while IsLess(i,asp_count(this_object.arrDataSeries))
		j = 0
		if bValue(row) then
			j = 1
			setArrElement arrSer,"series" & CSmartStr(i),(((((((((("<" & CSmartStr(this_object.type_gauge)) & " template=""template_") & CSmartStr(this_object.chart_xmlencode_p1(ArrayElement(this_object.arrDataSeries,i)))) & CSmartStr(i)) & """  orientation=""") & CSmartStr(this_object.orientation)) & """ name=""") & CSmartStr(this_object.chart_xmlencode_p1(ArrayElement(this_object.arrDataSeries,i)))) & CSmartStr(i)) & "_gauge"">") & vblf
			setArrElement arrSer,"series" & CSmartStr(i),CSmartStr(ArrayElement(arrSer,"series" & CSmartStr(i))) & ("<pointers>" & vblf)
			setArrElement arrSer,"series" & CSmartStr(i),CSmartStr(ArrayElement(arrSer,"series" & CSmartStr(i))) & ((((((("<pointer name=""" & CSmartStr(this_object.chart_xmlencode_p1(ArrayElement(this_object.arrDataSeries,i)))) & "_point"" type=""") & CSmartStr(this_object.pointer_type())) & """ value=""") & CSmartStr(this_object.chart_xmlencode_p1(CSmartDbl(ArrayElement(row,ArrayElement(this_object.arrDataSeries,i)))+0))) & """ color=""#75B7E1"">") & vblf)
			setArrElement arrSer,"series" & CSmartStr(i),CSmartStr(ArrayElement(arrSer,"series" & CSmartStr(i))) & ("<animation enabled=""true"" start_time=""0"" duration=""1"" interpolation_type=""Elastic""/>" & vblf)
			setArrElement arrSer,"series" & CSmartStr(i),CSmartStr(ArrayElement(arrSer,"series" & CSmartStr(i))) & CSmartStr(this_object.pointer_style())
			setArrElement arrSer,"series" & CSmartStr(i),CSmartStr(ArrayElement(arrSer,"series" & CSmartStr(i))) & ("</pointer>" & vblf)
			setArrElement arrSer,"series" & CSmartStr(i),CSmartStr(ArrayElement(arrSer,"series" & CSmartStr(i))) & ("</pointers>" & vblf)
			setArrElement arrSer,"series" & CSmartStr(i),CSmartStr(ArrayElement(arrSer,"series" & CSmartStr(i))) & ((("</" & CSmartStr(this_object.type_gauge)) & ">") & vblf)
		end if
		if bValue(refr) then
			ResponseWrite CSmartStr(ArrayElement(this_object.arrDataSeries,i)) & vblf
			ResponseWrite this_object.chart_xmlencode_p1(CSmartDbl(ArrayElement(row,ArrayElement(this_object.arrDataSeries,i)))+0)
		else
			if IsLess(0,j) then
				ResponseWrite ArrayElement(arrSer,"series" & CSmartStr(i))
			end if
		end if
		if not bValue(refr) or IsLess(i,CSmartDbl(asp_count(this_object.arrDataSeries))-1) then
			ResponseWrite vblf
		end if
		i = CSmartDbl(i)+1
	loop
	db_close conn
End Function
Function method_Chart_Gauge_pointer_type(byref this_object)
	if IsEqual(this_object.type_gauge,"circular") then
		method_Chart_Gauge_pointer_type = "Needle"
		Exit Function
	else
		method_Chart_Gauge_pointer_type = "Marker"
		Exit Function
	end if
End Function
Function method_Chart_Gauge_pointer_label(byref this_object,ByVal series)
	Dim y,this
	if IsEqual(this_object.type_gauge,"circular") then
		y = 90
	else
		if IsEqual(this_object.orientation,"vertical") then
			y = 99
		else
			y = 80
		end if
	end if
	ResponseWrite (("<label enabled=""" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sval"))) & """>") & vblf
	ResponseWrite (((("<format>" & CSmartStr(this_object.chart_xmlencode_p1(ArrayElement(this_object.arrDataSeries,series)))) & ": ") & CSmartStr(this_object.valueFormat_p2(series,true))) & "</format>") & vblf
	ResponseWrite (("<position placement_mode=""ByPoint"" x=""50"" y=""" & CSmartStr(y)) & """ valign=""Center"" halign=""Center""/>") & vblf
	ResponseWrite "<background>" & vblf
	ResponseWrite "<fill type=""Solid"" color=""White"" opacity=""0.8""/>" & vblf
	ResponseWrite "<border type=""Solid"" color=""Black"" opacity=""0.2""/>" & vblf
	ResponseWrite "<corners type=""Rounded"" all=""5""/>" & vblf
	ResponseWrite "<effects enabled=""false""/>" & vblf
	ResponseWrite "</background>" & vblf
	ResponseWrite "</label>" & vblf
End Function
Function method_Chart_Gauge_pointer_style(byref this_object)
	Dim res
	if IsEqual(this_object.type_gauge,"circular") then
		res = "<needle_pointer_style base_radius=""-50"">" & vblf
		res = CSmartStr(res) & ("<cap>" & vblf)
		res = CSmartStr(res) & ("<background>" & vblf)
		res = CSmartStr(res) & ("<fill type=""Gradient"">" & vblf)
		res = CSmartStr(res) & ("<gradient type=""Linear"" angle=""45"">" & vblf)
		res = CSmartStr(res) & ("<key color=""#D3D3D3""/>" & vblf)
		res = CSmartStr(res) & ("<key color=""#6F6F6F""/>" & vblf)
		res = CSmartStr(res) & ("</gradient>" & vblf)
		res = CSmartStr(res) & ("</fill>" & vblf)
		res = CSmartStr(res) & ("<border color=""Black"" opacity=""0.9""/>" & vblf)
		res = CSmartStr(res) & ("</background>" & vblf)
		res = CSmartStr(res) & ("<effects enabled=""true"">" & vblf)
		res = CSmartStr(res) & ("<bevel enabled=""true"" distance=""2"" shadow_opacity=""0.6"" highlight_opacity=""0.6""/>" & vblf)
		res = CSmartStr(res) & ("<drop_shadow enabled=""true"" distance=""1.5"" blur_x=""2"" blur_y=""2"" opacity=""0.4""/>" & vblf)
		res = CSmartStr(res) & ("</effects>" & vblf)
		res = CSmartStr(res) & ("</cap>" & vblf)
		res = CSmartStr(res) & ("</needle_pointer_style>" & vblf)
	else
		res = "<marker_pointer_style align=""Outside"" padding=""18.5""/>" & vblf
	end if
	doAssignmentByRef method_Chart_Gauge_pointer_style,res
	Exit Function
End Function
Function method_Chart_Gauge_get_frame(byref this_object)
	if IsEqual(this_object.type_gauge,"circular") then
		ResponseWrite "<frame type=""Rectangular"">" & vblf
		ResponseWrite "<inner_stroke enabled=""false""/>" & vblf
		ResponseWrite "<outer_stroke enabled=""false""/>" & vblf
		ResponseWrite "<corners type=""Rounded"" all=""15""/>" & vblf
		ResponseWrite "<background>" & vblf
		ResponseWrite (("<border enabled=""true"" color=""" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color81"))) & """ opacity=""0.5""/>") & vblf
		ResponseWrite "</background>" & vblf
		ResponseWrite "</frame>" & vblf
	end if
End Function
Function method_Chart_Gauge_get_axis(byref this_object,ByVal series)
	Dim diff,slog,muls,m,numDec,pos,this,val
	this_object.start_angle = 30
	this_object.sweep_angle = 300
	doClassAssignment this_object,"scale_min",ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),series),"gaugeMinValue")
	doClassAssignment this_object,"scale_max",ArrayElement(ArrayElement(ArrayElement(this_object.chrt_array,"parameters"),series),"gaugeMaxValue")
	if not bValue(asp_is_numeric(this_object.scale_min)) then
		this_object.scale_min = 0
	end if
	if not bValue(asp_is_numeric(this_object.scale_max)) then
		this_object.scale_max = 100
	end if
	diff = CSmartDbl(this_object.scale_max)-CSmartDbl(this_object.scale_min)
	doAssignmentByRef slog,asp_floor(log10(diff))
	doClassAssignment this_object,"major_interval",pow(10,CSmartDbl(slog)-2)
	Set muls = (CreateDictionary5(Empty,1,Empty,2,Empty,3,Empty,5,Empty,10))
	do while bValue(true)
		GetCollectionBounds muls,c_charts_loopIdx20,c_charts_loopMax20
		do while c_charts_loopIdx20<=c_charts_loopMax20
			c_charts_arrKey20 = GetCollectionKey(muls,c_charts_loopIdx20)
			doAssignment m,ArrayElement(muls,c_charts_arrKey20)
			if IsLessOrEqual(CSmartDbl(diff)/(CSmartDbl(this_object.major_interval)*CSmartDbl(m)),10) then
				this_object.major_interval = CSmartDbl(this_object.major_interval)*CSmartDbl(m)
				exit do
			end if
			c_charts_loopIdx20=c_charts_loopIdx20+1
		loop
		if IsLessOrEqual(CSmartDbl(diff)/CSmartDbl(this_object.major_interval),10) then
			exit do
		end if
		this_object.major_interval = CSmartDbl(this_object.major_interval)*10
	loop
	numDec = -CSmartDbl(asp_floor(log10(this_object.major_interval)))
	if IsLess(numDec,0) then
		numDec = 0
	end if
	pos = ""
	if IsEqual(this_object.type_gauge,"circular") then
		pos = "align=""Inside"" padding=""40"""
		ResponseWrite (((("<axis start_angle=""" & CSmartStr(this_object.start_angle)) & """ sweep_angle=""") & CSmartStr(this_object.sweep_angle)) & """>") & vblf
	else
		ResponseWrite "<axis>" & vblf
	end if
	ResponseWrite (((((("<scale minimum=""" & CSmartStr(this_object.scale_min)) & """ maximum=""") & CSmartStr(this_object.scale_max)) & """ major_interval=""") & CSmartStr(this_object.major_interval)) & """/>") & vblf
	ResponseWrite (("<labels enabled=""" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sval"))) & """>") & vblf
	ResponseWrite (("<format>{%Value}{numDecimals:" & CSmartStr(numDec)) & "}</format>") & vblf
	ResponseWrite "</labels>" & vblf
	this_object.get_tickmark 
	if IsLess(0,asp_count(this_object.arrGaugeColor)) and bValue(asp_array_key_exists(series,this_object.arrGaugeColor)) then
		ResponseWrite "<color_ranges>" & vblf
		GetCollectionBounds ArrayElement(this_object.arrGaugeColor,series),c_charts_loopIdx21,c_charts_loopMax21
		do while c_charts_loopIdx21<=c_charts_loopMax21
			ind = GetCollectionKey(ArrayElement(this_object.arrGaugeColor,series),c_charts_loopIdx21)
			doAssignment val,ArrayElement(ArrayElement(this_object.arrGaugeColor,series),ind)
			ResponseWrite (((((((("<color_range start=""" & CSmartStr(ArrayElement(val,0))) & """ end=""") & CSmartStr(ArrayElement(val,1))) & """ color=""") & CSmartStr(ArrayElement(val,2))) & """ ") & CSmartStr(pos)) & ">") & vblf
			if IsEqual(this_object.type_gauge,"circular") then
				ResponseWrite "<border enabled=""true"" color=""Black"" opacity=""0.2""/>" & vblf
				ResponseWrite "<fill opacity=""0.7""/>" & vblf
			end if
			ResponseWrite "</color_range>" & vblf
			c_charts_loopIdx21=c_charts_loopIdx21+1
		loop
		ResponseWrite "</color_ranges>" & vblf
	end if
	this_object.get_scale_bar 
	this_object.get_labels 
	ResponseWrite "</axis>" & vblf
End Function
Function method_Chart_Gauge_get_tickmark(byref this_object)
	if not IsEqual(this_object.type_gauge,"circular") then
		ResponseWrite "<major_tickmark shape=""Rectangle"" width=""1.3"" length=""10"" align=""Center"" padding=""0"">" & vblf
		ResponseWrite "<fill type=""Solid"" color=""White""/>" & vblf
		ResponseWrite "<border enabled=""true"" color=""#494949"" opacity=""0.5""/>" & vblf
		ResponseWrite "</major_tickmark>" & vblf
		ResponseWrite "<minor_tickmark shape=""Line"" align=""Center"" length=""7"">" & vblf
		ResponseWrite "<border enabled=""true"" color=""#494949"" opacity=""1""/>" & vblf
		ResponseWrite "</minor_tickmark>" & vblf
		ResponseWrite "<scale_bar enabled=""false""/>" & vblf
		ResponseWrite "<scale_line enabled=""false""/>" & vblf
	end if
End Function
Function method_Chart_Gauge_get_scale_bar(byref this_object)
	if IsEqual(this_object.type_gauge,"circular") then
		ResponseWrite "<scale_bar>" & vblf
		ResponseWrite "<fill color=""Rgb(200,200,200)""/>" & vblf
		ResponseWrite "</scale_bar>" & vblf
	end if
End Function
Function method_Chart_Gauge_get_labels(byref this_object)
	Dim this
	if not IsEqual(this_object.type_gauge,"circular") then
		ResponseWrite "<labels align=""Inside"" padding=""1"">" & vblf
		ResponseWrite (("<format>" & CSmartStr(this_object.valueFormat_p2(0,true))) & "</format>") & vblf
		ResponseWrite "</labels>" & vblf
	end if
End Function

'------ Class Chart_Ohlc extends Chart ------
Class Chart_Ohlc
	Public ohcl_type
	Public strSQL
	Public label2
	Public numRecordsToShow
	Public totalRecords
	Public header
	Public footer
	Public strLabel
	Public arrDataLabels
	Public arrDataSeries
	Public arrDataColor
	Public arrFormatCurrency
	Public arrFormatDecimal
	Public arrFormatCustomer
	Public arrFormatCustomerStr
	Public arrDataSize
	Public arrAxesColor
	Public arrGaugeColor
	Public arrOHLC_high
	Public arrOHLC_low
	Public arrOHLC_open
	Public arrOHLC_close
	Public arrOHLC_candle
	Public arrOHLC_color
	Public arrOHLC_color_up
	Public arrOHLC_color_down
	Public sleg
	Public scol
	Public chrt_array
	Public webchart
	Public cname
	Public gstrOrderBy
	Public table_type
	Public sessionPrefix
	Public Function write()
		DoAssignmentByRef write,method_Chart_Ohlc_write(me)
	End Function
	Public Function init_Chart_Ohlc_p2(ByRef ch_array,ByVal param)
		DoAssignmentByRef init_Chart_Ohlc_p2,method_Chart_Ohlc_Chart_Ohlc(me,ch_array,param)
	End Function
	Public Function write_data()
		DoAssignmentByRef write_data,method_Chart_Ohlc_write_data(me)
	End Function
	Public Function get_series_type()
		DoAssignmentByRef get_series_type,method_Chart_Ohlc_get_series_type(me)
	End Function
	Public Function write_dps()
		DoAssignmentByRef write_dps,method_Chart_Ohlc_write_dps(me)
	End Function
	Public Function ohls_styles()
		DoAssignmentByRef ohls_styles,method_Chart_Ohlc_ohls_styles(me)
	End Function
	Public Function get_ohcl_tooltip()
		DoAssignmentByRef get_ohcl_tooltip,method_Chart_Ohlc_get_ohcl_tooltip(me)
	End Function
	Public Function write_chart_settings()
		DoAssignmentByRef write_chart_settings,method_Chart_Ohlc_write_chart_settings(me)
	End Function
	Public Function get_data_p1(ByVal refr)
		DoAssignmentByRef get_data_p1,method_Chart_Ohlc_get_data(me,refr)
	End Function
	Public Function get_point_p2(ByVal row,ByVal series)
		DoAssignmentByRef get_point_p2,method_Chart_Ohlc_get_point(me,row,series)
	End Function
	Public Function write_Logarithmic()
		DoAssignmentByRef write_Logarithmic,method_Chart_Ohlc_write_Logarithmic(me)
	End Function
	Public Function write_legend_tag()
		DoAssignmentByRef write_legend_tag,method_Chart_Ohlc_write_legend_tag(me)
	End Function
	Public Function write_extra()
		DoAssignmentByRef write_extra,method_Chart_Ohlc_write_extra(me)
	End Function
	Public Function setFooter_p1(ByVal name)
		DoAssignmentByRef setFooter_p1,method_Chart_setFooter(me,name)
	End Function
	Public Function getFooter()
		DoAssignmentByRef getFooter,method_Chart_getFooter(me)
	End Function
	Public Function setHeader_p1(ByVal name)
		DoAssignmentByRef setHeader_p1,method_Chart_setHeader(me,name)
	End Function
	Public Function getHeader()
		DoAssignmentByRef getHeader,method_Chart_getHeader(me)
	End Function
	Public Function setLabelField_p1(ByVal name)
		DoAssignmentByRef setLabelField_p1,method_Chart_setLabelField(me,name)
	End Function
	Public Function getLabelField()
		DoAssignmentByRef getLabelField,method_Chart_getLabelField(me)
	End Function
	Public Function setSeriaColor_p2(ByVal color,ByVal index)
		DoAssignmentByRef setSeriaColor_p2,method_Chart_setSeriaColor(me,color,index)
	End Function
	Public Function getSeriaColor_p1(ByVal index)
		DoAssignmentByRef getSeriaColor_p1,method_Chart_getSeriaColor(me,index)
	End Function
	Public Function setScrollingState_p1(ByVal scroll)
		DoAssignmentByRef setScrollingState_p1,method_Chart_setScrollingState(me,scroll)
	End Function
	Public Function getScrollingState()
		DoAssignmentByRef getScrollingState,method_Chart_getScrollingState(me)
	End Function
	Public Function setMaxBarScroll_p1(ByVal num)
		DoAssignmentByRef setMaxBarScroll_p1,method_Chart_setMaxBarScroll(me,num)
	End Function
	Public Function getMaxBarScroll()
		DoAssignmentByRef getMaxBarScroll,method_Chart_getMaxBarScroll(me)
	End Function
	Public Function write_legend()
		DoAssignmentByRef write_legend,method_Chart_write_legend(me)
	End Function
	Public Function write_format()
		DoAssignmentByRef write_format,method_Chart_write_format(me)
	End Function
	Public Function formatCurrency_p2(ByVal val,ByVal series)
		DoAssignmentByRef formatCurrency_p2,method_Chart_formatCurrency(me,val,series)
	End Function
	Public Function write_axes_custom()
		DoAssignmentByRef write_axes_custom,method_Chart_write_axes_custom(me)
	End Function
	Public Function write_Grid()
		DoAssignmentByRef write_Grid,method_Chart_write_Grid(me)
	End Function
	Public Function write_chart_background()
		DoAssignmentByRef write_chart_background,method_Chart_write_chart_background(me)
	End Function
	Public Function color_series_p1(ByVal series)
		DoAssignmentByRef color_series_p1,method_Chart_color_series(me,series)
	End Function
	Public Function labelFormat_p2(ByVal fieldName,ByVal data)
		DoAssignmentByRef labelFormat_p2,method_Chart_labelFormat(me,fieldName,data)
	End Function
	Public Function getDefaultValue_p2(ByVal series,ByVal row)
		DoAssignmentByRef getDefaultValue_p2,method_Chart_getDefaultValue(me,series,row)
	End Function
	Public Function valueFormat_p2(ByVal series,ByVal x_axis)
		DoAssignmentByRef valueFormat_p2,method_Chart_valueFormat(me,series,x_axis)
	End Function
	Public Function valueFormat_p1(ByVal series)
		DoAssignmentByRef valueFormat_p1,method_Chart_valueFormat(me,series,false)
	End Function
	Public Function tooltipFormat_p2(ByVal series,ByVal row)
		DoAssignmentByRef tooltipFormat_p2,method_Chart_tooltipFormat(me,series,row)
	End Function
	Public Function chart_xmlencode_p1(ByVal str)
		DoAssignmentByRef chart_xmlencode_p1,method_Chart_chart_xmlencode(me,str)
	End Function
	Public Function write_plot_background()
		DoAssignmentByRef write_plot_background,method_Chart_write_plot_background(me)
	End Function
' serialize stuff
	Public Function ASPserialize
		dim out
		set out=CreateDictionary()
		setArrElement out,"ohcl_type", ohcl_type
		setArrElement out,"strSQL", strSQL
		setArrElement out,"label2", label2
		setArrElement out,"numRecordsToShow", numRecordsToShow
		setArrElement out,"totalRecords", totalRecords
		setArrElement out,"header", header
		setArrElement out,"footer", footer
		setArrElement out,"strLabel", strLabel
		setArrElement out,"arrDataLabels", arrDataLabels
		setArrElement out,"arrDataSeries", arrDataSeries
		setArrElement out,"arrDataColor", arrDataColor
		setArrElement out,"arrFormatCurrency", arrFormatCurrency
		setArrElement out,"arrFormatDecimal", arrFormatDecimal
		setArrElement out,"arrFormatCustomer", arrFormatCustomer
		setArrElement out,"arrFormatCustomerStr", arrFormatCustomerStr
		setArrElement out,"arrDataSize", arrDataSize
		setArrElement out,"arrAxesColor", arrAxesColor
		setArrElement out,"arrGaugeColor", arrGaugeColor
		setArrElement out,"arrOHLC_high", arrOHLC_high
		setArrElement out,"arrOHLC_low", arrOHLC_low
		setArrElement out,"arrOHLC_open", arrOHLC_open
		setArrElement out,"arrOHLC_close", arrOHLC_close
		setArrElement out,"arrOHLC_candle", arrOHLC_candle
		setArrElement out,"arrOHLC_color", arrOHLC_color
		setArrElement out,"arrOHLC_color_up", arrOHLC_color_up
		setArrElement out,"arrOHLC_color_down", arrOHLC_color_down
		setArrElement out,"sleg", sleg
		setArrElement out,"scol", scol
		setArrElement out,"chrt_array", chrt_array
		setArrElement out,"webchart", webchart
		setArrElement out,"cname", cname
		setArrElement out,"gstrOrderBy", gstrOrderBy
		setArrElement out,"table_type", table_type
		setArrElement out,"sessionPrefix", sessionPrefix
		set ASPserialize = out
	End Function
	Public Sub ASPunserialize(dict)
		DoAssignment ohcl_type, dict("ohcl_type")
		DoAssignment strSQL, dict("strSQL")
		DoAssignment label2, dict("label2")
		DoAssignment numRecordsToShow, dict("numRecordsToShow")
		DoAssignment totalRecords, dict("totalRecords")
		DoAssignment header, dict("header")
		DoAssignment footer, dict("footer")
		DoAssignment strLabel, dict("strLabel")
		DoAssignment arrDataLabels, dict("arrDataLabels")
		DoAssignment arrDataSeries, dict("arrDataSeries")
		DoAssignment arrDataColor, dict("arrDataColor")
		DoAssignment arrFormatCurrency, dict("arrFormatCurrency")
		DoAssignment arrFormatDecimal, dict("arrFormatDecimal")
		DoAssignment arrFormatCustomer, dict("arrFormatCustomer")
		DoAssignment arrFormatCustomerStr, dict("arrFormatCustomerStr")
		DoAssignment arrDataSize, dict("arrDataSize")
		DoAssignment arrAxesColor, dict("arrAxesColor")
		DoAssignment arrGaugeColor, dict("arrGaugeColor")
		DoAssignment arrOHLC_high, dict("arrOHLC_high")
		DoAssignment arrOHLC_low, dict("arrOHLC_low")
		DoAssignment arrOHLC_open, dict("arrOHLC_open")
		DoAssignment arrOHLC_close, dict("arrOHLC_close")
		DoAssignment arrOHLC_candle, dict("arrOHLC_candle")
		DoAssignment arrOHLC_color, dict("arrOHLC_color")
		DoAssignment arrOHLC_color_up, dict("arrOHLC_color_up")
		DoAssignment arrOHLC_color_down, dict("arrOHLC_color_down")
		DoAssignment sleg, dict("sleg")
		DoAssignment scol, dict("scol")
		DoAssignment chrt_array, dict("chrt_array")
		DoAssignment webchart, dict("webchart")
		DoAssignment cname, dict("cname")
		DoAssignment gstrOrderBy, dict("gstrOrderBy")
		DoAssignment table_type, dict("table_type")
		DoAssignment sessionPrefix, dict("sessionPrefix")
	End Sub
' end serialize
End Class
' Chart_Ohlc implementation 
Function method_Chart_Ohlc_write(byref this_object)
	Dim this
	ResponseWrite "<?xml version=""1.0"" standalone=""yes""?>" & vblf
	ResponseWrite "<anychart>" & vblf
	ResponseWrite "<charts>" & vblf
	this_object.write_data 
	this_object.write_dps 
	this_object.write_chart_settings 
	ResponseWrite "</chart>" & vblf
	ResponseWrite "</charts>" & vblf
	ResponseWrite "</anychart>" & vblf
End Function
Function method_Chart_Ohlc_Chart_Ohlc(byref this_object,ByRef ch_array,ByVal param)
	doClassAssignment this_object,"arrDataLabels",CreateDictionary()
	doClassAssignment this_object,"arrDataSeries",CreateDictionary()
	doClassAssignment this_object,"arrDataColor",CreateDictionary()
	doClassAssignment this_object,"arrFormatCurrency",CreateDictionary()
	doClassAssignment this_object,"arrFormatDecimal",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomer",CreateDictionary()
	doClassAssignment this_object,"arrFormatCustomerStr",CreateDictionary()
	doClassAssignment this_object,"arrDataSize",CreateDictionary()
	doClassAssignment this_object,"arrAxesColor",CreateDictionary()
	doClassAssignment this_object,"arrGaugeColor",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_high",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_low",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_open",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_close",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_candle",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_up",CreateDictionary()
	doClassAssignment this_object,"arrOHLC_color_down",CreateDictionary()
	doClassAssignment this_object,"chrt_array",CreateDictionary()
	this_object.sessionPrefix = ""
	doClassAssignment this_object,"ohcl_type",ArrayElement(param,"ohcl_type")
	this_object.sleg = "Series"
	method_Chart_Chart this_object,ch_array,param
End Function
Function method_Chart_Ohlc_write_data(byref this_object)
	Dim this
	ResponseWrite "<chart plot_type=""CategorizedVertical"">" & vblf
	ResponseWrite "<data>" & vblf
	this_object.get_data_p1 false
	ResponseWrite "</data>" & vblf
End Function
Function method_Chart_Ohlc_get_series_type(byref this_object)
	if IsEqual(this_object.ohcl_type,"ohcl") then
		method_Chart_Ohlc_get_series_type = "OHLC"
		Exit Function
	else
		method_Chart_Ohlc_get_series_type = "Candlestick"
		Exit Function
	end if
End Function
Function method_Chart_Ohlc_write_dps(byref this_object)
	Dim this
	ResponseWrite (("<data_plot_settings default_series_type=""" & CSmartStr(this_object.get_series_type())) & """>") & vblf
	this_object.get_ohcl_tooltip 
	ResponseWrite "</data_plot_settings>" & vblf
	this_object.ohls_styles 
End Function
Function method_Chart_Ohlc_ohls_styles(byref this_object)
	Dim i,attr
	ResponseWrite "<styles>" & vblf
	i = 0
	do while IsLess(i,asp_count(this_object.arrOHLC_open))
		if IsEqual(this_object.ohcl_type,"ohcl") then
			ResponseWrite (("<ohlc_style name=""style" & CSmartStr(CSmartDbl(i)+1)) & """>") & vblf
			attr = "line thickness=""1"""
		else
			ResponseWrite (("<candlestick_style name=""style" & CSmartStr(CSmartDbl(i)+1)) & """>") & vblf
			attr = "fill"
		end if
		ResponseWrite "<up>" & vblf
		ResponseWrite (((("<" & CSmartStr(attr)) & " color=""") & CSmartStr(ArrayElement(this_object.arrOHLC_color_up,i))) & """/>") & vblf
		ResponseWrite "</up>" & vblf
		ResponseWrite "<down>" & vblf
		ResponseWrite (((("<" & CSmartStr(attr)) & " color=""") & CSmartStr(ArrayElement(this_object.arrOHLC_color_down,i))) & """/>") & vblf
		ResponseWrite "</down>" & vblf
		ResponseWrite "<states>" & vblf
		ResponseWrite "<hover>" & vblf
		ResponseWrite "<up>" & vblf
		ResponseWrite (((("<" & CSmartStr(attr)) & " color=""LightColor(") & CSmartStr(ArrayElement(this_object.arrOHLC_color_up,i))) & ")""/>") & vblf
		ResponseWrite "</up>" & vblf
		ResponseWrite "<down>" & vblf
		ResponseWrite (((("<" & CSmartStr(attr)) & " color=""LightColor(") & CSmartStr(ArrayElement(this_object.arrOHLC_color_down,i))) & ")""/>") & vblf
		ResponseWrite "</down>" & vblf
		ResponseWrite "</hover>" & vblf
		ResponseWrite "</states>" & vblf
		if IsEqual(this_object.ohcl_type,"ohcl") then
			ResponseWrite "</ohlc_style>" & vblf
		else
			ResponseWrite "</candlestick_style>" & vblf
		end if
		i = CSmartDbl(i)+1
	loop
	ResponseWrite "</styles>" & vblf
End Function
Function method_Chart_Ohlc_get_ohcl_tooltip(byref this_object)
	if IsEqual(this_object.ohcl_type,"ohcl") then
		ResponseWrite "<ohlc_series>" & vblf
	else
		ResponseWrite "<candlestick_series>" & vblf
	end if
	ResponseWrite "<tooltip_settings enabled=""True"">" & vblf
	ResponseWrite "</tooltip_settings>" & vblf
	if IsEqual(this_object.ohcl_type,"ohcl") then
		ResponseWrite "</ohlc_series>" & vblf
	else
		ResponseWrite "</candlestick_series>" & vblf
	end if
End Function
Function method_Chart_Ohlc_write_chart_settings(byref this_object)
	Dim this,scroll
	ResponseWrite "<chart_settings>" & vblf
	ResponseWrite "<title enabled=""true"" padding=""15"">" & vblf
	ResponseWrite (("<text>" & CSmartStr(this_object.header)) & "</text>") & vblf
	ResponseWrite (("<font color=""#" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"color101"))) & """/>") & vblf
	ResponseWrite "</title>" & vblf
	ResponseWrite "<axes>" & vblf
	ResponseWrite "<y_axis>" & vblf
	ResponseWrite "<title>" & vblf
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"saxes"),"true") then
		ResponseWrite (("<font color=""DarkColor(" & CSmartStr(ArrayElement(this_object.arrOHLC_color,0))) & ")""/>") & vblf
	else
		ResponseWrite (("<font color=""DarkColor(" & CSmartStr(ArrayElement(this_object.arrAxesColor,0))) & ")""/>") & vblf
	end if
	ResponseWrite (("<text>" & CSmartStr(ArrayElement(this_object.arrDataLabels,0))) & "</text>") & vblf
	ResponseWrite "</title>" & vblf
	this_object.write_Logarithmic 
	ResponseWrite "</y_axis>" & vblf
	ResponseWrite "<x_axis>" & vblf
	scroll = "false"
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"cscroll"),"true") and IsLess(this_object.numRecordsToShow,this_object.totalRecords) then
		scroll = "true"
	end if
	ResponseWrite (((("<zoom enabled=""" & CSmartStr(scroll)) & """ allow_drag=""false"" visible_range=""") & CSmartStr(this_object.numRecordsToShow)) & """/>") & vblf
	ResponseWrite "<title>" & vblf
	ResponseWrite (("<text>" & CSmartStr(this_object.label2)) & "</text>") & vblf
	ResponseWrite "</title>" & vblf
	ResponseWrite "</x_axis>" & vblf
	this_object.write_extra 
	ResponseWrite "</axes>" & vblf
	this_object.write_legend 
	ResponseWrite "<chart_background>" & vblf
	this_object.write_chart_background 
	ResponseWrite "</chart_background>" & vblf
	this_object.write_plot_background 
	ResponseWrite "</chart_settings>" & vblf
End Function
Function method_Chart_Ohlc_get_data(byref this_object,ByVal refr)
	Dim arrSer,i,this,rs,j,recPerRow,row
	Set arrSer = (CreateDictionary())
	i = 0
	do while IsLess(i,asp_count(this_object.arrOHLC_open))
		setArrElement this_object.arrOHLC_color_up,i,ArrayElement(this_object.arrOHLC_color,i)
		if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"chart_type"),"type"),"candlestick") then
			setArrElement this_object.arrOHLC_color_down,i,ArrayElement(this_object.arrOHLC_candle,i)
		else
			setArrElement this_object.arrOHLC_color_down,i,ArrayElement(this_object.arrOHLC_color,i)
		end if
		setArrElement arrSer,"series" & CSmartStr(i),(((((((("<series id=""" & CSmartStr(this_object.chart_xmlencode_p1(ArrayElement(this_object.arrOHLC_open,i)))) & """ name=""") & CSmartStr(this_object.chart_xmlencode_p1(ArrayElement(this_object.arrDataLabels,i)))) & """ color=""") & CSmartStr(ArrayElement(this_object.arrOHLC_color_up,i))) & """ style=""style") & CSmartStr(CSmartDbl(i)+1)) & """>") & vblf
		i = CSmartDbl(i)+1
	loop
	doAssignmentByRef rs,db_query(this_object.strSQL,conn)
	j = 0
	doAssignment recPerRow,this_object.numRecordsToShow
	do while bValue(doAssignmentByRef(row,db_fetch_array(rs)))
		j = CSmartDbl(j)+1
		if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"cscroll"),"true") then
			recPerRow = CSmartDbl(recPerRow)+1
		end if
		if IsLess(recPerRow,j) and IsLess(0,recPerRow) then
			exit do
		end if
		i = 0
		do while IsLess(i,asp_count(this_object.arrOHLC_open))
			setArrElement arrSer,"series" & CSmartStr(i),CSmartStr(ArrayElement(arrSer,"series" & CSmartStr(i))) & (CSmartStr(this_object.get_point_p2(row,i)) & vblf)
			i = CSmartDbl(i)+1
		loop
	loop
	doClassAssignment this_object,"totalRecords",j
	i = 0
	do while IsLess(i,asp_count(this_object.arrOHLC_open))
		if bValue(refr) then
			ResponseWrite CSmartStr(ArrayElement(this_object.arrOHLC_open,i)) & vblf
			setArrElement arrSer,"series" & CSmartStr(i),asp_str_replace(CreateDictionary2(Empty,"\",Empty,vblf),CreateDictionary2(Empty,"\\",Empty,"\n"),ArrayElement(arrSer,"series" & CSmartStr(i)))
		end if
		if IsLess(0,j) then
			ResponseWrite CSmartStr(ArrayElement(arrSer,"series" & CSmartStr(i))) & "</series>"
		end if
		if not bValue(refr) or IsLess(i,CSmartDbl(asp_count(this_object.arrDataSeries))-1) then
			ResponseWrite vblf
		end if
		i = CSmartDbl(i)+1
	loop
	db_close conn
End Function
Function method_Chart_Ohlc_get_point(byref this_object,ByVal row,ByVal series)
	Dim strLabelFormat,this,str
	doAssignmentByRef strLabelFormat,this_object.labelFormat_p2(this_object.strLabel,row)
	str = ("<point name=""" & CSmartStr(strLabelFormat)) & """ "
	str = CSmartStr(str) & (((((((("high=""" & CSmartStr(this_object.chart_xmlencode_p1(CSmartDbl(ArrayElement(row,ArrayElement(this_object.arrOHLC_high,series)))+0))) & """  low=""") & CSmartStr(this_object.chart_xmlencode_p1(CSmartDbl(ArrayElement(row,ArrayElement(this_object.arrOHLC_low,series)))+0))) & """ open=""") & CSmartStr(this_object.chart_xmlencode_p1(CSmartDbl(ArrayElement(row,ArrayElement(this_object.arrOHLC_open,series)))+0))) & """ close=""") & CSmartStr(this_object.chart_xmlencode_p1(CSmartDbl(asp_str_replace(",",".",ArrayElement(row,ArrayElement(this_object.arrOHLC_close,series))))+0))) & """>")
	str = CSmartStr(str) & ((("<tooltip enabled=""True""><format>" & CSmartStr(this_object.tooltipFormat_p2(series,row))) & "</format></tooltip>") & vblf)
	str = CSmartStr(str) & ("</point>" & vblf)
	doAssignmentByRef method_Chart_Ohlc_get_point,str
	Exit Function
End Function
Function method_Chart_Ohlc_write_Logarithmic(byref this_object)
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"slog"),"true") then
		ResponseWrite "<scale type=""Logarithmic"" log_base=""10""/>" & vblf
	end if
End Function
Function method_Chart_Ohlc_write_legend_tag(byref this_object)
	ResponseWrite "<legend enabled=""true"" position=""Bottom"" ignore_auto_item=""true"" align=""Spread"" padding=""15"" height=""20%"">" & vblf
End Function
Function method_Chart_Ohlc_write_extra(byref this_object)
	Dim i,position,this
	if IsEqual(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"saxes"),"true") then
		ResponseWrite "<extra>" & vblf
		i = 1
		do while IsLess(i,asp_count(this_object.arrOHLC_open))
			doAssignment position,IIF(i mod 2=0,"Normal","Opposite")
			ResponseWrite (((("<y_axis name=""" & CSmartStr(this_object.chart_xmlencode_p1(ArrayElement(this_object.arrOHLC_high,i)))) & """ position=""") & CSmartStr(position)) & """ enabled=""true"">") & vblf
			ResponseWrite "<title enabled=""true"" align=""Center"">" & vblf
			ResponseWrite (("<text>" & CSmartStr(ArrayElement(this_object.arrDataLabels,i))) & "</text>") & vblf
			ResponseWrite (("<font color=""DarkColor(" & CSmartStr(ArrayElement(this_object.arrOHLC_color_up,i))) & ")""/>") & vblf
			ResponseWrite "</title>" & vblf
			ResponseWrite (("<labels align=""Inside"" enabled=""" & CSmartStr(ArrayElement(ArrayElement(this_object.chrt_array,"appearance"),"sval"))) & """>") & vblf
			ResponseWrite (("<font color=""DarkColor(" & CSmartStr(ArrayElement(this_object.arrOHLC_color_up,i))) & ")"" />") & vblf
			ResponseWrite (("<format>" & CSmartStr(this_object.valueFormat_p2(i,true))) & "</format>") & vblf
			ResponseWrite "</labels>" & vblf
			ResponseWrite "</y_axis>" & vblf
			i = CSmartDbl(i)+1
		loop
		ResponseWrite "</extra>" & vblf
	end if
End Function
%>
