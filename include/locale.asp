<%

'//	locale settings

'	locale settings

set locale_info = CreateObject("Scripting.Dictionary")

'	date settings
locale_info.Add "LOCALE_ICENTURY", "1"
locale_info.Add "LOCALE_IDATE", "2"
locale_info.Add "LOCALE_ILDATE", "1"
locale_info.Add "LOCALE_SDATE", "/"
locale_info.Add "LOCALE_SLONGDATE", "dd MMMM yyyy"
locale_info.Add "LOCALE_SSHORTDATE", "yyyy/MM/dd"
'	weekday names
locale_info.Add "LOCALE_IFIRSTDAYOFWEEK", "6"
locale_info.Add "LOCALE_SDAYNAME1", "Monday"
locale_info.Add "LOCALE_SDAYNAME2", "Tuesday"
locale_info.Add "LOCALE_SDAYNAME3", "Wednesday"
locale_info.Add "LOCALE_SDAYNAME4", "Thursday"
locale_info.Add "LOCALE_SDAYNAME5", "Friday"
locale_info.Add "LOCALE_SDAYNAME6", "Saturday"
locale_info.Add "LOCALE_SDAYNAME7", "Sunday"
locale_info.Add "LOCALE_SABBREVDAYNAME1", "Mon"
locale_info.Add "LOCALE_SABBREVDAYNAME2", "Tue"
locale_info.Add "LOCALE_SABBREVDAYNAME3", "Wed"
locale_info.Add "LOCALE_SABBREVDAYNAME4", "Thu"
locale_info.Add "LOCALE_SABBREVDAYNAME5", "Fri"
locale_info.Add "LOCALE_SABBREVDAYNAME6", "Sat"
locale_info.Add "LOCALE_SABBREVDAYNAME7", "Sun"
'	month names
locale_info.Add "LOCALE_SMONTHNAME1", "January"
locale_info.Add "LOCALE_SMONTHNAME2", "February"
locale_info.Add "LOCALE_SMONTHNAME3", "March"
locale_info.Add "LOCALE_SMONTHNAME4", "April"
locale_info.Add "LOCALE_SMONTHNAME5", "May"
locale_info.Add "LOCALE_SMONTHNAME6", "June"
locale_info.Add "LOCALE_SMONTHNAME7", "July"
locale_info.Add "LOCALE_SMONTHNAME8", "August"
locale_info.Add "LOCALE_SMONTHNAME9", "September"
locale_info.Add "LOCALE_SMONTHNAME10", "October"
locale_info.Add "LOCALE_SMONTHNAME11", "November"
locale_info.Add "LOCALE_SMONTHNAME12", "December"
locale_info.Add "LOCALE_SABBREVMONTHNAME1", "Jan"
locale_info.Add "LOCALE_SABBREVMONTHNAME2", "Feb"
locale_info.Add "LOCALE_SABBREVMONTHNAME3", "Mar"
locale_info.Add "LOCALE_SABBREVMONTHNAME4", "Apr"
locale_info.Add "LOCALE_SABBREVMONTHNAME5", "May"
locale_info.Add "LOCALE_SABBREVMONTHNAME6", "Jun"
locale_info.Add "LOCALE_SABBREVMONTHNAME7", "Jul"
locale_info.Add "LOCALE_SABBREVMONTHNAME8", "Aug"
locale_info.Add "LOCALE_SABBREVMONTHNAME9", "Sep"
locale_info.Add "LOCALE_SABBREVMONTHNAME10", "Oct"
locale_info.Add "LOCALE_SABBREVMONTHNAME11", "Nov"
locale_info.Add "LOCALE_SABBREVMONTHNAME12", "Dec"
'	time settings
locale_info.Add "LOCALE_ITIME", "0"
locale_info.Add "LOCALE_ITIMEMARKPOSN", "0"
locale_info.Add "LOCALE_ITLZERO", "1"
locale_info.Add "LOCALE_S1159", "AM"
locale_info.Add "LOCALE_S2359", "PM"
locale_info.Add "LOCALE_STIME", ":"
locale_info.Add "LOCALE_STIMEFORMAT", "hh:mm:ss tt"
'	currency settings
locale_info.Add "LOCALE_ICURRDIGITS", "2"
locale_info.Add "LOCALE_ICURRENCY", "2"
locale_info.Add "LOCALE_INEGCURR", "2"
locale_info.Add "LOCALE_SCURRENCY", "R"
locale_info.Add "LOCALE_SMONDECIMALSEP", ","
locale_info.Add "LOCALE_SMONGROUPING", "3;0"
locale_info.Add "LOCALE_SMONTHOUSANDSEP", " "
'	numbers formatting settings
locale_info.Add "LOCALE_IDIGITS", "2"
locale_info.Add "LOCALE_INEGNUMBER", "1"
locale_info.Add "LOCALE_SDECIMAL", "."
locale_info.Add "LOCALE_SGROUPING", "3;0"
locale_info.Add "LOCALE_SNEGATIVESIGN", "-"
locale_info.Add "LOCALE_SPOSITIVESIGN", ""
locale_info.Add "LOCALE_STHOUSAND", " "


locale_info("LOCALE_ILONGDATE")=GetLongDateFormat()
 
'//	locale functions
'//	number, currency, date & time functions


function str_format_number(val,valDigits)
	dim sign, vint, frac, out, ptr, gi, fmul, i, sfrac, iDigits
	dim grouping
	if not isNumeric(val) then 
		str_format_number=val
		exit function
	end if
	iDigits=valDigits
	if vartype(valDigits)=vbBoolean then
		if iDigits=false then iDigits=locale_info("LOCALE_IDIGITS")
	end if
	val = round(val,iDigits)
	if val>=0 then
	  sign=1
	  vint = int(val)
	  frac = val-vint
	else 
	  sign=-1
	  vint = int(-val)
	  frac = -val-vint
	end if
	out = formatnumber(vint,0)
'//	grouping
    grouping=split(locale_info("LOCALE_SGROUPING"),";")
	if uBound(grouping)>0 and grouping(0)<>"" then
		ptr=len(out)
		for gi=0 to uBound(grouping)-1
			if not grouping(gi)<>"" then gi=gi-1
			if ptr<=grouping(gi) then
				ptr=0
				exit for
			end if
			out=substr(out,1,ptr-grouping(gi)) & locale_info("LOCALE_STHOUSAND") & substr(out,ptr-grouping(gi))
			ptr=ptr-grouping(gi)
		next
	end if
''//	fractional digits
    if iDigits>0 then
      fmul=1
      for i=0 to iDigits-1
        fmul=fmul*10
      next
      frac=round(frac*fmul)
	  sfrac=cstr(frac)
	  dl=len(sfrac)
	  while dl<cint(iDigits)
	    sfrac="0" & sfrac
	    dl=dl+1
	  wend
	  out=out & locale_info("LOCALE_SDECIMAL") & sfrac
    end if
''//	format output
	if sign>0 then
		str_format_number = locale_info("LOCALE_SPOSITIVESIGN") & out
		exit function
	else
		select case locale_info("LOCALE_INEGNUMBER")
			case 0 str_format_number = "(" & out & ")"
								exit function
			case 1 str_format_number = "-" & out
								exit function
			case 2 str_format_number = "- " & out
								exit function
			case 3 str_format_number = out & "-"
								exit function
			case 4 str_format_number = out & " -"
								exit function
		end select
	end if
	str_format_number=val
end function


function str_format_currency(val)
	if isnull(val) then 
	        str_format_currency=""
        	exit function
	end if
	str_format_currency = FormatCurrency(CSmartDbl(val))
	exit function
	dim sign, vint, frac, out, ptr, gi, fmul, sfrac
	dim grouping
	if not isNumeric(val) then 
	    str_format_currency = val
	    exit function
	end if
	if val>=0 then
		sign=1
		vint = int(val)
		frac = val-vint
	else 
	  sign=-1
	  vint = int(-val)
	  frac = -val-vint
	end if
	out = formatnumber(vint,0)
'//	grouping
    grouping=split(locale_info("LOCALE_SMONGROUPING"),";")
	if uBound(grouping)>0 and grouping(0)<>"" then
		ptr=len(out)
		for gi=0 to uBound(grouping)-1
			if not grouping(gi)<>"" then gi=gi-1
			if ptr<=grouping(gi) then
				ptr=0
				exit for
			end if
			out=substr(out,1,ptr-grouping(gi)) & locale_info("LOCALE_SMONTHOUSANDSEP") & substr(out,ptr-grouping(gi))
			ptr=ptr-grouping(gi)
		next
	end if
'//	fractional digits
    if locale_info("LOCALE_ICURRDIGITS")>0 then
      fmul=1
      for i=0 to locale_info("LOCALE_ICURRDIGITS")-1
        fmul=fmul*10
      next
      frac=round(frac*fmul)
	  sfrac=cstr(frac)
	  dl=len(sfrac)
	  while dl<cint(locale_info("LOCALE_ICURRDIGITS"))
	    sfrac="0" & sfrac
	    dl=dl+1
	  wend
	  out=out & locale_info("LOCALE_SMONDECIMALSEP") & sfrac
    end if
'//	format output
	if sign>0 then
		select case locale_info("LOCALE_ICURRENCY")
			case 0 str_format_currency = cstr(locale_info("LOCALE_SCURRENCY")) & cstr(out)
								exit function
			case 1 str_format_currency = cstr(out) & cstr(locale_info("LOCALE_SCURRENCY"))
								exit function
			case 2 str_format_currency = cstr(locale_info("LOCALE_SCURRENCY")) & " " & cstr(out)
								exit function
			case 3 str_format_currency = cstr(out) & " " & cstr(locale_info("LOCALE_SCURRENCY"))
								exit function
		end select
	else
		select case locale_info("LOCALE_INEGCURR")
			case 0 str_format_currency = "(" & cstr(locale_info("LOCALE_SCURRENCY")) & cstr(out) & ")"
								exit function
			case 1 str_format_currency = "-" & cstr(locale_info("LOCALE_SCURRENCY")) & cstr(out)
								exit function
			case 2 str_format_currency = cstr(locale_info("LOCALE_SCURRENCY")) & "-" & cstr(out)
								exit function
			case 3 str_format_currency = cstr(locale_info("LOCALE_SCURRENCY")) & cstr(out)
								exit function
			case 4 str_format_currency = "(" & cstr(out) & cstr(locale_info("LOCALE_SCURRENCY")) & ")"
								exit function
			case 5 str_format_currency = "-" & cstr(out) & cstr(locale_info("LOCALE_SCURRENCY"))
								exit function
			case 6 str_format_currency = cstr(out) & "-" & cstr(locale_info("LOCALE_SCURRENCY"))
								exit function
			case 7 str_format_currency = cstr(out) & cstr(locale_info("LOCALE_SCURRENCY")) & "-"
								exit function
			case 8 str_format_currency = "-" & cstr(out) & " " & cstr(locale_info("LOCALE_SCURRENCY"))
								exit function
			case 9 str_format_currency = "-" & cstr(locale_info("LOCALE_SCURRENCY")) & " " & cstr(out)
								exit function
			case 10 str_format_currency = cstr(out) & " " & cstr(locale_info("LOCALE_SCURRENCY")) & "-"
								exit function
			case 11 str_format_currency = cstr(locale_info("LOCALE_SCURRENCY")) & " " & cstr(out) & "-"
								exit function
			case 12 str_format_currency = cstr(locale_info("LOCALE_SCURRENCY")) & " -" & cstr(out)
								exit function
			case 13 str_format_currency = cstr(out) & "- " & cstr(locale_info("LOCALE_SCURRENCY"))
								exit function
			case 14 str_format_currency = "(" & cstr(locale_info("LOCALE_SCURRENCY")) & " " & cstr(out) & ")"
								exit function
			case 15 str_format_currency = "(" & cstr(out) & " " & cstr(locale_info("LOCALE_SCURRENCY")) & ")"
								exit function
		end select
	end if
	str_format_currency = val
end function


'//	converts mysql datetime to array(year,month,day,hour,minute,second)
function db2time(val)
	dim arr
	set arr=server.createObject("Scripting.Dictionary")
	if isnull(val) then
		set db2time=arr
		exit function
	end if
	if isdate(val) then
		arr(0)=year(val)
		arr(1)=month(val)
		arr(2)=day(val)
		arr(3)=hour(val)
		arr(4)=minute(val)
		arr(5)=second(val)
		set db2time=arr
		exit function
	end if
	str=CStr(val)
	dim isdst, havedate, havetime, pattern, y, mo, d, h, mi, s, vlen
	dim vnow(3)
	Dim regEx, Match, Matches
	Dim matchesCount
	Set regEx = New RegExp
	regEx.IgnoreCase = True
	regEx.Global = True
	pattern=""
	vnow(0)=year(now)
	vnow(1)=month(now)
	vnow(2)=day(now)
	
    havedate=0
	havetime=0
	if isNumeric(str) then
'//	timestamp
		havedate=1
		vlen=len(str)
		if vlen>=10 then havetime=1
		select case vlen
		  case 14 pattern="(\d{4})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})"
		  case 12 pattern="(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})"
		  case 10 pattern="(\d{2})(\d{2})(\d{2})(\d{2})(\d{2})"
		  case 8 pattern="(\d{4})(\d{2})(\d{2})"
		  case 6 pattern="(\d{2})(\d{2})(\d{2})"
		  case 4 pattern="(\d{2})(\d{2})"
		  case 2 pattern="(\d{2})"
	      case else
	        set db2time = arr
		    exit function
	    end select
	    regEx.Pattern = pattern

		Set Matches = regEx.Execute(str)
		matchesCount = Matches.Count

		If matchesCount > 0 Then 
			set m=Matches(0)
			y = m.subMatches(0).Value

			If matchesCount > 1 Then
				mo = m.subMatches(1)
			Else
				mo = 1
			End If
			If matchesCount > 2 Then
				d = m.subMatches(2)
			Else
				d = 1
			End If
			If matchesCount > 3 Then
				h = m.subMatches(3)
			Else
				h = 0
			End If
			If matchesCount > 4 Then
				mi = m.subMatches(4)
			Else
				mi = 0
			end if
			If matchesCount > 5 Then
				s = m.subMatches(5)
			Else
				s = 0
			End If
		else
			set db2time = arr
		    exit function
		end if
	else 
		if not isNumeric(str) and not isnull(str) and trim(str)<>"" then
	'// date,time,datetime
'			str=year(str) & "-" & month(str) & "-" & day(str) & " " & hour(str) & ":" & minute(str) & ":" & second(str)
			regEx.Pattern = "(\d{4})-(\d{1,2})-(\d{1,2}) (\d{1,2}):(\d{1,2}):(\d{1,2})"

			Set Matches = regEx.Execute(str)
			matchesCount = Matches.Count

			If matchesCount > 0 Then 
				set m=Matches(0)
				y = m.subMatches(0)
				mo = m.subMatches(1)
				d = m.subMatches(2)
				h = m.subMatches(3)
				mi = m.subMatches(4)
				s = m.subMatches(5)
				havedate=1
				havetime=1
			else 
				regEx.Pattern = "(\d{4})-(\d{1,2})-(\d{1,2})"
				Set Matches = regEx.Execute(str)
				matchesCount = Matches.Count

				If matchesCount > 0 Then 
					set m=Matches(0)
					y = m.subMatches(0)
					mo = m.subMatches(1)
					d = m.subMatches(2)
					h = 0
					mi = 0
					s = 0
					havedate=1
				else 
					regEx.Pattern = "(\d{2})-(\d{1,2})-(\d{1,2})"
					Set Matches = regEx.Execute(str)
					matchesCount = Matches.Count

					If matchesCount > 0 Then
						set m=Matches(0)
						y=vnow(0)
						mo=vnow(1)+1
						d=vnow(2)
						h = m.subMatches(3)
						mi = m.subMatches(4)
						s = m.subMatches(5)
						havetime=1
					else 
						set db2time = arr
						exit function
					end if
				end if
			end if
		else
			set db2time = arr
			exit function
		end if
	end if
	if havetime=0 then
		h=0
		mi=0
		s=0
	end if
	if havedate=0 then
		y=vnow(0)
		mo=vnow("1")+1
		d=vnow("2")
	end if
	arr(0)=y
	arr(1)=mo
	arr(2)=d
	arr(3)=h
	arr(4)=mi
	arr(5)=s
	set db2time = arr
end function

function str_format_datetime(ttime())
	str_format_datetime = format_datetime_custom(ttime,locale_info("LOCALE_SSHORTDATE") & " " & locale_info("LOCALE_STIMEFORMAT"))
end function

function str_format_time(ttime())
	str_format_time = format_datetime_custom(ttime,locale_info("LOCALE_STIMEFORMAT"))
end function

function format_shortdate(ttime())
	format_shortdate = format_datetime_custom(ttime,locale_info("LOCALE_SSHORTDATE"))
end function

function format_longdate(ttime())
	format_longdate = format_datetime_custom(ttime,locale_info("LOCALE_SLONGDATE"))
end function

function simpledate2db(strdate,formatid)
	dim sstr, mmonth, dday, yyear
	dim numbers
	sstr=strdate
	set numbers=parsenumbers(sstr)
	if numbers.count=0 then simpledate2db = strdate 
	while numbers.count<3
		numbers(numbers.count)=1
	wend
	if formatid=0 then
		mmonth=numbers(0)
		dday=numbers(1)
		yyear=numbers(2)
	else 
		if formatid=1 then
			dday=numbers(0)
			mmonth=numbers(1)
			yyear=numbers(2)
		else 
			if formatid=2 then
				yyear=numbers(0)
				mmonth=numbers(1)
				dday=numbers(2)
			else
				simpledate2db = strdate
			end if
		end if
	end if
	if yyear<100 then
		if yyear<60 then
			yyear=yyear+2000
		else
			yyear=yyear+1900
		end if
	end if
	simpledate2db = yyear & "-" & mmonth & "-" & dday
end function


function localdate2db(strdate)
	localdate2db = simpledate2db(strdate,locale_info("LOCALE_IDATE"))
end function

function localtime2db(strtime)
'//	check if we use 12hours clock
	dim use12, pos, pm, amstr, pmstr, str, h, k
	dim numbers
	use12=0
	pos=instr(1,locale_info("LOCALE_STIMEFORMAT"),"h" & locale_info("LOCALE_STIME"))
	if pos>0 then 
		use12=1
'	determine am/pm
		pm=0
		amstr=locale_info("LOCALE_S1159")
		pmstr=locale_info("LOCALE_S2359")
		posam=instr(1,strtime,amstr)
		pospm=instr(1,strtime,pmstr)
		
		if posam=0 and pospm>0 then
			pm=1
		elseif posam>0 and pospm=0 then
			pm=0
		elseif posam=0 and pospm=0 then
			use12=0
		else
			if posam>pospm then _
				pm=1
		end if
	end if
	str=strtime
	set numbers=parsenumbers(str)
	while numbers.count<3
		numbers(numbers.count)=0
	wend
	h=numbers(0)
	m=numbers(1)
	s=numbers(2)
	if use12<>0 and h<>0 then
		if pm=0 and h=12 then h=0
		if pm=1 and h<12 then h=h+12
	end if
	localtime2db = cstr(h) & ":" & cstr(m) & ":" & cstr(s)
end function


function localdatetime2db(strdatetime,format)
	dim use12, locale_idate, pm, amstr, pmstr, pos, tm, mmonth, dday, yyear, h, m, s, l
	dim numbers
	locale_idate=locale_info("LOCALE_IDATE")
	if format="dmy" then locale_idate=1
	if format="mdy" then locale_idate=0
	if format="ymd" then locale_idate=2
'	check if we use 12hours clock
	use12=0
	pos=instr(1,locale_info("LOCALE_STIMEFORMAT"),"h" & locale_info("LOCALE_STIME"))
	if pos>0 then 
		use12=1
'	determine am/pm
		pm=0
		amstr=locale_info("LOCALE_S1159")
		pmstr=locale_info("LOCALE_S2359")
		posam=instr(1,strdatetime,amstr)
		pospm=instr(1,strdatetime,pmstr)
		
		if posam=0 and pospm>0 then
			pm=1
		elseif posam>0 and pospm=0 then
			pm=0
		elseif posam=0 and pospm=0 then
			use12=0
		else
			if posam>pospm then _
				pm=1
		end if
	end if
	set numbers=parsenumbers(strdatetime)
	if numbers.count<2 then	
		localdatetime2db = "null"
		exit function
	end if
'	add current year if not specified
	if numbers.count<3 then
		if locale_idate<>1 then
			mmonth=numbers(0)
			dday=numbers(1)
		else
			mmonth=numbers(1)
			dday=numbers(0)
		end if
		yyear=year(now)
	else
		if locale_idate=0 then	
			mmonth=numbers(0)
			dday=numbers(1)
			yyear=numbers(2)
		else 
			if locale_idate=1 then
				dday=numbers(0)
				mmonth=numbers(1)
				yyear=numbers(2)
			else 
				if locale_idate=2 then
					yyear=numbers(0)
					mmonth=numbers(1)
					dday=numbers(2)
				end if
			end if
		end if
	end if
	if mmonth=0 or dday=0 then 
		localdatetime2db = "null"
		exit function
	end if
	while numbers.count<6
		numbers(numbers.count)=0
	wend
	h=numbers(3)
	m=numbers(4)
	s=numbers(5)
	if use12=1 and h<>0 then
		if pm=0 and h=12 then h=0
		if pm=1 and h<12 then h=h+12
	end if
	if yyear<100 then
		if yyear<60 then 
			yyear=yyear+2000
		else
			yyear=yyear+1900
		end if
	end if
	localdatetime2db = yyear & "-" & mmonth & "-" & dday & " " & h & ":" & m & ":" & s
end function

function parsenumbers(str)
	dim i, num, pos, j
	dim ret
	set ret=server.createObject("Scripting.Dictionary")
	i=1
	num=0
	pos=1
	j=0
	if isnull(str) then
		set parsenumbers = ret
		exit function
	end if
	if len(str)=0 then
		set parsenumbers = ret
		exit function
	end if
	while i<=len(str)
		if isNumeric(mid(str,i,1)) and num=0 then
			num=1
			pos=i
		else 
			if not isNumeric(mid(str,i,1)) and num<>0 then
				ret(j)=cint(mid(str,pos,i-pos))
				j=j+1
				num=0
			end if
		end if
		i=i+1
	wend
	if num<>0 then 
		ret(j)=cint(mid(cstr(str),pos,i-pos+1))
		j=j+1
	end if
	set parsenumbers = ret
end function

'//	returns day of week (1-7) for (monday-sunday)
function format_datetime_custom(ttime,format)
	dim i,weekday, hour12, am, out, inquot, n, key
	dim keys
	Set subst = CreateObject("Scripting.Dictionary")
	if isnull(ttime) then
		format_datetime_custom = ""
		exit function
	else
		if asp_count(ttime)<3 or ttime(0)="" then 
			format_datetime_custom = ""
			exit function
		end if
	end if
	if ttime(1)=0 then _
		ttime(1)=1
	i=1
	weekday=getdayofweek(ttime)

	subst.Add "dddd",locale_info("LOCALE_SDAYNAME" & weekday)
	subst.Add "ddd",locale_info("LOCALE_SABBREVDAYNAME" & weekday)
	if len(cstr(ttime(2)))=1 then 
		subst.Add "dd","0" & cstr(ttime(2))
	else
		subst.Add "dd",cstr(ttime(2))
	end if
	subst.Add "d",ttime(2)
	subst.Add "MMMM",locale_info("LOCALE_SMONTHNAME" & ttime(1))
	subst.Add "MMM",locale_info("LOCALE_SABBREVMONTHNAME" & ttime(1))
	if len(cstr(ttime(1)))=1 then 
		subst.Add "MM","0" & cstr(ttime(1))
	else
		subst.Add "MM",cstr(ttime(1))
	end if
	subst.Add "M",ttime(1)

	var = CStr(ttime(0))
	while len(var)<4
	     var = "0" & var
	wend 
	subst.Add "yyyy", var
	
	var = CStr((ttime(0) mod 100))
	while len(var)<2
	     var = "0" & var
	wend 
	subst.Add "yy", var

	subst.Add "y",(ttime(0) mod 10)
	subst.Add "gg",""
	if len(cstr(ttime(3)))=1 then 
		subst.Add "HH","0" & cstr(ttime(3))
	else
		subst.Add "HH",cstr(ttime(3))
	end if
	subst.Add "H",ttime(3)
	if len(cstr(ttime(4)))=1 then 
		subst.Add "mm","0" & cstr(ttime(4))
	else
		subst.Add "mm",cstr(ttime(4))
	end if
	subst.Add "m",ttime(4)
	if len(cstr(ttime(5)))=1 then 
		subst.Add "ss","0" & cstr(ttime(5))
	else
		subst.Add "ss",cstr(ttime(5))
	end if
	subst.Add "s",ttime(5)
	hour12=ttime(3)
	am=1
	if hour12>=12 then
		am=0
		hour12=hour12-12
	end if
	if hour12=0 then hour12=12
	subst.Add "hh",cstr(hour12)
	subst.Add "h",hour12
	if am=1 then
		subst.Add "tt",locale_info("LOCALE_S1159")
		subst.Add "t",mid(locale_info("LOCALE_S1159"),1,1)
	else
		subst.Add "tt",locale_info("LOCALE_S2359")
		subst.Add "t",mid(locale_info("LOCALE_S2359"),1,1)
	end if
	out=format
	inquot=0
	while i<=len(out)
		if mid(out,i,1)="'" then
			inquot=1-inquot
			out=mid(out,1,i) & mid(out,i+2)
			flag=1
		else 
			if inquot=0 then
				for each key in subst
					if mid(out,i,len(key))=key then
						out=mid(out,1,i-1) & subst(key) & mid(out,len(key)+i)
						i=i+len(subst(key))-1
						exit for
					end if
				next
			end if
		end if
		i=i+1
	wend
	format_datetime_custom = out
end function

function getdayofweek(ttime())
	dim daydif, i
'//	January 1, 2004 - Thursday
'//	Get the differewnce in days between January 1, 2004 and January 1 of given year
	daydif=0
	if ttime(0)>=2004 then
		for i=2004 to ttime(0)-1
			if isleapyear(i) then
				daydif=daydif+366
			else
				daydif=daydif+365
			end if
		next
	else
		for i=2003 to ttime(0) step -1
			if isleapyear(i) then
				daydif=daydif-366
			else
				daydif=daydif-365
			end if
		next
	end if
'//	to given month
	dim mdays(13)
	mdays(1)=31
	mdays(2)=28
	mdays(3)=31
	mdays(4)=30
	mdays(5)=31
	mdays(6)=30
	mdays(7)=31
	mdays(8)=31
	mdays(9)=30
	mdays(10)=31
	mdays(11)=30
	mdays(12)=31
	
	if isleapyear(ttime(0)) then mdays(2)=29
	
	for i=1 to ttime(1)-1
		daydif=daydif+mdays(i)
	next
'//	to given day
	daydif=daydif+ttime(2)-1
	if daydif>0 then 
		getdayofweek = (4+daydif-1) mod 7 + 1
		exit function
	end if
	getdayofweek = 7-(3-daydif) mod 7
end function

function getweekstart(ttime)
	dim wday,diff
	wday = getdayofweek(ttime)
	if wday>=locale_info("LOCALE_IFIRSTDAYOFWEEK")+1 then
		diff=wday - locale_info("LOCALE_IFIRSTDAYOFWEEK")-1
	else
		diff=wday+7 - locale_info("LOCALE_IFIRSTDAYOFWEEK")-1
	end if
	set getweekstart=adddays(ttime,-diff)
end function


function adddays(byref pttime,days)
	dim mdays(13),ttime
	copyDictionary pttime,ttime
	mdays(1)=31
	mdays(2)=28
	mdays(3)=31
	mdays(4)=30
	mdays(5)=31
	mdays(6)=30
	mdays(7)=31
	mdays(8)=31
	mdays(9)=30
	mdays(10)=31
	mdays(11)=30
	mdays(12)=31
	
	if isleapyear(ttime(0)) then mdays(2)=29

	if days>0 then
		for i=1 to days 
			if ttime(2)<mdays(ttime(1)) then
				ttime(2) = ttime(2)+1
			else
				ttime(2)=1
				ttime(1) = ttime(1) + 1
				if ttime(1)>12 then
					ttime(1) = 1
					ttime(0) = ttime(0)+1
					if isleapyear(ttime(0)) then 
						mdays(2)=29
					else
						mdays(2)=28
					end if
				end if
			end if
		next
	else
		for i=1 to -days 
			if ttime(2)>1 then
				ttime(2) = ttime(2) - 1
			else
				ttime(1) = ttime(1) - 1
				if ttime(1)<1 then
					ttime(0) = ttime(0) - 1
					if isleapyear(ttime(0)) then 
						mdays(2)=29
					else
						mdays(2)=28
					end if
					ttime(1)=12
				end if
				ttime(2)=mdays(ttime(1))
			end if
		next
	end if
	set adddays=ttime
end function



function getweeknumber(ttime)
	dim daydif, i
	if locale_info("LOCALE_IFIRSTDAYOFWEEK")<=3 then
		startweekday=3-locale_info("LOCALE_IFIRSTDAYOFWEEK")
	else
		startweekday=10-locale_info("LOCALE_IFIRSTDAYOFWEEK")
	end if

'//	January 1, 2004 - Thursday
'//	Get the differewnce in days between January 1, 2004 and January 1 of given year
	daydif=0
	if ttime(0)>=2004 then
		for i=2004 to ttime(0)-1
			if isleapyear(i) then
				daydif=daydif+366
			else
				daydif=daydif+365
			end if
		next
	else
		for i=2003 to ttime(0) step -1
			if isleapyear(i) then
				daydif=daydif-366
			else
				daydif=daydif-365
			end if
		next
	end if
'//	to given month
	dim mdays(13)
	mdays(1)=31
	mdays(2)=28
	mdays(3)=31
	mdays(4)=30
	mdays(5)=31
	mdays(6)=30
	mdays(7)=31
	mdays(8)=31
	mdays(9)=30
	mdays(10)=31
	mdays(11)=30
	mdays(12)=31
	
	if isleapyear(ttime(0)) then mdays(2)=29
	
	for i=1 to ttime(1)-1
		daydif=daydif+mdays(i)
	next
'//	to given day
	daydif=daydif+ttime(2)-1

	daydif=daydif+startweekday
	daydif = daydif-(daydif mod 7)
	getweeknumber=daydif/7
end function


function isleapyear(y)
	if y mod 4 <>0 then 
		isleapyear = false
		exit function
	end if
	if y mod 100 <>0 then 
		isleapyear = true
		exit function
	end if
	if (y/100) mod 4 <> 0 then 
		isleapyear = false
		exit function
	end if
	isleapyear = true
end function

function GetLongDateFormat()
	dim format, dstart, inquote, dindex, mindex, yindex, i, c, f

	format=locale_info("LOCALE_SLONGDATE")

'//	dd,d - day
'//	MMMM, MMM, MM, M - month
'//	yyyy, yy, y - year
'//	dddd, ddd - day of week, ignore it
'//	'sdsd' - quoted string, ignore it.
	dstart=-1
	inquote=false
	dindex=-1
	mindex=-1
	yindex=-1
	i=0
	f=1
	while f=1
		c=""
		if i<len(format) then c=mid(format,i+1,1)
		if dstart>=0 and c<>"d" then
			if i-dstart<=2 then dindex=dstart
			dstart=-1
		end if
		if not inquote and c="\'" then
			inquote=true
		else 
			if c="\'" then
				inquote=false
			else 
				if not inquote then
					if dindex<0 and c="d" then
						if dstart<0 then dstart=i
					end if
					if yindex<0 and c="y" then yindex=i
					if mindex<0 and c="M" then mindex=i
				end if
			end if
		end if
		if i>=len(format) then f=0
		i=i+1
	wend
	if dindex<0 or mindex<0 or yindex<0 then 
		GetLongDateFormat = -1
		exit function
	end if
	if dindex<mindex and mindex<yindex then 
		GetLongDateFormat = 1	'// DMY 
		exit function
	end if
	if mindex<dindex and dindex<yindex then 
		GetLongDateFormat = 0	'// MDY
		exit function
	end if
	if yindex<mindex and mindex<dindex then 
		GetLongDateFormat = 2	'// YMD
		exit function
	end if
	if yindex<dindex and dindex<mindex then 
		GetLongDateFormat = 1	'// YDM
		exit function
	end if
	GetLongDateFormat = -1
end function

%>
