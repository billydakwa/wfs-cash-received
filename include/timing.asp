<%

set t_timings = Server.CreateObject("Scripting.Dictionary")
set t_counts = Server.CreateObject("Scripting.Dictionary")

class Timing
	private func
	private start
	public sub SetFuncName(fname)
		func=fname
		start= timer()
	end sub
	private sub Class_Terminate()
		dim tm
		tm=timer()
		if not t_timings.Exists(func) then
			t_timings(func)=tm-start
			t_counts(func)=1
		else
			t_timings(func)=t_timings(func) + tm-start
			t_counts(func)=t_counts(func)+1
		end if
	end sub
end class

class PrintTiming
	private start
	private sub Class_Initialize()
		start=timer()
	end sub
	private sub Class_Terminate()
		dim keys,k,n,tt
		keys = t_timings.keys
		n=0.0
		tt=timer()
'		for each k in keys
'			n = n + t_counts(k)*0.00002
'		next
		response.write "<br><br><br><br><br><br><br>"
		response.write "<br>Total: " & formatnumber(tt-start-n,3)
		response.write "<table border=0>"
		for each k in keys
			response.write "<tr><td>" & k & "<td>" & formatnumber(t_timings(k),2) & "<td>" & t_counts(k) & "<td>" & formatnumber(t_timings(k)*1000/t_counts(k),2)
		next
		response.write "</table>"
	end sub
end class

set t_print = new PrintTiming
%>