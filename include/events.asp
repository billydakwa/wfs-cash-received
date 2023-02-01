<%

'------ Class eventsBase ------
Class eventsBase
	Public events
	Public captchas
	public function exists_p1(ByVal var_event)
		exists_p1 = eventsBase_exists(me,var_event)
	end function 
	public function existsCAPTCHA_p1(ByVal page)
		existsCAPTCHA_p1 = eventsBase_existsCAPTCHA(me,page)
	end function 
	public function callCAPTCHA_p1(ByRef pageObject)
		eventsBase_callCAPTCHA me,pageObject
	end function 
	Public Function init_eventsBase()
		set events = CreateDictionary()
		set captchas = CreateDictionary()
	end function
end class

Function eventsBase_exists(byref obj, ByVal var_event)
	eventsBase_exists = not IsFalse(asp_array_key_exists(var_event,obj.events))
End Function

Function eventsBase_existsCAPTCHA(byref obj, ByVal page)
	eventsBase_existsCAPTCHA = not IsFalse(asp_array_key_exists(page,obj.captchas))
End Function

Function eventsBase_callCAPTCHA(byref obj, ByRef pageObject)
	if IsEqual(pageObject.pageType,"add") then
		obj.displayCaptchaOnAdd 
	else
		if IsEqual(pageObject.pageType,"edit") then
			obj.displayCaptchaOnEdit 
		else
			if IsEqual(pageObject.pageType,"register") then
				obj.displayCaptchaOnRegister 
			end if
		end if
	end if
End Function


'------ Class class_GlobalEvents extends eventsBase ------
Class class_GlobalEvents
	Public events
	Public captchas
	public function exists_p1(ByVal var_event)
		exists_p1 = eventsBase_exists(me,var_event)
	end function 
	public function existsCAPTCHA_p1(ByVal page)
		existsCAPTCHA_p1 = eventsBase_existsCAPTCHA(me,page)
	end function 
	public function callCAPTCHA_p1(ByRef pageObject)
		eventsBase_callCAPTCHA me,pageObject
	end function 
	Public Function init_class_GlobalEvents()
	set events = CreateDictionary()
	set captchas = CreateDictionary()
	End Function
'	handlers
'	handler wrappers

'	onscreen events

End Class

%>