<%
	function search_cyear2cmonth2(cyear, cmonth, cyear2, cmonth2)
		if cyear =  "" then cyear = Cstr(Year(date))
		if cmonth = "" then cmonth = Cstr(Month(Date))
		if cyear2 =  "" then cyear2 = Cstr(Year(date))
		if cmonth2 = "" then cmonth2 = Cstr(Month(Date))
		dim intLoop
		dim strIntLoop

		response.write "<select id='cyear' name='cyear' >"
		for intLOop = 2008 to YEar(date)
		response.write "<option value="&intLOop&"
		if cyear = CSTR(intLoop) then response.write " selected"
		response.write ">"
		response.write "</select>"

		response.write "<select id='cyear2' name='cyear2' >"
		for intLOop = 2008 to YEar(date)
		response.write "<option value="&intLOop&"
		if cyear = CSTR(intLoop) then response.write " selected"
		response.write ">"
		response.write "</select>"

		response.write "<select id='cmonth' name='cmonth' >"
		for intLOop = 1 to 12
		if Len(intLoop) = 1 then strIntLoop = "0" & cmonth else strIntLoop = intLOop
		response.write "<option value="&strIntLoop&"
		if cmonth = CSTR(strIntLoop) then response.write " selected"
		response.write ">"
		response.write "</select>"

		response.write "<select id='cmonth2' name='cmonth2' >"
		for intLOop = 1 to 12
		if Len(intLoop) = 1 then strIntLoop = "0" & intloop else strIntLoop = intLOop
		response.write "<option value="&strIntLoop&"
		if cmonth2 = CSTR(strIntLoop) then response.write " selected"
		response.write ">"
		response.write "</select>"

	end function

	function search_cyearcmonth2(cyar, cmonth, cmonth2)
		if cyear =  "" then cyear = Cstr(Year(date))
		if cmonth = "" then cmonth = Cstr(Month(Date))
		if cmonth2 = "" then cmonth2 = Cstr(Month(Date))
		dim intLoop
		dim strIntLoop

		response.write "<select id='cyear' name='cyear' >"
		for intLOop = 2008 to YEar(date)
		response.write "<option value="&intLOop&"
		if cyear = CSTR(intLoop) then response.write " selected"
		response.write ">"
		response.write "</select>"


		response.write "<select id='cmonth' name='cmonth' >"
		for intLOop = 1 to 12
		if Len(intLoop) = 1 then strIntLoop = "0" & cmonth else strIntLoop = intLOop
		response.write "<option value="&strIntLoop&"
		if cmonth = CSTR(strIntLoop) then response.write " selected"
		response.write ">"
		response.write "</select>"

		response.write "<select id='cmonth2' name='cmonth2' >"
		for intLOop = 1 to 12
		if Len(intLoop) = 1 then strIntLoop = "0" & intloop else strIntLoop = intLOop
		response.write "<option value="&strIntLoop&"
		if cmonth2 = CSTR(strIntLoop) then response.write " selected"
		response.write ">"
		response.write "</select>"
	end function

	function search_cyearcmonth(cyear, cmonth)
		if cyear =  "" then cyear = Cstr(Year(date))
		if cmonth = "" then cmonth = Cstr(Month(Date))
		dim intLoop
		dim strIntLoop

		response.write "<select id='cyear' name='cyear' >"
		for intLOop = 2008 to YEar(date)
		response.write "<option value="&intLOop&"
		if cyear = CSTR(intLoop) then response.write " selected"
		response.write ">"
		response.write "</select>"


		response.write "<select id='cmonth' name='cmonth' >"
		for intLOop = 1 to 12
		if Len(intLoop) = 1 then strIntLoop = "0" & cmonth else strIntLoop = intLOop
		response.write "<option value="&strIntLoop&"
		if cmonth = CSTR(strIntLoop) then response.write " selected"
		response.write ">"
		response.write "</select>"

	end function

	function search_cyear(cyear)
		if cyear =  "" then cyear = Cstr(Year(date))
		dim intLoop

		response.write "<select id='cyear' name='cyear' >"
		for intLOop = 2008 to YEar(date)
		response.write "<option value="&intLOop&"
		if cyear = CSTR(intLoop) then response.write " selected"
		response.write ">"
		response.write "</select>"

	end function

	function search_custcode(Custcode)
		dim sql: sql = "select highcustcode, custname from sc_Cust_hdr where medflag='a' order by custname"
		dim cmd : set cmd = server.createobject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandTYpe = adcmdtext
		cmd.commandtext = sql
		dim rs : set rs = cmd.execute

		response.write "<select id='custcode' name='custcode'>"
		do until rs.eof
		response.write "<option value='" & rs(0)&"' "
		if rs(0) = custcode then response.write " selected "
		response.write ">"
		rs.movenext
		loop
		rs.close
		set rs = nothing
		set cmd = nothing
	end function
%>