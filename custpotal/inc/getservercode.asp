<%
	Sub md_part_deps1()
		Dim rs : Set rs = server.CreateObject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.source = "SELECT DEPSIDX, DEPSNAME FROM dbo.WEB_SP_PARTCODE WHERE HIGHDEPSIDX IS NULL"
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.open 

		Dim depsidx, depsname
		If Not rs.eof Then
			Set depsidx = rs(0)
			Set depsname = rs(1)
			response.write "<select name='deps1' onchange='requestMiddlecode(this.selectedIndex)'>" & vbcrlf
			response.write "<option value=''>매체 대분류를 선택하세요</option>" & vbcrlf
			Do Until rs.eof 
				response.write "<option value='" & depsidx.value & "'>" & depsname.value & "</option>" & vbcrlf
			rs.movenext
			Loop
			response.write "</select>"
		End If
		rs.close
		Set rs = nothing
	End sub

	
	sub getSideCode(s)
		dim side : side = application("side")
		dim intLoop
		response.write"<select name='selside'>" & vbCrLf
		response.write"<option value=''></option>"
		for intLoop = 0 TO ubound(side)
			response.write"<option value='"&side(intLoop)&"' "
				if side(intLoop) = s then response.write "SELECTED"
			response.write " >" & side(intLoop) & "</option>" & VbCrLf
		next
		response.write"</select>"
	end sub

	sub getQualityCode(s)
		dim quality : quality = application("quality")
		dim intLoop
		response.write"<select name='selquality'>" & vbCrLf
		response.write"<option value=''></option>"
		for intLoop = 0 TO ubound(quality)
			response.write"<option value='"&quality(intLoop)&"' "
				if quality(intLoop) = s then response.write "SELECTED"
			response.write " >" & quality(intLoop) & "</option>" & VbCrLf
		next
		response.write"</select>"
	end sub

	sub getRegion(s)
		dim region : region = application("region")
		dim intLoop
		response.write"<select name='selregion'>" & vbCrLf
		response.write"<option value=''></option>"
		for intLoop = 0 TO ubound(region)
			response.write"<option value='"&region(intLoop)&"' "
				if region(intLoop) = s Then Response.write " SELECTED"
			response.write " >" & region(intLoop) & "</option>" & VbCrLf
		next
		response.write"</select>"
	end sub

	sub getCategory(idx)
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "SELECT C.CATEGORYNAME, C1.CATEGORYNAME,  C2.CATEGORYNAME, C3.CATEGORYNAME FROM DBO.WEB_CATEGORY C " & _
						" INNER JOIN DBO.WEB_CATEGORY C1 ON C.CATEGORYIDX = C1.HIGHCATEGORYIDX " & _
						" INNER JOIN DBO.WEB_CATEGORY C2 ON C1.CATEGORYIDX = C2.HIGHCATEGORYIDX " & _
						" LEFT OUTER JOIN DBO.WEB_CATEGORY C3 ON C2.CATEGORYIDX = C3.HIGHCATEGORYIDX " & _
						" WHERE (C2.CATEGORYIDX = " & idx & " OR C3.CATEGORYIDX =  " & idx & " ) AND C2.CATEGORYLVL = 2"
		rs.open
		if not rs.eof then
			response.write rs(0) & " > " & rs(1) &  " > " & rs(2)
			if not isNull(rs(3)) then response.write " > " & rs(3)
		end if
		rs.close
		set rs = nothing
	end sub

	sub getEditCategory(idx)
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "SELECT C.CATEGORYNAME, C1.CATEGORYNAME,  C2.CATEGORYNAME, C3.CATEGORYNAME FROM DBO.WEB_CATEGORY C " & _
						" INNER JOIN DBO.WEB_CATEGORY C1 ON C.CATEGORYIDX = C1.HIGHCATEGORYIDX " & _
						" INNER JOIN DBO.WEB_CATEGORY C2 ON C1.CATEGORYIDX = C2.HIGHCATEGORYIDX " & _
						" LEFT OUTER JOIN DBO.WEB_CATEGORY C3 ON C2.CATEGORYIDX = C3.HIGHCATEGORYIDX " & _
						" WHERE (C2.CATEGORYIDX = " & idx & " OR C3.CATEGORYIDX =  " & idx & " ) AND C2.CATEGORYLVL = 2"
		rs.open
		if not rs.eof then
			response.write rs(0) & " > " & rs(1) & " > "
			if not isNull(rs(3)) then response.write " > " & rs(2)& " > "
		end if
		rs.close
		set rs = nothing
	end Sub


	Sub getMiddleCategory()
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "SELECT CATEGORYIDX, CATEGORYNAME FROM dbo.WEB_CATEGORY WHERE CATEGORYLVL = 1"
		rs.open

		response.write"<select name='selcategory'>" & vbCrLf
		response.write"<option value=''></option>"
		Do Until rs.eof
			response.write "<option value='"&rs("categoryidx")&"'> " & rs("categoryname") &"</option>"&vbCRLf
			rs.movenext
		loop
		response.write"</select>"
		rs.close
		set rs = nothing
	End sub

	sub getDeptcode()
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "SELECT distinct custcode, custname  FROM dbo.SC_CUST_TEMP "
		rs.open
		response.write"<select name='selcustcode'>" & vbCrLf
		response.write"<option value=''>광고주를 선택하세요.</option>"
		Do Until rs.eof
			response.write "<option value='"&rs("custcode")&"'> " & rs("custname") &"</option>"&vbCRLf
			rs.movenext
		loop
		response.write"</select>"
		rs.close
		set rs = nothing
	end sub

	sub getCustCode()
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "SELECT custcode, custname  FROM dbo.SC_CUST_TEMP where custcode = highcustcode "
		rs.open
		response.write"<select name='selcustcode'>" & vbCrLf
		response.write"<option value=''>광고주를 선택하세요.</option>"
		Do Until rs.eof
			response.write "<option value='"&rs("custcode")&"'> " & rs("custname") &"</option>"&vbCRLf
			rs.movenext
		loop
		response.write"</select>"
		rs.close
		set rs = nothing

	end sub

	sub getParamCust(code)
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "SELECT distinct custcode, custname  FROM dbo.SC_CUST_TEMP WHERE MEDFLAG = 'A' AND CUSTCODE = HIGHCUSTCODE "
		rs.open
		response.write"<select name='selcustcode' onchange='changeForDept()'>" & vbCrLf
		response.write"<option value=''>광고주를 선택하세요.</option>"
		Do Until rs.eof
			response.write "<option value='"&rs("custcode")&"' "
				if rs("custcode") = code then response.write " selected "
			response.write " > " & rs("custname") &"</option>"&vbCRLf
			rs.movenext
		loop
		response.write"</select>"
		rs.close
		set rs = nothing

	end sub

	sub getParamDept(code, deptcode)
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "SELECT distinct custcode, custname  FROM dbo.SC_CUST_TEMP WHERE highcustcode = '" & code &"' "
		rs.open
		response.write"<select name='seldeptcode' onchange='checkForSearch();'>" & vbCrLf
		response.write"<option value=''>광고주를 선택하세요.</option>"
		Do Until rs.eof
			response.write "<option value='"&rs("custcode")&"' "
				if rs("custcode") = deptcode then response.write " selected "
			response.write " > " & rs("custname") &"</option>"&vbCRLf
			rs.movenext
		loop
		response.write"</select>"
		rs.close
		set rs = nothing

	end sub

	sub getDateTime(yname, mname, dname, dat)
		dim y : y = year(dat)
		dim m : m = month(dat)
		dim d : d = day(dat)

		dim i
		response.write "<select name='"&yname&"' >"
		for i = 2000 to y
			response.write "<option value='" & i & "' "
			if i = y then response.write " selected "
			response.write " > " & i & " </option>"
		response.write "</select>"
		next
		response.write "<select name='"&mname&"' >"
		for i = 1 to 12
			response.write "<option value='" & i & "' "
			if i = m then response.write " selected "
			response.write " > " & i & " </option>"
		response.write "</select>"
		next
		response.write "<select name='"&dname&"' >"
		for i = 1 to 31
			response.write "<option value='" & i & "' "
			if i = d then response.write " selected "
			response.write " > " & i & " </option>"
		response.write "</select>"
		next


	end sub
%>
