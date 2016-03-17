<script language = "javascript">
function test(msg){
	alert(msg);
}
</script>



<%

	function checksession(pUserid, pClass)
		dim objrs, sql
		If Left(pUserid, 1) = "B" Then
			sql = "select class from wb_med_employee where empid = '" & pUserid & "' "
		Else
			sql = "select class from dbo.wb_Account where userid = '" & pUserid & "' "
		End If
		call get_recordset(objrs, sql)

		if objrs.eof then
			checksession = true
		else
			if objrs("class") <> pClass then
				checksession = true
			else
				checksession = false
			end if
		end if
	end function

	'레코드셋 가져오기
	function get_recordset(objrs, sql)
		set objrs = server.CreateObject("adodb.recordset")
		objrs.activeconnection = application("connectionstring")
		objrs.cursorlocation = aduseclient
		objrs.cursortype = adopenforwardonly
		objrs.locktype = adlockreadonly
		objrs.source = sql
		get_recordset = objrs.open
	end function
	' 광고비 집행 을 위한 테스트 서버 접속용
	function get_recordset2(objrs, sql)
		set objrs = server.CreateObject("adodb.recordset")
		objrs.activeconnection = application("connectionstring2")
		objrs.cursorlocation = aduseclient
		objrs.cursortype = adopenforwardonly
		objrs.locktype = adlockreadonly
		objrs.source = sql
		get_recordset2 = objrs.open
	end function

	' 레코드셋으로 데이터 입력, 수정, 삭제하기
	function set_recordset(objrs, sql)
		set objrs = server.CreateObject("adodb.recordset")
		objrs.activeconnection = application("connectionstring")
		objrs.cursorlocation = aduseclient
		objrs.cursortype = adopenstatic
		objrs.locktype = adlockoptimistic
		objrs.source = sql
		set_recordset = objrs.open
	end function

	' 광고주/매체사/외주업체 목록 상자
	sub get_custcode_total(custcode, mode, url)
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "SELECT custcode, custname  FROM dbo.SC_CUST_TEMP where medflag in ('A', 'B') and custcode = highcustcode and attr10 = 1 order by custname"
		rs.open
		response.write"<select name='seltotalcustcode'"
			if not isnull(url) then response.write " onchange='go_page("""&url&""");' "
			if not isnull(mode) then response.write " disabled "
		response.write " >" & vbCrLf
		response.write"<option value=''>광고주를 선택하세요.</option>"
		Do Until rs.eof
			response.write "<option value='"&rs("custcode")&"' "
				if custcode = rs("custcode") then response.write " selected "
			response.write "> " & rs("custname") &"</option>"&vbCRLf
			rs.movenext
		loop
		response.write"</select>"
		rs.close
		set rs = nothing
	end sub


	' 광고주 목록 상자
	sub get_custcode_mst(custcode, mode, url)
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "SELECT highcustcode, custname  FROM dbo.SC_CUST_HDR where medflag='A'  order by custname"
		rs.open
		response.write"<select name='selcustcode'"
			if not isnull(url) then response.write " onchange='go_page("""&url&""");' "
			if not isnull(mode) then response.write " disabled "
		response.write " >" & vbCrLf
		response.write"<option value=''>광고주를 선택하세요.</option>"
		Do Until rs.eof
			response.write "<option value='"&rs("highcustcode")&"' "
				if custcode = rs("highcustcode") then response.write " selected "
			response.write "> " & rs("custname") &"</option>"&vbCRLf
			rs.movenext
		loop
		response.write"</select>"
		rs.close
		set rs = nothing
	end sub

	' 광고주 사업부 목록 상자 -> 광고주 팀 목록 상자
	sub get_custcode_dept(custcode)
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "SELECT custcode, custname  FROM dbo.SC_CUST_DTL where highcustcode like '" & custcode & "%' and medflag='A'  order by custname"
		rs.open
		response.write"<select name='seldeptcode'>" & vbCrLf
		response.write"<option value=''>사업부를 선택하세요.</option>"
		Do Until rs.eof
			response.write "<option value='"&rs("custcode")&"' "
				if custcode = rs("custcode") then response.write " selected "
			response.write "> " & rs("custname") &"</option>"&vbCRLf
			rs.movenext
		loop
		response.write"</select>"
		rs.close
		set rs = nothing
	end sub

	sub get_blank_select(str, wid)
		response.write "<select name='' style='width:"&wid&";'><option value=''>"&str&"</option></select>"
	end sub

	' 광고주별 사업부 목록 상자
	sub get_custcode_custcode2(custcode, custcode2, mode)
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "SELECT custcode, custname  FROM dbo.SC_CUST_DTL where medflag='A' and highcustcode = '" & custcode & "' order by custname "
		rs.open
		response.write"<select name='selcustcode2' "
		if not isnull(mode) then response.write "disabled" end if
		response.write ">"
		response.write"<option value=''>사업부를 선택하세요.</option>"
		Do Until rs.eof
			response.write "<option value='"&rs("custcode")&"' "
				if custcode2 = rs("custcode") then response.write " selected "
			response.write "> " & rs("custname") &"</option>"&vbCRLf
			rs.movenext
		loop
		response.write"</select>"
		rs.close
		set rs = nothing
	end sub



	' 광고주별 사업부 목록 상자(소재 관리용)
	sub get_custcode_custcode2_job(custcode, custcode2, url, mode)
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "SELECT custcode, custname  FROM dbo.SC_CUST_TEMP where medflag='A' and custcode <>  highcustcode and highcustcode = '" & custcode & "'  and attr10 = 1 order by custname"
		rs.open
		response.write"<select name='selcustcode2' "
			if not isnull(mode) then response.write "disabled" end if
			if not isnull(url) then response.write " onchange='go_page("""&url&""");' "
		response.write ">"
		response.write"<option value=''>사업부를 선택하세요.</option>"
		Do Until rs.eof
			response.write "<option value='"&rs("custcode")&"' "
				if custcode2 = rs("custcode") then response.write " selected "
			response.write "> " & rs("custname") &"</option>"&vbCRLf
			rs.movenext
		loop
		response.write"</select>"
		rs.close
		set rs = nothing
	end sub

	' 매체사 목록 상자
	sub get_medium_custcode(custcode, mode)
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "SELECT custcode, custname  FROM dbo.SC_CUST_TEMP where meddiv ='5'  and attr10 = 1 order by custname"
		rs.open
		response.write"<select name='selcustcode' "
			if not isnull(mode) then response.write " disabled "
		response.write ">" & vbCrLf
		response.write"<option value=''>매체사를 선택하세요.</option>"
		Do Until rs.eof
			response.write "<option value='"&rs("custcode")&"'"
				if trim(custcode) = rs("custcode") then response.write " selected "
			response.write "> " & rs("custname") &"</option>"&vbCRLf
			rs.movenext
		loop
		response.write"</select>"
		rs.close
		set rs = nothing

	end sub

	' 매체사 목록 상자
	sub get_custcode_custcode3(custcode, mode)
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "SELECT custcode, custname  FROM dbo.SC_CUST_TEMP where medflag='B' and custcode = highcustcode and meddiv ='5'   and attr10 = 1 order by custname"
		rs.open
		response.write"<select name='selcustcode3' "
			if not isnull(mode) then response.write " disabled "
		response.write ">" & vbCrLf
		response.write"<option value=''>매체사를 선택하세요.</option>"
		Do Until rs.eof
			response.write "<option value='"&rs("custcode")&"'"
				if trim(custcode) = rs("custcode") then response.write " selected "
			response.write "> " & rs("custname") &"</option>"&vbCRLf
			rs.movenext
		loop
		response.write"</select>"
		rs.close
		set rs = nothing

	end sub

	' 매체사 사업부 목록 상자
	sub get_medium_depttcode(custcode, mode)
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "SELECT custcode, custname  FROM dbo.SC_CUST_TEMP where medflag='B' and custcode <> highcustcode  and attr10 = 1 order by custname"
		rs.open
		response.write"<select name='seldeptcode' "
			if mode = "r" then response.write " disabled "
		response.write ">" & vbCrLf
		response.write"<option value=''>매체사를 선택하세요." & custcode & "</option>"
		Do Until rs.eof
			response.write "<option value='"&rs("custcode")&"'"
			if not isnull(custcode) then
				if trim(custcode) = rs("custcode") then response.write " selected "
			end if
			response.write "> " & rs("custname") &"</option>"&vbCRLf
			rs.movenext
		loop
		response.write"</select>"
		rs.close
		set rs = nothing

	end sub

	'광고주별 브랜드 목록 상자
	sub get_jobcust(custcode, seqno, mode, url)
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "select seqno, seqname from dbo.sc_jobcust where clientsubcode = '"&custcode&"' order by seqname"
		rs.open
		response.write"<select name='selseqno' "
			if not isnull(mode) then response.write " disabled "
			if not isnull(url) then response.write " onchange='go_page(""" & url &""")' "
		response.write " style='width:207'>" & vbCrLf
		response.write"<option value=''>브랜드를 선택하세요.</option>"
		Do Until rs.eof
			response.write "<option value='"&rs("seqno")&"'"
				if trim(seqno) = rs("seqno") then response.write " selected "
			response.write "> " & rs("seqname") &"</option>"&vbCRLf
			rs.movenext
		loop
		response.write"</select>"
		rs.close
		set rs = nothing
	end sub


'광고주별 소재관리
	sub get_jobcust_subject(custcode, mode, url, job)
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "select j.jobidx, j.thema from dbo.wb_jobcust j inner join dbo.sc_jobcust j2 on j.seqno = j2.seqno where j2.custcode = '"&custcode&"' order by j.thema"
		rs.open
		response.write"<select name='selsubject' "
			if not isnull(mode) then response.write " disabled "
			if not isnull(url) then response.write " onchange='go_page(""" & url &""")' "
		response.write " style='width:207px;' >" & vbCrLf
		response.write"<option value=''>소재를 선택하세요.</option>"
		Do Until rs.eof
			response.write "<option value='"&rs("jobidx")&"'"
				if job = rs("jobidx") then response.write " selected "
			response.write "> " & rs("thema") &"</option>"&vbCRLf
			rs.movenext
		loop
		response.write"</select>"
		rs.close
		set rs = nothing
	end sub


	' 매체 분류 (중분류) 목록 상자
	sub get_middle_categoty(categoryidx)
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "SELECT categoryidx, categoryname  FROM dbo.wb_category where categorylvl=1"
		rs.open
		response.write"<select name='selcategory'>" & vbCrLf
		response.write"<option value=''>카테고리를 선택하세요.</option>"
		Do Until rs.eof
			response.write "<option value='"&rs("categoryidx")&"' "
			if not isnull(categoryidx) then
				if int(rs("categoryidx")) = int(categoryidx) then response.write " selected "
			end if
			response.write "> " & rs("categoryname") &"</option>"&vbCRLf
			rs.movenext
		loop
		response.write"</select>"
		rs.close
		set rs = nothing

	end sub

	'면 정보 목록 상자
	sub get_side_code(s)
		dim side : side = application("side")
		dim intLoop
		response.write"<select name='selside' "
		response.write ">"
		response.write"<option value=''></option>"
		for intLoop = 0 TO ubound(side)
			response.write"<option value='"&side(intLoop)&"' "
				if side(intLoop) = s then response.write "SELECTED"
			response.write " >" & side(intLoop) & "</option>"
		next
		response.write"</select>"
	end sub

	'재질 목록 상자
	sub get_quality_code(s)
		dim quality : quality = application("quality")
		dim intLoop
		response.write"<select name='selquality'  style='width:132px;'>"
		response.write"<option value=''></option>"
		for intLoop = 0 TO ubound(quality)
			response.write"<option value='"&quality(intLoop)&"' "
				if quality(intLoop) = s then response.write "SELECTED"
			response.write " >" & quality(intLoop) & "</option>"
		next
		response.write"</select>"
	end sub

	'지역 목록 상자
	sub get_region_code(s, mode)
		dim region : region = application("region")
		dim intLoop
		response.write"<select name='selregion'"
			if mode = "r" then response.write " disabled "
		response.write ">" & vbCrLf
		response.write"<option value=''></option>"
		for intLoop = 0 TO ubound(region)
			response.write"<option value='"&region(intLoop)&"' "
				if region(intLoop) = s Then Response.write " SELECTED"
			response.write " >" & region(intLoop) & "</option>" & VbCrLf
		next
		response.write"</select>"
	end sub

	'매체 전체 분류 항목 가져오기
	sub get_medium_catetory(categoryidx)
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "dbo.vw_medium_category"
		rs.open
		rs.find = "mdidx = " & categoryidx
		dim category
		if not isnull(rs("ggroupname")) then
			category = rs("ggroupname")
			if not isnull(rs("mgroupname")) then
				category = category & " > " & rs("mgroupname")
				if not	  isnull(rs("sgroupname")) then
					category = category & " > " & rs("sgroupname")
					if not isnull(rs("dgroupname")) then
						category = category & " > " & rs("dgroupname")
					end if
				end if
			end if
		end if
		response.write category
	end sub

	sub get_year(y)
		dim str, intLoop
		str = "<select name='cyear'>"
		for intLoop = 2005 to year(date)+5
			str = str & "<option value='" & intLoop &"'"
				if cint(y) = intLoop then str = str & " selected"
			str = str & ">" & intLoop & "</option>"
		next
		str = str & "</select>"
		response.write str
	end sub

	sub get_month(m)
		dim str, intLoop
		str = "<select name='cmonth'>"
		for intLoop = 1 to 12
			str = str & "<option value='" & intLoop &"'"
				if cint(m) = intLoop then str = str & " selected"
			str = str & ">" & intLoop & "</option>"
		next
		str = str & "</select>"
		response.write str
	end sub

		sub get_year2(y)
		dim str, intLoop
		str = "<select name='cyear2'>"
		for intLoop = 2005 to year(date)+5
			str = str & "<option value='" & intLoop &"'"
				if cint(y) = intLoop then str = str & " selected"
			str = str & ">" & intLoop & "</option>"
		next
		str = str & "</select>"
		response.write str
	end sub

	sub get_month2(m)
		dim str, intLoop
		str = "<select name='cmonth2'>"
		for intLoop = 1 to 12
			str = str & "<option value='" & intLoop &"'"
				if cint(m) = intLoop then str = str & " selected"
			str = str & ">" & intLoop & "</option>"
		next
		str = str & "</select>"
		response.write str
	end sub

	sub pagesplit(totalrecord, gotopage, pagesize, page, searchstring, custcode, custcode2, cyear, cmonth, midx, mtitle)
		dim blockpage : blockpage = Fix((totalrecord-1)/pagesize)+1
		dim blockno : blockno = Fix((gotopage-1)/pagesize)*pagesize+1

		if blockno > 1 then
			response.write " <a href='" & page &".asp?gotopage="&blockno-pagesize&"&searchstring="&searchstring&"&selcustcode="&custcode&"&selcustcode2="&custcode2&"&cyear="&cyear&"&cmonth="&cmonth&"&midx="&midx&"&mtitle="&mtitle&"'><img src='/images/icon_prev.gif' width='5' height='8' border='0' align='absmiddle' vspace='5'></a> "
		end if

		dim intLoop
		for intLoop = blockno to ((blockno-1)+pagesize)
			if intLoop <= int(blockpage) then
				if intLoop = int(gotopage) then
					response.write " "& intLoop & " "
				else
					response.write " <a href='" & page &".asp?gotopage="&intLoop&"&searchstring="&searchstring&"&selcustcode="&custcode&"&selcustcode2="&custcode2&"&cyear="&cyear&"&cmonth="&cmonth&"&midx="&midx&"&mtitle="&mtitle&"' class='pagesplit' >"&intLoop&"</a> "
				end if
			end if
		next

		if intLoop < blockpage then
			response.write " <a href='" & page &".asp?gotopage="&blockno+pagesize&"&searchstring="&searchstring&"&selcustcode="&custcode&"&selcustcode2="&custcode2&"&cyear="&cyear&"&cmonth="&cmonth&"&midx="&midx&"&mtitle="&mtitle&"'><img src='/images/icon_next.gif' width='5' height='8' border='0' align='absmiddle' vspace='5'></a> "
		end if
	end sub


	sub boardpagesplit2(totalrecord, gotopage, pagesize, searchstring,midx)
		dim blockpage : blockpage = Int((totalrecord-1)/pagesize)+1
		dim blockno : blockno = Int((gotopage-1)/pagesize)*pagesize+1


		if blockno <> 1 then
			response.write " <a href='list.asp?gotopage="&blockno-pagesize&"&searchstring="&searchstring&"&midx="&midx&"'><img src='/images/icon_prev.gif' width='5' height='8' border='0' align='absmiddle' vspace='5'></a> "
		end if

		dim intLoop
		for intLoop = blockno to ((blockno-1)+pagesize)
			if intLoop <= int(blockpage) then
				if intLoop = int(gotopage) then
					response.write " "& intLoop & " "
				else
					response.write " <a href='list.asp?gotopage="&intLoop&"&searchstring="&searchstring&"&midx="&midx&"' class='pagesplit'>"&intLoop&"</a> "
				end if
			end if
		next

		if intLoop < blockpage then
			response.write "<a href='list.asp?gotopage="&blockno-pagesize&"&searchstring="&searchstring&"&midx="&midx&"'><img src='/images/icon_next.gif' width='17' height='13' border='0' align='absmiddle' vspace='5'></a> "
		end if
	end Sub
	
	sub boardpagesplit(totalrecord, gotopage, pagesize, searchstring,midx, custcode, title)
		dim blockpage : blockpage = Int((totalrecord-1)/pagesize)+1
		dim blockno : blockno = Int((gotopage-1)/pagesize)*pagesize+1

		If custcode = "" Then custcode = " "
		If title = "" Then title = " "

		if blockno <> 1 then
			response.write " <a href=""#""  onclick='get_pageSRC(""list.asp"", "&midx&", """&custcode&""", """",  """&title&""", "&blockno-pagesize&");'> "& intLoop & "</a> "
		end if

		dim intLoop
		for intLoop = blockno to ((blockno-1)+pagesize)
			if intLoop <= int(blockpage) then
				if intLoop = int(gotopage) then
					response.write " "& intLoop & " "
				Else
					response.write "<a href=""#"" class='pagesplit' onclick='get_pageSRC(""list.asp"", "&midx&", """&custcode&""", """",  """&title&""", "&intLoop&");'> "& intLoop & "</a>"
				end if
			end if
		next

		if intLoop < blockpage then
			response.write "<a href=""#""  onclick='get_pageSRC(""list.asp"", "&midx&", """&custcode&""", """",  """&title&""", "&blockno-pagesize&");'> "& intLoop & "</a> "
		end if


	end sub

	sub getSendMail(mail, tomail, title, content)
		dim objMail
		if isnull(mail) then mail = ""
		Set objMail = Server.CreateObject("CDO.Message")
		objMail.From = mail
		objMail.To = tomail
		objMail.Subject = title
		objMail.TextBody = content
		objMail.Send
		response.write "<script> alert('메일발송이 완료되었습니다.'); </script>"
		Set objMail = Nothing
	end sub


	sub get_category_grand(idx, mode, func)
		dim rs, sql, str
		sql = "select categoryidx, categoryname from dbo.wb_category where categorylvl  is null order by categoryidx "
		call get_recordset(rs, sql)

		str = "<select name='selgcategory' style='width:320px;' "
		if not isnull(func) then str = str & " onclick = '" & func & "' "
		if not isnull(mode) then str = str & " disabled "
		str = str & " > <option value = ''> 대분류를 선택하세요. </option>"
		do until rs.eof
			str = str & "<option value ='" & rs("categoryidx") & "'  >" & rs("categoryname") & " </option>"
		rs.movenext
		loop
		str = str & "</select>"
		rs.close
		set rs = nothing
		response.write str
	end sub

	function get_cyear(y)
		dim str, intLoop
		str = "<select name='cyear'>"
		for intLoop = 2005 to year(date)+5
			str = str & "<option value='" & intLoop &"'"
				if cint(y) = intLoop then str = str & " selected"
			str = str & ">" & intLoop & "</option>"
		next
		str = str & "</select>"
		get_cyear = str
	end function


	function get_cyear2(y)
		dim str, intLoop
		str = "<select name='cyear2'>"
		for intLoop = 2005 to year(date)+5
			str = str & "<option value='" & intLoop &"'"
				if cint(y) = intLoop then str = str & " selected"
			str = str & ">" & intLoop & "</option>"
		next
		str = str & "</select>"
		get_cyear2 = str
	end function


	function get_cmonth2(m)
		dim str, intLoop
		str = "<select name='cmonth2'>"
		for intLoop = 1 to 12
			str = str & "<option value='" & intLoop &"'"
				if cint(m) = intLoop then str = str & " selected"
			str = str & ">" & intLoop & "</option>"
		next
		str = str & "</select>"
		get_cmonth2 = str
	end function

	function get_cmonth(m)
		dim str, intLoop
		str = "<select name='cmonth'>"
		for intLoop = 1 to 12
			str = str & "<option value='" & intLoop &"'"
				if cint(m) = intLoop then str = str & " selected"
			str = str & ">" & intLoop & "</option>"
		next
		str = str & "</select>"
		get_cmonth = str
	end function

	function get_custcode2(custcode, custcode2)
		dim rs, sql, str, intLoop
		sql = "select custcode, custname from dbo.sc_cust_temp where highcustcode = '" & custcode & "' "
		call get_recordset(rs, sql)

		str = "<select name='custcode2'>"
		do until rs.eof
			str = str & "<option value='" & rs("custcode") & "' "
				if custcode2 = rs("custcode") then str = str & "selected"
			str = str &">" & rs("custname") & "</option>"
		rs.movenext
		loop
		str = str & "</select>"
		get_custcode2 = str
	end Function


	sub get_use_custcode(custcode, mode, url)
		dim rs : set rs = server.createobject("adodb.recordset")
		rs.activeconnection = application("connectionstring")
		rs.cursorlocation = aduseclient
		rs.cursortype = adopenforwardonly
		rs.locktype = adlockreadonly
		rs.source = "select distinct h.oldcustcode, h.custname from dbo.sc_cust_dtl d inner join dbo.wb_contact_mst m on d.oldcustcode = m.custcode inner join dbo.sc_cust_hdr h on d.highcustcode = h.highcustcode where d.use_flag = '1' order by h.custname desc"
		rs.open
		response.write"<select name='selcustcode'"
			if not isnull(url) then response.write " onchange='go_page("""&url&""");' "
			if not isnull(mode) then response.write " disabled "
		response.write " >" & vbCrLf
		response.write"<option value=''>광고주를 선택하세요.</option>"
		Do Until rs.eof
			response.write "<option value='"&rs("oldcustcode")&"' "
				if custcode = rs("oldcustcode") then response.write " selected "
			response.write "> " & rs("custname") &"</option>"&vbCRLf
			rs.movenext
		loop
		response.write"</select>"
		rs.close
		set rs = nothing
	end sub

	sub set_initclipingLevel(userid)
		Dim sql : sql = "update wb_account set clipinglevel = 0 where userid =?"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		cmd.parameters.append cmd.createparameter("userid", advarChar, adParamInput, 12)
		cmd.parameters("userid").value = userid
		cmd.execute ,, adExecuteNoRecords
		set cmd = nothing
	end sub

	sub set_initclipingLevel2(userid)
		Dim sql : sql = "update wb_med_employee set clipinglevel = 0 where empid =?"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		cmd.parameters.append cmd.createparameter("empid", advarChar, adParamInput, 12)
		cmd.parameters("empid").value = userid
		cmd.execute ,, adExecuteNoRecords
		set cmd = nothing
	end sub

	sub set_clipingLevel(userid, lvl)
		Dim sql : sql = "update wb_account set clipinglevel = ? where userid =?"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		cmd.parameters.append cmd.createparameter("clipinglevel", adUnsignedTinyint, adParamInput)
		cmd.parameters.append cmd.createparameter("userid", advarChar, adParamInput, 12)
		cmd.parameters("clipinglevel").value = lvl
		cmd.parameters("userid").value = userid
		cmd.execute ,, adExecuteNoRecords
		set cmd = nothing
	end sub

	sub set_clipingLevel2(userid, lvl)
		Dim sql : sql = "update wb_med_employee set clipinglevel = ? where empid =?"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		cmd.parameters.append cmd.createparameter("clipinglevel", adUnsignedTinyint, adParamInput)
		cmd.parameters.append cmd.createparameter("userid", advarChar, adParamInput, 12)
		cmd.parameters("clipinglevel").value = lvl
		cmd.parameters("userid").value = userid
		cmd.execute ,, adExecuteNoRecords
		set cmd = nothing
	end sub


	sub set_isuse(userid)
		Dim sql : sql = "update wb_account set clipinglevel = 0, isuse='N' where userid =?"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		cmd.parameters.append cmd.createparameter("userid", advarChar, adParamInput, 12)
		cmd.parameters("userid").value = userid
		cmd.execute ,, adExecuteNoRecords
		set cmd = nothing
	end sub

	sub set_isuse2(userid)
		Dim sql : sql = "update wb_med_employee set clipinglevel = 0, isuse='N' where empid =?"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		cmd.parameters.append cmd.createparameter("userid", advarChar, adParamInput, 12)
		cmd.parameters("userid").value = userid
		cmd.execute ,, adExecuteNoRecords
		set cmd = nothing
	end sub

	function getyearmon(cyear, cmonth)
		if isnull(cyear) or cyear="" then 	cyear = Year(date)
		if isnull(cmonth) or cmonth = "" then cmonth = Month(date)
		getyearmon = cyear&cmonth
	end function

	function search_cyear2cmonth2(cyear, cmonth, cyear2, cmonth2, custcode)
		if cyear =  "" then cyear = Cstr(Year(date))
		if cmonth = "" then cmonth = Cstr(Month(Date))
		if cyear2 =  "" then cyear2 = Cstr(Year(date))
		if cmonth2 = "" then cmonth2 = Cstr(Month(Date))
		dim intLoop
		dim strIntLoop

		response.write "<select id='cyear' name='cyear' style='margin-left:20px;'>"
		for intLOop = 2008 to YEar(date)
		response.write "<option value='"&intLOop&"'"
		if cyear = CSTR(intLoop) then response.write " selected"
		response.write ">" & intLoop & "</option>" & vbcrlf
		next
		response.write "</select>"

		response.write "<select id='cmonth' name='cmonth' style='margin-left:3px;'>"
		for intLOop = 1 to 12
		if Len(intLoop) = 1 then strIntLoop = "0" & intLOop else strIntLoop = intLoop
		response.write "<option value='"&strIntLoop&"'"
		if cmonth = CSTR(strIntLoop) then response.write " selected"
		response.write ">" & strintLoop & "</option>" & vbcrlf
		next
		response.write "</select>"


		response.write " - <select id='cyear2' name='cyear2' style='margin-left:5px;'>"
		for intLOop = 2008 to YEar(date)
		response.write "<option value='"&intLOop&"'"
		if cyear2 = CSTR(intLoop) then response.write " selected"
		response.write ">" & intLoop & "</option>" & vbcrlf
		next
		response.write "</select>"

		response.write "<select id='cmonth2' name='cmonth2' style='margin-left:3px;'>"
		for intLOop = 1 to 12
		if Len(intLoop) = 1 then strIntLoop = "0" & intloop else strIntLoop = intLOop
		response.write "<option value='"&strIntLoop&"'"
		if cmonth2 = CSTR(strIntLoop) then response.write " selected"
		response.write ">" & strintLoop & "</option>" & vbcrlf
		next
		response.write "</select>"

		call search_custcode(custcode)
	end function


	function search_SELECT_cyear2cmonth2(cyear, cmonth, cyear2,  cmonth2)
		if cyear =  "" then cyear = Cstr(Year(date))
		if cmonth = "" then cmonth = Cstr(Month(Date))
		if cyear2 =  "" then cyear2 = Cstr(Year(date))
		if cmonth2 = "" then cmonth2 = Cstr(Month(Date))
		dim intLoop
		dim strIntLoop

		response.write "<select id='cyear' name='cyear' >"
		for intLOop = 2008 to YEar(date)
		response.write "<option value='"&intLOop&"'"
		if cyear = CSTR(intLoop) then response.write " selected"
		response.write ">" & intLoop & "</option>" & vbcrlf
		next
		response.write "</select>"


		response.write "<select id='cmonth' name='cmonth' style='margin-left:3px;' >"
		for intLOop = 1 to 12
		if Len(intLoop) = 1 then strIntLoop = "0" & intLOop else strIntLoop = intLOop
		response.write "<option value='"&strIntLoop&"'"
		if cmonth = CSTR(strIntLoop) then response.write " selected"
		response.write ">" & strintLoop & "</option>" & vbcrlf
		next
		response.write "</select>"

		response.write " - <select id='cyear2' name='cyear2' style='margin-left:5px;'>"
		for intLOop = 2008 to YEar(date)
		response.write "<option value='"&intLOop&"'"
		if cyear2 = CSTR(intLoop) then response.write " selected"
		response.write ">" & intLoop & "</option>" & vbcrlf
		next
		response.write "</select>"

		response.write " - <select id='cmonth2' name='cmonth2' >"
		for intLOop = 1 to 12
		if Len(intLoop) = 1 then strIntLoop = "0" & intloop else strIntLoop = intLOop
		response.write "<option value='"&strIntLoop&"'"
		if cmonth2 = CSTR(strIntLoop) then response.write " selected"
		response.write ">" & strintLoop & "</option>" & vbcrlf
		next
		response.write "</select>"

		'call search_hidden_custcode(custcode)
	end function

	function search_cyearcmonth2(cyear, cmonth, cmonth2, custcode)
		if cyear =  "" then cyear = Cstr(Year(date))
		if cmonth = "" then cmonth = Cstr(Month(Date))
		if cmonth2 = "" then cmonth2 = Cstr(Month(Date))
		dim intLoop
		dim strIntLoop

		response.write "<select id='cyear' name='cyear' >"
		for intLOop = 2008 to YEar(date)
		response.write "<option value='"&intLOop&"'"
		if cyear = CSTR(intLoop) then response.write " selected"
		response.write ">" & intLoop & "</option>" & vbcrlf
		next
		response.write "</select>"


		response.write "<select id='cmonth' name='cmonth' style='margin-left:3px;' >"
		for intLOop = 1 to 12
		if Len(intLoop) = 1 then strIntLoop = "0" & intLOop else strIntLoop = intLOop
		response.write "<option value='"&strIntLoop&"'"
		if cmonth = CSTR(strIntLoop) then response.write " selected"
		response.write ">" & strintLoop & "</option>" & vbcrlf
		next
		response.write "</select>"

		response.write " - <select id='cmonth2' name='cmonth2' >"
		for intLOop = 1 to 12
		if Len(intLoop) = 1 then strIntLoop = "0" & intloop else strIntLoop = intLOop
		response.write "<option value='"&strIntLoop&"'"
		if cmonth2 = CSTR(strIntLoop) then response.write " selected"
		response.write ">" & strintLoop & "</option>" & vbcrlf
		next
		response.write "</select>"

		call search_custcode(custcode)
	end function


	function search_SELECT_cyearcmonth2(cyear, cmonth, cmonth2)
		if cyear =  "" then cyear = Cstr(Year(date))
		if cmonth = "" then cmonth = Cstr(Month(Date))
		if cmonth2 = "" then cmonth2 = Cstr(Month(Date))
		dim intLoop
		dim strIntLoop

		response.write "<select id='cyear' name='cyear' >"
		for intLOop = 2008 to YEar(date)
		response.write "<option value='"&intLOop&"'"
		if cyear = CSTR(intLoop) then response.write " selected"
		response.write ">" & intLoop & "</option>" & vbcrlf
		next
		response.write "</select>"


		response.write "<select id='cmonth' name='cmonth' style='margin-left:3px;' >"
		for intLOop = 1 to 12
		if Len(intLoop) = 1 then strIntLoop = "0" & intLOop else strIntLoop = intLOop
		response.write "<option value='"&strIntLoop&"'"
		if cmonth = CSTR(strIntLoop) then response.write " selected"
		response.write ">" & strintLoop & "</option>" & vbcrlf
		next
		response.write "</select>"

		response.write " - <select id='cmonth2' name='cmonth2' >"
		for intLOop = 1 to 12
		if Len(intLoop) = 1 then strIntLoop = "0" & intloop else strIntLoop = intLOop
		response.write "<option value='"&strIntLoop&"'"
		if cmonth2 = CSTR(strIntLoop) then response.write " selected"
		response.write ">" & strintLoop & "</option>" & vbcrlf
		next
		response.write "</select>"

	end function

	function search_cyearcmonth(cyear, cmonth, custcode)
		if cyear =  "" then cyear = Cstr(Year(date))
		if cmonth = "" then cmonth = Cstr(Month(Date))
		dim intLoop
		dim strIntLoop

		response.write "<select id='cyear' name='cyear' >"
		for intLOop = 2008 to YEar(date)
		response.write "<option value='"&intLOop&"'"
		if cyear = CSTR(intLoop) then response.write " selected"
		response.write ">" & intLoop & "</option>" & vbcrlf
		next
		response.write "</select>"


		response.write "<select id='cmonth' name='cmonth' style='margin-left:3px;' >"
		for intLOop = 1 to 12
		if Len(intLoop) = 1 then strIntLoop = "0" & intLOop else strIntLoop = intLOop
		response.write "<option value='"&strIntLoop&"'"
		if cmonth = CSTR(strIntLoop) then response.write " selected"
		response.write ">" & strintLoop & "</option>" & vbcrlf
		next
		response.write "</select>"

		call search_custcode(custcode)
	end function

	function search_SELECT_cyearcmonth(cyear, cmonth )
		if cyear =  "" then cyear = Cstr(Year(date))
		if cmonth = "" then cmonth = Cstr(Month(Date))
		dim intLoop
		dim strIntLoop

		response.write "<select id='cyear' name='cyear' >"
		for intLOop = 2008 to YEar(date)
		response.write "<option value='"&intLOop&"'"
		if cyear = CSTR(intLoop) then response.write " selected"
		response.write ">" & intLoop & "</option>" & vbcrlf
		next
		response.write "</select>"


		response.write "<select id='cmonth' name='cmonth' style='margin-left:3px;' >"
		for intLOop = 1 to 12
		if Len(intLoop) = 1 then strIntLoop = "0" & intLOop else strIntLoop = intLOop
		response.write "<option value='"&strIntLoop&"'"
		if cmonth = CSTR(strIntLoop) then response.write " selected"
		response.write ">" & strintLoop & "</option>" & vbcrlf
		next
		response.write "</select>"


	end function

	function search_cyear(cyear, custcode)
		if cyear =  "" then cyear = Cstr(Year(date))
		dim intLoop

		response.write "<select id='cyear' name='cyear' >"
		for intLOop = 2008 to YEar(date)
		response.write "<option value='"&intLOop&"'"
		if cyear = CSTR(intLoop) then response.write " selected"
		response.write ">" & intLoop & "</option>" & vbcrlf
		next
		response.write "</select>"

		call search_custcode(custcode)
	end function

	function search_SELECT_cyear(cyear)
		if cyear =  "" then cyear = Cstr(Year(date))
		dim intLoop

		response.write "<select id='cyear' name='cyear' >"
		for intLOop = 2008 to YEar(date)
		response.write "<option value='"&intLOop&"'"
		if cyear = CSTR(intLoop) then response.write " selected"
		response.write ">" & intLoop & "</option>" & vbcrlf
		next
		response.write "</select>"

	end function

	function search_custcode(Custcode)
'		dim sql: sql = "select highcustcode, custname from sc_Cust_hdr where medflag='a' order by custname" 기존:20100426바꿈
		dim sql: sql = "select highcustcode, custname from MD_CLIENTCODE_LIST_V order by custname"
		dim cmd : set cmd = server.createobject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandTYpe = adcmdtext
		cmd.commandtext = sql
		dim rs : set rs = cmd.execute

		response.write "<select id='custcode2' name='custcode2' style='width:200px;margin-left:20px;'>"
		response.write "<option value='' "
		if custcode = "" then response.write " selected"
		response.write "> -- Grand Total -- </option>"
		do until rs.eof
		response.write "<option value='" & rs(0)&"' "
		if rs(0) = custcode then response.write " selected "
		response.write ">" & rs(1) & "</option>" & VbCrLF
		rs.movenext
		loop
		response.write "</select>"
		rs.close
		set rs = nothing
		set cmd = nothing
	end function


	function search_hidden_custcode(Custcode)
		dim sql: sql = "select highcustcode, custname from sc_Cust_hdr where medflag='a' order by custname"
		dim cmd : set cmd = server.createobject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandTYpe = adcmdtext
		cmd.commandtext = sql
		dim rs : set rs = cmd.execute

		response.write "<select style='VISIBILITY: hidden'  id='custcode2' name='custcode2' style='width:200px;margin-left:20px;'>"
		response.write "<option value='' "
		if custcode = "" then response.write " selected"
		response.write "> -- Grand Total -- </option>"
		do until rs.eof
		response.write "<option value='" & rs(0)&"' "
		if rs(0) = custcode then response.write " selected "
		response.write ">" & rs(1) & "</option>" & VbCrLF
		rs.movenext
		loop
		response.write "</select>"
		rs.close
		set rs = nothing
		set cmd = nothing
	end Function
	



%>