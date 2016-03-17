<%
	Sub validRequest(args)
		Dim item
		For Each item In args
			response.write item &  " : " & args(item) & "<br>"
		Next
		'Response.End
	End Sub

	Sub getquerystringparameter
		Dim item
		For Each item In request.querystring
			response.write item &  " : " & request.querystring(item) & "<br>"
		Next
	End Sub

	Sub qetformparameter
		Dim item
		For Each item In request.form
			response.write item &  " : " & request.form(item) & "<br>"
		Next
	End Sub

	'년도
	Sub getyear(cyear)
		If IsNull(cyear) Then cyear = Year(date)
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = "select MIN(startdate) as startdate from wb_contact_mst"
		cmd.commandType = adcmdtext
		Dim rs : Set rs = cmd.execute
		Dim startdate : startdate = Year(rs(0))
		If rs.eof Then startdate = 2008

		Dim intLoop
		response.write "<select name='cyear' id='cyear'>"
		For intLoop = startdate To Year(date) + 5
			response.write "<option value='" & intLoop &"' "
				If CInt(cyear) = intLoop Then response.write "selected"
			response.write ">" & intLoop & "</option>"
		Next
		response.write "</select>"
	End Sub

	' 해당월
	Sub getmonth(cmonth)
		If IsNull(cmonth) Then cmonth = Month(date)
		If Len(cmonth) = 1 Then cmonth = "0"&cmonth
		Dim intLoop
		response.write "<select name='cmonth' id='cmonth'>"
		For intLoop = 1 To 12
			If Len(intLoop) = 1 Then intLoop = "0"&intLoop
			response.write "<option value='" & intLoop &"' "
				If CStr(cmonth) = CStr(intLoop) Then response.write "selected"
			response.write ">" & intLoop & "</option>"
		Next
		response.write "</select>"
	End Sub

	Sub getyear2(cyear2)
		If IsNull(cyear) Then cyear = Year(date)
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = "select MIN(startdate) as startdate from wb_contact_mst"
		cmd.commandType = adcmdtext
		Dim rs : Set rs = cmd.execute
		Dim startdate : startdate = Year(rs(0))
		If rs.eof Then startdate = 2008

		Dim intLoop
		response.write "<select name='cyear2' id='cyear2'>"
		For intLoop = startdate To Year(date) + 5
			response.write "<option value='" & intLoop &"' "
				If CInt(cyear2) = intLoop Then response.write "selected"
			response.write ">" & intLoop & "</option>"
		Next
		response.write "</select>"
	End Sub

	' 해당월2
	Sub getmonth2(cmonth2)
		If IsNull(cmonth) Then cmonth = Month(date)
		If Len(cmonth2) = 1 Then cmonth2 = "0"&cmonth2
		Dim intLoop
		response.write "<select name='cmonth2' id='cmonth2'>"
		For intLoop = 1 To 12
			If Len(intLoop) = 1 Then intLoop = "0"&intLoop
			response.write "<option value='" & intLoop &"' "
				If CStr(cmonth2) = CStr(intLoop) Then response.write "selected"
			response.write ">" & intLoop & "</option>"
		Next
		response.write "</select>"
	End Sub

	Sub getyear3(cyear)
		If IsNull(cyear) Then cyear = Year(date)
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = "select MIN(startdate) as startdate from wb_contact_mst"
		cmd.commandType = adcmdtext
		Dim rs : Set rs = cmd.execute
		Dim startdate : startdate = Year(rs(0))
		If rs.eof Then startdate = 2008

		Dim intLoop
		Response.write "<select id='cyear' name='cyear'>"
		response.write "<option value='' selected> --ALL-- </option>"
		For intLoop = startdate To Year(date) + 5
			response.write "<option value='" & intLoop &"' "
				'If CInt(cyear) = intLoop Then response.write "selected"
			response.write ">" & intLoop & "</option>"
		Next
		response.write "</select>"
	End Sub

	' 해당월2
	Sub getmonth3(cmonth)
		If IsNull(cmonth) Then cmonth = Month(date)
		If Len(cmonth) = 1 Then cmonth = "0"&cmonth
		Dim intLoop
		Response.write "<select id='cmonth' name='cmonth'>"
		response.write "<option value='' selected> --ALL-- </option>"
		For intLoop = 1 To 12
			If Len(intLoop) = 1 Then intLoop = "0"&intLoop
			response.write "<option value='" & intLoop &"' "
				'If CStr(cmonth) = CStr(intLoop) Then response.write "selected"
			response.write ">" & intLoop & "</option>"
		Next
		response.write "</select>"
	End Sub

	' 면 정보 한글화
	Function  getside(side)
		Select Case Trim(side)
			Case "F"
				getside = "정면"
			Case "B"
				getside = "후면"
			Case "L"
				getside = "우측"
			Case "R"
				getside = "좌측"
		End Select
	End Function

' 광고 제목 가져오기
	Function gettitle(contidx)
		Dim sql : sql = "select title from wb_contact_mst where contidx = " & contidx
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		If rs.eof Then gettitle = "" Else gettitle = rs(0)
		Set cmd = Nothing
	End Function

'광고주명 가져오기
	Function getcustname(custcode)
		Dim sql : sql = "select h.custname from sc_cust_hdr h inner join sc_cust_dtl d on h.highcustcode = d.highcustcode where d.custcode = '" & custcode & "' "
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		If rs.eof Then getcustname = "" Else getcustname = rs(0)
		Set cmd = Nothing
	End Function

' 사업부서 이름 가져오기
	Function getdeptname(custcode)
		Dim sql : sql = "select h.custname from sc_cust_dtl h inner join sc_cust_dtl d on d.custcode = h.clientsubcode where d.custcode = '" & custcode & "' "
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		If rs.eof Then getdeptname = "" Else getdeptname = rs(0)
		Set cmd = Nothing
	End Function


	' 운영팀 이름 가져오기
	Function getteamname(custcode)
		Dim sql : sql = "select custname from sc_cust_dtl  where custcode = '" & custcode & "' "
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		If rs.eof Then getteamname = "" Else getteamname = rs(0)
		Set cmd = Nothing
	End Function

' 매체사 이름 가져오기
	Function getmedname(medcode)
		Dim sql : sql = "select custname from sc_cust_hdr where highcustcode = '" & medcode & "' "
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		If rs.eof Then getmedname = "" Else getmedname = rs(0)
		Set cmd = Nothing
	End Function

' 매체사 담당자 이름 가져오기
	Function getempname(empid)
		Dim sql : sql = "select empname from wb_med_employee where empid = '" & empid & "' "
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		If rs.eof Then getempname = "" Else getempname = rs(0)
		Set cmd = Nothing
	End Function

' 계약별 년, 월 별 광고 수량 가져오기
	Function getmonthlyqty(contidx, cyear, cmonth)
		Dim sql : sql = "select sum(qty) from wb_contact_mst s inner join wb_contact_md m on s.contidx = m.contidx and s.contidx = "&contidx&" inner join wb_contact_exe e on m.mdidx = e.mdidx and e.cyear = '"&cyear&"' and e.cmonth='"&cmonth&"' group by s.contidx, e.cyear, e.cmonth"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		If rs.eof Then getmonthlyqty = "0" Else getmonthlyqty = rs(0)
		Set cmd = Nothing
	End Function

' 브랜드 이름 가져오기
	Function getbrand(thmno)
		If IsNull(thmno) Then getbrand= ""
		Dim sql : sql = "select c.highseqname from wb_subseq_dtl a inner join wb_subseq_mst b on a.subno=b.subno inner join sc_subseq_hdr c on b.seqno = c.highseqno where a.thmno = '" & thmno &"' "
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		Set cmd = nothing
		If rs.eof Then getbrand = "" Else getbrand = rs(0)
	End Function


' 서브브랜드 이믈 가져오기
	Function getsubbrand(thmno)
		If IsNull(thmno) Then getsubbrand= ""
		Dim sql : sql = "select b.subname from wb_subseq_dtl a inner join wb_subseq_mst b on a.subno=b.subno where a.thmno = '" & thmno &"' "
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		Set cmd = nothing
		If rs.eof Then getsubbrand = "" Else getsubbrand = rs(0)
	End Function

	'소재명 가져오기
	Function getthmname(thmno)
		If IsNull(thmno) Then getthmname= ""
		Dim sql : sql = "select thmname from wb_subseq_dtl  where thmno = '" & thmno &"' "
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		Set cmd = nothing
		If rs.eof Then getthmname = "" Else getthmname = rs(0)
	End Function

	' 선택된 매체의 년월에 집행한 브랜드 이름 가져오기
	Function getcurrentbrandname(mdidx, cyear, cmonth, side)
			Dim sql : sql = "select d.highseqname from wb_subseq_exe a inner join wb_subseq_dtl b on a.thmno = b.thmno inner join wb_subseq_mst c on b.subno = c.subno inner join sc_subseq_hdr d on c.seqno = d.highseqno where a.seq = (select max(seq) from wb_subseq_exe where mdidx = ? and side=? and cyear+cmonth <= '"&cyear&cmonth&"' )"
			If side = "" Then side = "F"
			Dim cmd : Set cmd = server.CreateObject("adodb.command")
			cmd.activeconnection = application("connectionstring")
			cmd.commandText = sql
			cmd.commandType = adCmdText
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput,, mdidx)
			cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1, side)
			cmd.prepared = true
			Dim rs : Set rs = cmd.execute
			Set cmd = Nothing

			If rs.eof Then getcurrentbrandname = "" Else getcurrentbrandname = rs(0)
	End Function

	function getcurrentstandard(mdidx, cyear, cmonth, side, part)
		dim sql : sql = "select standard, quality from wb_contact_md_dtl where seq=(select max(seq) from wb_contact_md_dtl where mdidx=? and side=? and cyear+cmonth<= ?)"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput,, mdidx)
		cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1, side)
		cmd.parameters.append cmd.createparameter("yearmonth", adchar, adparaminput, 6, cyear&cmonth)
		cmd.prepared = true
		Dim rs : Set rs = cmd.execute
		Set cmd = nothing
		if part = "standard" then
		If rs.eof Then getcurrentstandard = "" Else getcurrentstandard = rs(0)
		else
		If rs.eof Then getcurrentstandard = "" Else getcurrentstandard = getStringQuality(rs(1))
		end if
	end function

	function getStringQuality(code)
		dim sql : sql = "select codename from wb_code_library where category = 'quality' and codevalue = ?"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		cmd.parameters.append cmd.createparameter("codevalue", advarchar, adparaminput,10, code)
		cmd.prepared = true
		Dim rs : Set rs = cmd.execute
		Set cmd = nothing
		if rs.eof then getStringQuality = null else getStringQuality = rs(0)


	end function

	' 선택된 매체의 년월에 집행한 소재 이름 가져오기
	Function getcurrentthemename(mdidx, cyear, cmonth, side)
		Dim sql : sql = "select  thmname from wb_subseq_exe a inner join wb_subseq_dtl b on a.thmno = b.thmno where a.seq = (select max(seq)  from wb_subseq_exe where mdidx = ? and side=?  and cyear+cmonth <= '"&cyear&cmonth&"')"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput,, mdidx)
			cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1, side)
			cmd.prepared = true
		Dim rs : Set rs = cmd.execute
		Set cmd = nothing

		If rs.eof Then getcurrentthemename = "" Else getcurrentthemename = rs(0)
	End Function

' 바이트수로 계약명 길이 조정
	Function cutTitle(title, num)
		dim i, sum, title_one, result , sumByte
		If IsNull(title) Then title = ""

		for i = 1 to len(title)
			title_one = MID(title, i, 1)
			if ASC(title_one)<0 then sumByte = sumByte + 2 else sumByte = sumByte + 1
			if sumByte>num then result = result &"..." : exit for else result = result + title_one
		next
		cutTitle = result
	End Function

' 광고주 콤보 박스
	Sub getcustcombo(custcode)
		Dim sql : sql = "select highcustcode, custname from sc_cust_hdr where medflag='A' and use_flag=1 order by custname"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		Set cmd = Nothing
		response.write "<select id='cmbcustcombo' name='cmbcustcombo' >"
		Do Until rs.eof
			response.write "<option value='" & rs("highcustcode") & "' "
			If custcode = rs("highcustcode") Then response.write " selected"
			response.write ">" & rs("custname") & "</option>"
			rs.movenext
		loop
		response.write "</select>"
	End Sub

' 매체사 콤보 박스

	Sub getmedcombo(medcode)
		Dim sql : sql = "select distinct a.highcustcode, a.custname from sc_cust_hdr a inner join sc_cust_dtl b on a.highcustcode=b.highcustcode where a.medflag='B' and a.use_flag=1 and b.med_out = '1' order by a.custname"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		Set cmd = Nothing
		response.write "<select id='cmbmed' name='cmbmed' id='cmbmed' class='med' >"
		response.write "<option value=''>매체사를 선택하세요 </option>"
		Do Until rs.eof
			response.write "<option value='" & rs("highcustcode") & "' "
			If medcode = rs("highcustcode") Then response.write " selected"
			response.write ">" & rs("custname") & "</option>"
			rs.movenext
		Loop
		response.write "</select>"
	End Sub

	'매체사 직원 계정 발급
	Function getempid(medcode)
		If medcode = "" Then medcode = Null
		Dim sql : sql = "select max(empid) from wb_med_employee where medcode=?"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		cmd.parameters.append cmd.createparameter("medcode", adchar, adparaminput, 6)
		cmd.parameters("medcode").value = medcode
		Dim rs : Set rs = cmd.execute
		If IsNull(rs(0)) Then
			getempid = medcode&"001"
		Else
			Dim num
			num = CInt(Right(rs(0),3))+1
			Do Until Len(num) > 2
				num = "0" & num
			Loop
			getempid = medcode&num
		End If


	End Function

	' 지역 정보 콤보 박스

	Sub getregion(region)
		If region = "" Then region = Null
		Dim sql : sql = "select codevalue, codename from wb_code_library where category = 'region' "
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		Set cmd = Nothing
		response.write "<select id='cmbregion' name='cmbregion' class='region' >"
		Do Until rs.eof
			response.write "<option value='" & rs("codevalue") & "' "
			If region = rs("codevalue") Then response.write " selected"
			response.write ">" & rs("codename") & "</option>"
			rs.movenext
		loop
		response.write "</select>"
	End Sub

	' 광고 매체 재질 정보 가져오기

	Sub getquality(quality)
		Dim sql : sql = "select codevalue, codename from wb_code_library where category = 'quality' order by codevalue"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		Set cmd = Nothing
		response.write "<select id='cmbquality' name='cmbquality' class='quality' style='width:150px;'>"
		Do Until rs.eof
			response.write "<option value='" & rs("codevalue") & "' "
			If quality = rs("codevalue") Then response.write " selected"
			response.write ">" & rs("codename") & "</option>"
			rs.movenext
		loop
		response.write "</select>"
	End Sub

	'매체 분류 가져오기
	Function getmediumname(categoryidx)
		If categoryidx = "" Or IsNull(categoryidx) Then categoryidx = 0
		Dim sql : sql = "select categoryname from wb_category  where categoryidx = "& categoryidx
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		If rs.eof Then getmediumname = "" Else getmediumname = rs(0)
		Set cmd = Nothing
	End Function

	'Command Paramaters Clear
	Sub clearParameter(obj)
		Do Until  (obj.parameters.count = 0)
			obj.parameters.delete obj.parameters.count-1
			Loop
	End Sub

	' 개별 이미지 가져오기
	Function getimage(photo, wid)
		'Dim ext : ext = Right(photo, InstrRev(photo, "."))
		If IsNull(photo) Or photo="" Then
			getimage = "<img src='/images/noimage.gif' width='"&wid&"' id='photo' class='noimage'>"
		Else
			getimage = "<img src='/pds/media/"&photo&"' id='photo' width='"&wid&"' class='photo'>"
		End If
	End Function

	' Debug Mode
	Sub Debug
		Dim item
		For Each item In request.querystring
			response.write item & " : " & request.querystring(item) & "<br>"
		Next
		response.write "Err.Number : " & Err.number & "<br>"
		response.write "Err.Description : " & Err.Description & "<br>"
		response.write "Err.Source : " & Err.Source
	End Sub

	' 모니터링 상태 한글 변환
	Function getmonitorstatus(code)
		if isnull(code) or code = "" then
			getmonitorstatus = ""
			exit function
		end if
		If code = "1" Then getmonitorstatus = "양호" Else getmonitorstatus = "불량"
	End Function

	' 모니터링 이미지 가져오기
	Function getmonitorimg(mdidx, side, cyear, cmonth)
		Dim sql : sql = "select COALESCE(img01, img02, img03, img04, 'no')  from wb_contact_monitor  where mdidx=? and side=? and cyear=? and cmonth=?"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		cmd.parameters.append cmd.createparameter("mdidx", adInteger, adParaminput)
		cmd.parameters.append cmd.createparameter("side", adchar, adParaminput, 1)
		cmd.parameters.append cmd.createparameter("cyear", adchar, adParaminput, 4)
		cmd.parameters.append cmd.createparameter("cmonth", adchar, adParaminput, 2)
		cmd.parameters("mdidx").value = mdidx
		cmd.parameters("side").value = side
		cmd.parameters("cyear").value = cyear
		cmd.parameters("cmonth").value = cmonth
		Dim rs : Set rs = cmd.execute
		If rs.eof Then getmonitorimg = null Else getmonitorimg = rs(0)
		Set cmd = Nothing
	End Function

	'월자리수 만들기
	Function setmonth(m)
		If Len(m) = 1 Then
			setmonth = "0" & m
		Else
			setmonth = m
		End If
	End Function

	Function MakeUrlToFile(ori_url, save_filename, contidx, cyear, cmonth)

		Dim xh : Set xh = CreateObject("MSXML2.ServerXMLHTTP")
		xh.OPen "GET", ori_url, False
		  lResolve = 5 * 1000
		  lConnect = 5 * 1000
		  lSend = 15 * 10000
		  lReceive = 15 * 10000
		  xh.setTimeouts lResolve,lConnect,lSend,lReceive
		xh.Send()

		Dim strData : strData = xh.ResponseBody
		Set xh = Nothing

'		Dim save_root : save_root = "\\11.0.12.201\adportal\print\"&save_filename
		Dim save_root : save_root = "C:\pds\print\"&save_filename
'		Dim fso : Set fso = CreateObject("scripting.filesystemobject")
'		If fso.fileExists(save_root) Then
'			fso.deleteFile(save_root)
'		End If
'		Set fso = Nothing

		Dim s : Set s = CreateObject("adodb.stream")
		s.open()
		s.type = 1
		s.write strData
		s.SaveToFIle  save_root, 2
		Set s = Nothing

		Dim sql : sql = "select * from wb_report_mst where contidx=? and cyear=? and cmonth=?"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")

		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		cmd.parameters.append cmd.createparameter("contidx", adInteger, adParamInput)
		cmd.parameters.append cmd.createparameter("cyear", adChar, adParaminput, 4)
		cmd.parameters.append cmd.createparameter("cmonth", adChar, adParaminput, 2)
		cmd.parameters("contidx").value = contidx
		cmd.parameters("cyear").value = cyear
		cmd.parameters("cmonth").value = cmonth
		Dim rs : Set rs = cmd.execute
		clearparameter(cmd)

		If  rs.eof Then
			sql = "insert wb_report_mst (contidx, cyear, cmonth, report) values (?, ?, ?, ?)"
		cmd.parameters.append cmd.createparameter("contidx", adInteger, adParamInput)
		cmd.parameters.append cmd.createparameter("cyear", adChar, adParaminput, 4)
		cmd.parameters.append cmd.createparameter("cmonth", adChar, adParaminput, 2)
		cmd.parameters.append cmd.createparameter("report", adVarChar, adParaminput, 200)
		cmd.parameters("contidx").value = contidx
		cmd.parameters("cyear").value = cyear
		cmd.parameters("cmonth").value = cmonth
		cmd.parameters("report").value = save_filename
		response.write cmd.parameters.count
		Else
			sql = "update wb_report_mst set report=? where contidx=? and cyear=? and cmonth=?"
		cmd.parameters.append cmd.createparameter("report", adVarChar, adParaminput, 200)
		cmd.parameters.append cmd.createparameter("contidx", adInteger, adParamInput)
		cmd.parameters.append cmd.createparameter("cyear", adChar, adParaminput, 4)
		cmd.parameters.append cmd.createparameter("cmonth", adChar, adParaminput, 2)
		cmd.parameters("report").value = save_filename
		cmd.parameters("contidx").value = contidx
		cmd.parameters("cyear").value = cyear
		cmd.parameters("cmonth").value = cmonth
		End If
		cmd.commandText = sql
		cmd.execute ,, adExecuteNoRecords
		Set cmd = Nothing


		'response.BinaryWrite strData
	End Function


	Function MakeUrlToPrint(ori_url)

		Dim xh : Set xh = CreateObject("MSXML2.ServerXMLHTTP")
		xh.OPen "GET", ori_url, False
		  lResolve = 5 * 1000
		  lConnect = 5 * 1000
		  lSend = 15 * 10000
		  lReceive = 15 * 10000
		  xh.setTimeouts lResolve,lConnect,lSend,lReceive
		xh.Send()

		Dim strData : strData = xh.ResponseBody
		Set xh = Nothing

		response.BinaryWrite strData
	End Function

	Sub getReportFile(mdidx, cyear, cmonth)
		Dim sql : sql = "select filename from wb_report_dtl where mdidx =? and cyear=? and cmonth=?"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		cmd.parameters.append cmd.createparameter("mdidx", adInteger, adParamInput)
		cmd.parameters.append cmd.createparameter("cyear", adChar, adParamInput, 4)
		cmd.parameters.append cmd.createparameter("cmonth", adChar, adParamInput, 2)
		cmd.parameters("mdidx").value = mdidx
		cmd.parameters("cyear").value = cyear
		cmd.parameters("cmonth").value = cmonth
		Dim rs : Set rs = cmd.execute

		If Not rs.eof Then
		Dim filename : filename = rs(0)
			response.write "<a href='http://mms.raed.co.kr/med/download.asp?filename="&filename&"'><img src='http://mms.raed.co.kr/images/m_ppt.gif' width='16' height='16' align='absmiddle' ></a>"
		Else
			response.write "<img src='http://mms.raed.co.kr/images/m2_ppt.gif' width='16' height='16' align='absmiddle' >"
		End If
		Set rs = Nothing
		Set cmd = Nothing
	End Sub

	function getmedmonthly(contidx, medcode, cyear, cmonth)
		dim sql : sql = "select sum(isnull(monthly,0)) from wb_contact_md a inner join wb_contact_exe c on a.mdidx=c.mdidx and cyear=? and cmonth=?   where a.medcode =? and contidx = ? group by medcode"
		dim cmd : set cmd = server.createobject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		cmd.parameters.append cmd.createparameter("cyear", adChar, adParaminput, 4)
		cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2)
		cmd.parameters.append cmd.createparameter("medcode", adChar, adParaminput, 6)
		cmd.parameters.append cmd.createparameter("contidx", adinteger, adParaminput)

		cmd.parameters("cyear").value = cyear
		cmd.parameters("cmonth").value = cmonth
		cmd.parameters("medcode").value = medcode
		cmd.parameters("contidx").value = contidx

		dim rs : set rs = cmd.execute
		if not rs.eof then getmedmonthly = rs(0) else getmedmonthly = 0
		rs.close
		set rs = nothing
		set cmd = nothing
	end function


%>