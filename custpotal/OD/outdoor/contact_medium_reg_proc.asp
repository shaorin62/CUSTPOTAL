<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>
<!--#include virtual="/inc/head.asp"-->			<% '초기 설정 페이지(에러 메세지 미출력) %>

<%
	' ****************************************************************
	' ******************************
	dim uploadform																						'자료 업로드용 컴포넌트
	dim contidx																								' 매체가 등록될 계약일련번호
	dim sidx																									' 계약별로 등록되는 매체의 일련번호
	Dim atag																								' xss관련 함수 허용문자를 써주면 됨
	dim totlaprice																							' 총 광고료
	dim tmpFile																								' 파일명을 저장하기 위한 임시파일명
	dim totalContactMonth																				' 계약 총개월수
	dim intLoop
	dim startDate																							' 계약 시작일
	dim endDate																							' 계약 종료일
	dim tmpMonthPrice																					' 임시 월광고료
	dim tmpExpense																						' 임시 월지급액
	set uploadform = server.createobject("DEXT.FileUpload")
	uploadform.defaultpath = "C:\map"						' 약도 저장 폴더 (/map/)

	dim region : region = uploadform("selregion")												' 지역
	dim locate : locate = uploadform("txtlocate")												' 설치위치
	dim categoryidx : categoryidx = clearXSS( uploadform("txtcategoryidx"), atag)						' 매체(면) 분류
	dim medcode : medcode = uploadform("selcustcode")									' 매체사
	dim map : map = uploadform("txtmap")														' 약도

	' 첨부파일에 등록가능 여부 판단
	Dim strFileChk
	If map = "" Then
		map = Null
	Else
		strFileChk = Check_Ext(map,"JPG,GIF,PNG")

		If strFileChk  = "error" Then
			Response.write "<script>"
			Response.write "alert('등록할 수 없는 파일입니다.\n\n이미지 파일(JPG,GIF,PNG)만 등록하십시오.');"
			Response.write " this.close();"
			Response.write "</script>"
			Response.End
		End if
	End If
	
	dim trust : trust = uploadform("rdotrust")													' 구분(정책, 일반)
	dim side : side = uploadform("selside")														' 면(L, R, F, B)
	dim unitprice : unitprice = uploadform("txtunitprice")									' 매체 단가
	dim qty : qty = uploadform("txtqty")															' 계약 수량
	dim unit : unit = uploadform("txtunit")														' 매체 단위
	dim standard : standard = clearXSS( uploadform("txtstandard")	, atag)								' 매체 규격
	dim quality : quality = uploadform("selquality")												' 매체 재질
	dim monthprice : monthprice = uploadform("txtmonthprice")							' 월 광고료
	dim expense : expense = uploadform("txtexpense")										' 월 지급액
	dim thema : thema = uploadform("selsubject")											' 집행 소재
	dim cyear : cyear = uploadform("cyear")
	dim cmonth : cmonth = cint(uploadform("cmonth"))
	dim objrs
	dim sql
	dim idx																									' 매체 면별 일련번호 코드
	dim tmpCyear
	dim tmpCmonth


	if region = "" then region = null
	if locate = "" then locate = null
	if map = "" then
		map = null														' 약도저장하지 않는다.
	else
		tmp = uploadform("txtmap").save(, false)			' 약도 파일에 동일한 파일이 존재하면 새로운 파일명으로 저장한다.
		map = right(tmp, len(tmp)-InStrRev(tmp, "\"))	' 새로운 파일 명를 추출
	end if
	if side = "" then side = null
	if unitprice = "" then unitprice = 0 else unitprice = replace(unitprice, ",","")
	if quality = "" then quality = null
	if monthprice = "" then monthprice = 0 else monthprice = replace(monthprice, ",","")
	if expense = "" then expense = 0 else expense = replace(expense, ",","")
	if thema = "" then thema = null

	contidx = uploadform("contidx")								'계약번호

	' ****************************************************************
	' ********** 계약 정보에서 시작일과 종료일을 가져온다.

	sql = "select startdate, enddate from dbo.wb_contact_mst where contidx = " & contidx

	call get_recordset(objrs, sql)

		startdate = objrs("startdate")
		enddate = objrs("enddate")

	objrs.close

	' ****************************************************************
	' ********** 계약 매체를 등록한다.

	sql = "select contidx, sidx, region ,medcode, locate, categoryidx, unit, trust, map, cuser, cdate, uuser, udate from dbo.wb_contact_md where contidx = " & contidx
	call set_recordset(objrs, sql)

	objrs.addnew
		objrs("contidx") = contidx
		objrs("region") = region
		objrs("locate") = locate
		objrs("categoryidx") = categoryidx
		objrs("medcode") = medcode
		objrs("unit") = unit
		objrs("trust") = trim(trust)
		objrs("map") = map
		objrs("cuser") = request.cookies("userid")
		objrs("cdate") = date
		objrs("uuser") = request.cookies("userid")
		objrs("udate") = date
	objrs.update

	sidx = objrs("sidx")														' 계약 매체의 일련번호(계약별 매체등록번호)

	objrs.close


	' ****************************************************************
	' ********** 계약 매체의 면(세부) 정보를 등록한다.

	sql = "select idx, sidx, side, unitprice, standard, quality from dbo.wb_contact_md_dtl d where sidx = " & sidx
	call set_recordset(objrs, sql)

		objrs.addnew
		objrs("sidx") = sidx
		objrs("side") = side
		objrs("unitprice") = unitprice
		objrs("standard") = standard
		objrs("quality") = quality
		objrs.update
	idx = objrs("idx")
	objrs.close

	' ****************************************************************
	' ********** 계약 매체의 광고비 및 정산관련 정보를 등록한다.

	sql = "select idx, cyear, cmonth, qty, jobidx, monthprice, expense, photo_1, photo_2, photo_3, photo_4, isPerform, performDate, performuser, isCancel, canceldate, canceluser, isClosing, closingdate from dbo.wb_contact_md_dtl_account where idx = " & idx
	call set_recordset(objrs, sql)

	' ********** 계약 총 개월수를 구한다.
	totalContactMonth = DateDiff("m", startDate, endDate)
	if (Day(startDate) = "01") then totalContactMonth = totalContactMonth - 1 end if

	' ********** 계약 기간의 월수 만큼 년,월을 증가시키면서 면별 정보를 입력한다.
	For intLoop = 0 to totalContactMonth
	if Len(Month(startDate)) = 1 then
		tmpCmonth = "0"&Month(startDate)
	else
		tmpCmonth = Month(startDate)
	end if
	tmpCyear = Year(startDate)

		objrs.addnew
		objrs("idx") = idx
		objrs("cyear") = tmpCyear
		objrs("cmonth") = tmpCmonth
		objrs("qty") = qty
		objrs("jobidx") = thema

		' ********** 계약 시작일이 1일 부터 시작인 경우
		if intLoop = 0 then
			if Day(startDate) = 1 then
				tmpMonthPrice = monthprice
				tmpExpense = expense
			elseif (Day(startDate) > 1 and Day(startDate) < 15) or (Day(startDate)  > 15) then
				tmpMonthPrice = 0
				tmpExpense = 0
			else
				tmpMonthPrice = monthprice
				tmpExpense = expense
			end if
		elseif intLoop = totalContactMonth then
			if Day(startDate) = 1 then
				tmpMonthPrice = monthprice
				tmpExpense = expense
			elseif Day(startDate)  < 15 then
				tmpMonthPrice = 0
				tmpExpense = 0
			else
				tmpMonthPrice = monthprice
				tmpExpense = expense
			end if
		else
			tmpMonthPrice = monthprice
			tmpExpense = expense
		end if

		objrs("monthprice") = tmpMonthPrice
		objrs("expense") = tmpExpense
		objrs("photo_1") = null
		objrs("photo_2") = null
		objrs("photo_3") = null
		objrs("photo_4") = null
		objrs("isPerform") = 0
		objrs("performDate") = null
		objrs("performuser") = null
		objrs("isCancel") = 0
		objrs("canceldate") = null
		objrs("canceluser") = null
		objrs("isClosing") = 0
		objrs("closingdate") = null

		objrs.update
		startDate = DateAdd("m", 1, startDate)  ' 저장 년월을 한달씩 증가시킨다.
	Next

	objrs.close

	' ********** 최초로 등록된 경우 계약 연장시 초기화 금액을 등록하기 위한 임시테이블에 금액을 저장한다..
	sql = "select contidx, idx, monthprice, expense, qty from dbo.wb_contact_tmp where contidx="&contidx

	call set_recordset(objrs, sql)

	objrs.addnew
	objrs("contidx") = contidx
	objrs("idx") = sidx
	objrs("monthprice") = monthprice
	objrs("expense") = expense
	objrs("qty") = qty
	objrs.update

	objrs.close

	set objrs = nothing
%>
<script language="JavaScript">
<!--
	window.opener.location.href="/od/outdoor/pop_contact_view.asp?contidx=<%=contidx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
	this.close();
//-->
</script>