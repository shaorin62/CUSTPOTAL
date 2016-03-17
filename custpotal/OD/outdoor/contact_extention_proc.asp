<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>
<!--#include virtual="/inc/head.asp"-->			<% '초기 설정 페이지(에러 메세지 미출력) %>

<%
'	Dim item
'	For Each item In request.Form
'		response.write item & " : " & request.Form(item) & "<br>"
'	Next
'	response.end

'연장시에 필요한 계약 정보를 가져온다.
	Dim atag
	Call dbcon

	dim totalContactMonth, tmpCyear, tmpCmonth, tmpMonthPrice, tmpExpense, tmpQty
	dim contidx : contidx = request("contidx")
	dim org_contidx : org_contidx = contidx
	dim title : title = clearXSS( request("txttitle"), atag)
	dim startdate : startdate = clearXSS( request("txtstartdate"), atag)
	dim firstdate : firstdate = clearXSS( request("txtfirstdate"), atag)
	dim enddate : enddate = clearXSS( request("txtenddate"), atag)
	dim custcode : custcode = request("txtdeptcode")
	dim mediummemo : mediummemo = clearXSS( request("txtmediummemo"), atag)
	dim regionmemo : regionmemo = clearXSS( request("txtregionmemo"), atag)
	dim comment : comment =clearXSS(  request("txtcomment"), atag)
	dim reporttype : reporttype =clearXSS(  request("reporttype"), atag)

	if mediummemo = "" then mediummemo = null
	if regionmemo = "" then regionmemo = null
	if comment = "" then comment = null

	dim sidx
	dim idx
	dim objrs, sql
	dim org_sidx, side, monthprice, expense, qty , objrs2, intLoop
	Dim objrs3, objrs4
	Dim org_idx
	Dim org_startDate, org_endDate

	' 계약 기초 정보를 저장한다.
	sql = "select contidx, custcode, title, firstdate, startdate, enddate,  regionmemo, mediummemo, comment, reporttype, cancel,canceldate, cuser, cdate, uuser, udate from dbo.wb_contact_mst where contidx = " & org_contidx

	call set_recordset(objrs, sql)

	objrs.addnew
	objrs.fields("custcode").value = custcode
	objrs.fields("title").value = title
	objrs.fields("firstdate").value = firstdate
	objrs.fields("startdate").value = startdate
	objrs.fields("enddate").value = enddate
	objrs.fields("comment").value = comment
	objrs.fields("regionmemo").value = regionmemo
	objrs.fields("mediummemo").value = mediummemo
	objrs.fields("cancel").value = 0
	objrs.fields("canceldate").value = enddate
	objrs.fields("reporttype").value = reporttype
	objrs.fields("cuser").value = request.cookies("userid")
	objrs.fields("cdate").value = date
	objrs.fields("uuser").value = request.cookies("userid")
	objrs.fields("udate").value = date
	objrs.update

	contidx = objrs.fields("contidx").value					' 연장된 계약일련번호

	objrs.close

	' 연장 이전의 계약 초기 정보를 가져온다.
	sql = "select contidx,  idx,  monthprice, expense, qty from dbo.wb_contact_tmp where contidx = " & org_contidx
	call set_recordset(objrs, sql)


	if not objrs.eof then
		set org_sidx = objrs("idx")							' 계약에 등록된 매체 정보
		set monthprice = objrs("monthprice")			' 계약매체당 월광고료
		set expense = objrs("expense")					' 계약매체당 월지급액
		set qty = objrs("qty")									' 계약매체당 계약수량
	end if



	' 계약에 해당하는 모든 매체를 재 등록한다.
	do until objrs.eof
		sql = "select contidx, sidx, locate, categoryidx, medcode, region, unit, trust, map, cuser, cdate, uuser, udate from dbo.wb_contact_md where sidx = " &   org_sidx

		call set_recordset(objrs2, sql)
		if not objrs2.eof then
			dim locate : locate = objrs2("locate")
			dim categoryidx : categoryidx = objrs2("categoryidx")
			dim medcode : medcode = objrs2("medcode")
			dim region : region = objrs2("region")
			dim unit : unit = objrs2("unit")
			dim trust : trust = objrs2("trust")
			dim map : map = objrs2("map")


			objrs2.addnew
			objrs2("contidx") = contidx
			objrs2("locate") = locate
			objrs2("categoryidx") = categoryidx
			objrs2("medcode") = medcode
			objrs2("region") = region
			objrs2("unit") = unit
			objrs2("trust") = trust
			objrs2("map") = map
			objrs2("cuser") = request.cookies("userid")
			objrs2("cdate") = date
			objrs2("uuser") = request.cookies("userid")
			objrs2("udate") = date
			objrs2.update

			sidx = objrs2("sidx")						' 새로운 계약매체 일련번호

			sql = "select idx, sidx, side, unitprice, standard, quality from dbo.wb_contact_md_dtl where sidx = " &  org_sidx
			call set_recordset(objrs3, sql)


			if not objrs3.eof Then
				dim unitprice : unitprice = objrs3("unitprice")
				dim standard : standard = objrs3("standard")
				dim quality : quality = objrs3("quality")
				if side = "" then side = null

				objrs3.addnew
				objrs3("sidx") = sidx
				objrs3("side") = side
				objrs3("unitprice") = unitprice
				objrs3("standard") = standard
				objrs3("quality") = quality
				objrs3.update
				idx = objrs3("idx")

				sql="select idx, cyear, cmonth, qty, jobidx, monthprice, expense, photo_1, photo_2, photo_3, photo_4, isPerform, performdate, performuser, isCancel, canceldate, canceluser, isClosing, closingdate from dbo.wb_contact_md_dtl_account where idx = " & idx
				call set_recordset(objrs4, sql)

				' ********** 계약 총 개월수를 구한다.
				totalContactMonth = 0
				org_startDate = startDate
				org_endDate = endDate

				totalContactMonth = DateDiff("m", startDate, endDate)
				if (Day(org_startDate) = "01") then totalContactMonth = totalContactMonth - 1 end if
				' ********** 계약 기간의 월수 만큼 년,월을 증가시키면서 면별 정보를 입력한다.
				For intLoop = 0 to totalContactMonth
					if Len(Month(org_startDate)) = 1 then
						tmpCmonth = "0"&Month(org_startDate)
					else
						tmpCmonth = Month(org_startDate)
					end if
					tmpCyear = Year(org_startDate)

					objrs4.addnew
					objrs4("idx") = idx
					objrs4("cyear") = tmpCyear
					objrs4("cmonth") = tmpCmonth
					objrs4("qty") = qty
					objrs4("jobidx") = null

					' ********** 계약 시작일이 1일 부터 시작인 경우
					if intLoop = 0 then
						if Day(org_startDate) = 1 then
							tmpMonthPrice = monthprice
							tmpExpense = expense
						elseif (Day(org_startDate) > 1 and Day(org_startDate) < 15) or (Day(org_startDate)  > 15) then
							tmpMonthPrice = 0
							tmpExpense = 0
						else
							tmpMonthPrice = monthprice
							tmpExpense = expense
						end if
					elseif intLoop = totalContactMonth then
						if Day(org_startDate) = 1 then
							tmpMonthPrice = monthprice
							tmpExpense = expense
						elseif Day(org_startDate)  < 15 then
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

					objrs4("monthprice") = tmpMonthPrice
					objrs4("expense") = tmpExpense
					objrs4("photo_1") = null
					objrs4("photo_2") = null
					objrs4("photo_3") = null
					objrs4("photo_4") = null
					objrs4("isPerform") = 0
					objrs4("performDate") = null
					objrs4("performuser") = null
					objrs4("isCancel") = 0
					objrs4("canceldate") = null
					objrs4("canceluser") = null
					objrs4("isClosing") = 0
					objrs4("closingdate") = null

					objrs4.update

					org_startDate = DateAdd("m", 1, org_startDate)  ' 저장 년월을 한달씩 증가시킨다.
				Next
				objrs2.close
				objrs3.close
				objrs4.close
			End if
		End if
		sql=" insert into wb_contact_tmp values(" & contidx & ", " & sidx & ", " & monthprice & ", " & expense & ", " & qty & ")"
		objconn.execute(sql)
'		response.write sql & "<p> "
		objrs.movenext
	Loop

	objrs.close

	set objrs4 = Nothing
	set objrs3 = nothing
	set objrs2 = nothing
	set objrs = Nothing

	Call dbclose
%>
<script language="JavaScript">
<!--
	this.close();
//-->
</script>