<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '���Ȱ��ö��̺귯�� %>
<!--#include virtual="/inc/head.asp"-->			<% '�ʱ� ���� ������(���� �޼��� �����) %>

<%
'	Dim item
'	For Each item In request.Form
'		response.write item & " : " & request.Form(item) & "<br>"
'	Next
'	response.end

'����ÿ� �ʿ��� ��� ������ �����´�.
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

	' ��� ���� ������ �����Ѵ�.
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

	contidx = objrs.fields("contidx").value					' ����� ����Ϸù�ȣ

	objrs.close

	' ���� ������ ��� �ʱ� ������ �����´�.
	sql = "select contidx,  idx,  monthprice, expense, qty from dbo.wb_contact_tmp where contidx = " & org_contidx
	call set_recordset(objrs, sql)


	if not objrs.eof then
		set org_sidx = objrs("idx")							' ��࿡ ��ϵ� ��ü ����
		set monthprice = objrs("monthprice")			' ����ü�� �������
		set expense = objrs("expense")					' ����ü�� �����޾�
		set qty = objrs("qty")									' ����ü�� ������
	end if



	' ��࿡ �ش��ϴ� ��� ��ü�� �� ����Ѵ�.
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

			sidx = objrs2("sidx")						' ���ο� ����ü �Ϸù�ȣ

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

				' ********** ��� �� �������� ���Ѵ�.
				totalContactMonth = 0
				org_startDate = startDate
				org_endDate = endDate

				totalContactMonth = DateDiff("m", startDate, endDate)
				if (Day(org_startDate) = "01") then totalContactMonth = totalContactMonth - 1 end if
				' ********** ��� �Ⱓ�� ���� ��ŭ ��,���� ������Ű�鼭 �麰 ������ �Է��Ѵ�.
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

					' ********** ��� �������� 1�� ���� ������ ���
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

					org_startDate = DateAdd("m", 1, org_startDate)  ' ���� ����� �Ѵ޾� ������Ų��.
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