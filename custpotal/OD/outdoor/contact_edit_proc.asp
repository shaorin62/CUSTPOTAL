<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
'	Dim item
'	For Each item In request.Form
'		response.write item & " : " & request.Form(item) & "<br>"
'	Next
'	response.end

	dim contidx : contidx = request("contidx")
	dim title : title = request("txttitle")
	dim startdate : startdate = request("txtstartdate")
	dim mediummemo : mediummemo = request("txtmediummemo")
	dim firstdate : firstdate = request("txtfirstdate")
	dim enddate : enddate = request("txtenddate")
	dim custcode : custcode = request("txtdeptcode")
	dim comment : comment = request("txtcomment")
	dim regionmemo : regionmemo = request("txtregionmemo")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")

	if mediummemo = "" then mediummemo = null
	if regionmemo = "" then regionmemo = null
	if comment = "" then comment = null

	dim nStartDate : nStartDate = startdate
	dim oStartDate
	dim nEndDate : nEndDate = enddate
	dim oEndDate

	dim objrs, objrs2, sql

	sql = "select  custcode, title, firstdate, startdate, enddate,  canceldate, regionmemo, mediummemo, comment,  uuser, udate from dbo.wb_contact_mst where contidx =" &contidx

	call get_recordset(objrs, sql)

	oStartDate = objrs("startdate")
	oEndDate = objrs("endDate")

	objrs.close

' ********************* 시작 변경 **********************************

	if datediff("m", oStartDate, nStartdate) > 0   then
		sql = "select * from dbo.wb_contact_md m inner join dbo.wb_contact_md_dtl d on m.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on d.idx = a.idx where a.isPerform = 1 and m.contidx=" & contidx & " and (a.cyear+a.cmonth between '" & oStartDate & "' and '" & nStartDate &"') "
		call get_recordset(objrs, sql)
		if not objrs.eof then
			response.write "<script> alert('수정하려는 기간 중에 정산내역이 존재합니다.\n\n정산내역을 취소한 후 수정하세요'); this.close(); </script>"
			respons.end
		end if
	end if
	if datediff("m", oEndDate , nEndDate) < 0   then
		sql = "select * from dbo.wb_contact_md m inner join dbo.wb_contact_md_dtl d on m.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on d.idx = a.idx where a.isPerform = 1 and m.contidx=" & contidx & " and (a.cyear+a.cmonth between '" & nEndDate & "' and '" & oEndDate &"') "
		call get_recordset(objrs, sql)
		if not objrs.eof then
			response.write "<script> alert('수정하려는 기간 중에 정산내역이 존재합니다.\n\n정산내역을 취소한 후 수정하세요'); this.close(); </script>"
			respons.end
		end if
	end if

	dim blankStartDate : blankStartDate = DateDiff("m", oStartDate, nStartDate)

	dim jobidx, qty, tmpMonthPrice, tmpExpense, idx


	if blankStartDate > 0 then ' 기존 날짜 보다 이후 날짜 이므로 해당 일자의 내역이 삭제 되어야 한다.

	sql = "select a.idx, a.cyear, a.cmonth  from dbo.wb_contact_md m inner join dbo.wb_contact_md_dtl d on m.sidx = d.sidx inner join  dbo.wb_contact_md_dtl_account a  on d.idx = a.idx where cyear+cmonth < '" & DateAdd("m", -1, nStartDate) & "' and m.contidx = " & contidx
	call get_recordset(objrs, sql)

		if not objrs.eof then
			do until objrs.eof
				sql = "select * from dbo.wb_contact_md_dtl_account where idx = " & objrs("idx") & " and cyear = '" & objrs("cyear") & "' and cmonth = '" & objrs("cmonth") & "' "
				call set_recordset(objrs2, sql)

				if not objrs2.eof then
					do until objrs2.eof
						objrs2.delete
					objrs2.movenext
					Loop
				end if
			objrs.movenext
			Loop
		end if
	objrs.close
	end if

	if blankStartDate < 0 then ' 기존 날짜 보다 이전 날짜 이므로 해당 기간이 추가 되어야 한다.

	sql = "select contidx, d.idx, monthprice, expense, qty from dbo.wb_contact_tmp  t inner join dbo.wb_contact_md_dtl d on t.idx = d.sidx  where contidx = " & contidx
	call get_recordset(objrs, sql)

	if not objrs.eof then
		sql = "select idx, cyear, cmonth, qty, jobidx, monthprice, expense, photo_1, photo_2, photo_3, photo_4, isPerform, performDate, performuser, isCancel, canceldate, canceluser, isClosing, closingdate from dbo.wb_contact_md_dtl_account where idx = " & objrs("idx")
		call set_recordset(objrs2, sql)

		jobidx = objrs2("jobidx")
		do until objrs.eof

			do until DateDiff("m", oStartDate, nStartDate) = 0

			objrs2.addnew
			objrs2("idx") = objrs("idx")
			objrs2("cyear") = left(nStartDate,4)
			objrs2("cmonth") = mid(nStartDate,6,2)
			objrs2("qty") = objrs("qty")
			objrs2("jobidx") = jobidx

			if Day(startdate) = 1 then
				tmpMonthPrice = objrs("monthprice")
				tmpExpense = objrs("expense")
			elseif (Day(startdate) > 1 and Day(startdate) < 15) or (Day(startdate)  > 15) then
				tmpMonthPrice = 0
				tmpExpense = 0
			else
				tmpMonthPrice = objrs("monthprice")
				tmpExpense =  objrs("expense")
			end if

			objrs2("monthprice") = tmpMonthPrice
			objrs2("expense") = tmpExpense
			objrs2("photo_1") = null
			objrs2("photo_2") = null
			objrs2("photo_3") = null
			objrs2("photo_4") = null
			objrs2("isPerform") = 0
			objrs2("performDate") = null
			objrs2("performuser") = null
			objrs2("isCancel") = 0
			objrs2("canceldate") = null
			objrs2("canceluser") = null
			objrs2("isClosing") = 0
			objrs2("closingdate") = null
			Response.write nStartDate & " , " & objrs("idx") & "<br>"

			objrs2.update
			nStartDate = DateAdd("m", 1, nStartDate)
			Loop

		objrs.movenext
		Loop
	end if

	end if

'************************** 종료일 변경

	dim blankEndDate : blankEndDate = DateDiff("m", nEndDate, oEndDate)


	if blankEndDate > 0 then ' 종료일자가 줄어들게 되므로 변경 날짜 이후는 삭제 한다.
	sql = "select a.idx, a.cyear, a.cmonth  from dbo.wb_contact_md m inner join dbo.wb_contact_md_dtl d on m.sidx = d.sidx inner join  dbo.wb_contact_md_dtl_account a  on d.idx = a.idx where cyear+cmonth > '" & nEndDate & "' and m.contidx = " & contidx

	call get_recordset(objrs, sql)

		if not objrs.eof then
			do until objrs.eof
				sql = "select * from dbo.wb_contact_md_dtl_account where idx = " & objrs("idx") & " and cyear = '" & objrs("cyear") & "' and cmonth = '" & objrs("cmonth") & "' "
				call set_recordset(objrs2, sql)

				if not objrs2.eof then
					do until objrs2.eof
						objrs2.delete
					objrs2.movenext
					Loop
				end if
			objrs.movenext
			Loop
		end if
	objrs.close
	end if

	if blankEndDate < 0 then '종료일자가 늘어나므로 날짜를 추가한다.

	sql = "select contidx, d.idx, monthprice, expense, qty from dbo.wb_contact_tmp  t inner join dbo.wb_contact_md_dtl d on t.idx = d.sidx  where contidx = " & contidx
	call get_recordset(objrs, sql)

	if not objrs.eof then
		sql = "select idx, cyear, cmonth, qty, jobidx, monthprice, expense, photo_1, photo_2, photo_3, photo_4, isPerform, performDate, performuser, isCancel, canceldate, canceluser, isClosing, closingdate from dbo.wb_contact_md_dtl_account where idx = " & objrs("idx")
		call set_recordset(objrs2, sql)

		jobidx = objrs2("jobidx")
		do until objrs.eof

			do until DateDiff("m", nEndDate, oEndDate) = 0
			oEndDate = DateAdd("m", 1, oEndDate)

			objrs2.addnew
			objrs2("idx") = objrs("idx")
			objrs2("cyear") = left(oEndDate,4)
			objrs2("cmonth") = mid(oEndDate,6,2)
			objrs2("qty") = objrs("qty")
			objrs2("jobidx") = jobidx

			if Day(enddate)  < 15  then
				tmpMonthPrice = 0
				tmpExpense = 0
			else
				tmpMonthPrice = objrs("monthprice")
				tmpExpense =  objrs("expense")
			end if

			objrs2("monthprice") = tmpMonthPrice
			objrs2("expense") = tmpExpense
			objrs2("photo_1") = null
			objrs2("photo_2") = null
			objrs2("photo_3") = null
			objrs2("photo_4") = null
			objrs2("isPerform") = 0
			objrs2("performDate") = null
			objrs2("performuser") = null
			objrs2("isCancel") = 0
			objrs2("canceldate") = null
			objrs2("canceluser") = null
			objrs2("isClosing") = 0
			objrs2("closingdate") = null
			Response.write oEndDate & " , " & objrs("idx") & "<br>"

			objrs2.update
			Loop

		objrs.movenext
		Loop
	end if

	End if

	sql = "select  custcode, title, firstdate, startdate, enddate,  canceldate, regionmemo, mediummemo, comment,  uuser, udate from dbo.wb_contact_mst where contidx =" &contidx

	call set_recordset(objrs, sql)

	oStartDate = objrs("startdate")
	oEndDate = objrs("endDate")

	objrs("custcode") = custcode
	objrs("title") = title
	objrs("firstdate") = firstdate
	objrs("startdate") = startdate
	objrs("enddate") = enddate
	objrs("canceldate") = enddate
	objrs("mediummemo") = mediummemo
	objrs("regionmemo") = regionmemo
	objrs("comment") = comment
	objrs("uuser") = request.cookies("userid")
	objrs("udate") = date
	objrs.update

	objrs.close


	set objrs = nothing
%>
<script language="JavaScript">
<!--
	window.opener.location.href="pop_contact_view.asp?contidx=<%=contidx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
	this.close();
//-->
</script>