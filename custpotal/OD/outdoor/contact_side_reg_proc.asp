<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%

'	dim item
'	for each item in request.form
'		response.write item & " : " & request.form(item) & "<br>"
'	next
'	response.end
	dim idx
	dim sidx : sidx = request("sidx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	dim side : side = request("selside")
	dim qty : qty = request("txtqty")
	dim thema : thema = request("selsubject")
	dim unitprice : unitprice = request("txtunitprice")
	dim monthprice : monthprice = request("txtmonthprice")
	dim unit : unit = request("txtunit")
	dim standard : standard = request("txtstandard")
	dim quality : quality = request("selquality")
	dim expense : expense = request("txtexpense")

	dim totalprice, totalqty, lastmonth, intLoop

	if side = "" then side = null
	if unitprice = "" then unitprice = 0 else unitprice = replace(unitprice, ",","")
	if quality = "" then quality = null
	if monthprice = "" then monthprice = 0 else monthprice = replace(monthprice, ",","")
	if expense = "" then expense = 0 else expense = replace(expense, ",","")
	if thema = "" then thema = null

	dim totalContactMonth ' ��� �Ѱ�����
	dim tmpMonthPrice		' �ӽ� �������
	dim tmpExpense			' �ӽ� �����޾�
	dim tmpCyear
	dim tmpCmonth

	dim objrs, sql
	sql = "select m.contidx, startdate, enddate from dbo.wb_contact_mst m inner join dbo.wb_contact_md m2 on m.contidx = m2.contidx where m2.sidx = " & sidx
	call get_recordset(objrs, sql)

	dim contidx, startdate , enddate
	contidx = objrs("contidx")
	startdate = objrs("startdate")
	enddate = objrs("enddate")

	objrs.close


	' ****************************************************************
	' ********** ��� ��ü�� ��(����) ������ ����Ѵ�.

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
	' ********** ��� ��ü�� ��(����) �ݾ� ������ ����Ѵ�.

	sql = "select idx, cyear, cmonth, qty, jobidx, monthprice, expense, photo_1, photo_2, photo_3, photo_4, isPerform, performDate, performuser, isCancel, canceldate, canceluser, isClosing, closingdate from dbo.wb_contact_md_dtl_account where idx = " & idx

	call set_recordset(objrs, sql)

	' ********** ��� �� �������� ���Ѵ�.
	totalContactMonth = DateDiff("m", startDate, endDate)
	if (Day(startDate) = "01") then totalContactMonth = totalContactMonth - 1 end if

	' ********** ��� �Ⱓ�� ���� ��ŭ ��,���� ������Ű�鼭 �麰 ������ �Է��Ѵ�.
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

		' ********** ��� �������� 1�� ���� ������ ���
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
		startDate = DateAdd("m", 1, startDate)  ' ���� ����� �Ѵ޾� ������Ų��.
	Next

	objrs.close

	' ********** ���ʷ� ��ϵ� ��� ��� ����� �ʱ�ȭ �ݾ��� ����ϱ� ���� �ӽ����̺� �ݾ��� �����Ѵ�..
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
	window.opener.location.href="pop_contact_view.asp?contidx=<%=contidx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
	this.close();
//-->
</script>