<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '���Ȱ��ö��̺귯�� %>
<!--#include virtual="/inc/head.asp"-->			<% '�ʱ� ���� ������(���� �޼��� �����) %>

<%
	' ****************************************************************
	' ******************************
	dim uploadform																						'�ڷ� ���ε�� ������Ʈ
	dim contidx																								' ��ü�� ��ϵ� ����Ϸù�ȣ
	dim sidx																									' ��ະ�� ��ϵǴ� ��ü�� �Ϸù�ȣ
	Dim atag																								' xss���� �Լ� ��빮�ڸ� ���ָ� ��
	dim totlaprice																							' �� �����
	dim tmpFile																								' ���ϸ��� �����ϱ� ���� �ӽ����ϸ�
	dim totalContactMonth																				' ��� �Ѱ�����
	dim intLoop
	dim startDate																							' ��� ������
	dim endDate																							' ��� ������
	dim tmpMonthPrice																					' �ӽ� �������
	dim tmpExpense																						' �ӽ� �����޾�
	set uploadform = server.createobject("DEXT.FileUpload")
	uploadform.defaultpath = "C:\map"						' �൵ ���� ���� (/map/)

	dim region : region = uploadform("selregion")												' ����
	dim locate : locate = uploadform("txtlocate")												' ��ġ��ġ
	dim categoryidx : categoryidx = clearXSS( uploadform("txtcategoryidx"), atag)						' ��ü(��) �з�
	dim medcode : medcode = uploadform("selcustcode")									' ��ü��
	dim map : map = uploadform("txtmap")														' �൵

	' ÷�����Ͽ� ��ϰ��� ���� �Ǵ�
	Dim strFileChk
	If map = "" Then
		map = Null
	Else
		strFileChk = Check_Ext(map,"JPG,GIF,PNG")

		If strFileChk  = "error" Then
			Response.write "<script>"
			Response.write "alert('����� �� ���� �����Դϴ�.\n\n�̹��� ����(JPG,GIF,PNG)�� ����Ͻʽÿ�.');"
			Response.write " this.close();"
			Response.write "</script>"
			Response.End
		End if
	End If
	
	dim trust : trust = uploadform("rdotrust")													' ����(��å, �Ϲ�)
	dim side : side = uploadform("selside")														' ��(L, R, F, B)
	dim unitprice : unitprice = uploadform("txtunitprice")									' ��ü �ܰ�
	dim qty : qty = uploadform("txtqty")															' ��� ����
	dim unit : unit = uploadform("txtunit")														' ��ü ����
	dim standard : standard = clearXSS( uploadform("txtstandard")	, atag)								' ��ü �԰�
	dim quality : quality = uploadform("selquality")												' ��ü ����
	dim monthprice : monthprice = uploadform("txtmonthprice")							' �� �����
	dim expense : expense = uploadform("txtexpense")										' �� ���޾�
	dim thema : thema = uploadform("selsubject")											' ���� ����
	dim cyear : cyear = uploadform("cyear")
	dim cmonth : cmonth = cint(uploadform("cmonth"))
	dim objrs
	dim sql
	dim idx																									' ��ü �麰 �Ϸù�ȣ �ڵ�
	dim tmpCyear
	dim tmpCmonth


	if region = "" then region = null
	if locate = "" then locate = null
	if map = "" then
		map = null														' �൵�������� �ʴ´�.
	else
		tmp = uploadform("txtmap").save(, false)			' �൵ ���Ͽ� ������ ������ �����ϸ� ���ο� ���ϸ����� �����Ѵ�.
		map = right(tmp, len(tmp)-InStrRev(tmp, "\"))	' ���ο� ���� �� ����
	end if
	if side = "" then side = null
	if unitprice = "" then unitprice = 0 else unitprice = replace(unitprice, ",","")
	if quality = "" then quality = null
	if monthprice = "" then monthprice = 0 else monthprice = replace(monthprice, ",","")
	if expense = "" then expense = 0 else expense = replace(expense, ",","")
	if thema = "" then thema = null

	contidx = uploadform("contidx")								'����ȣ

	' ****************************************************************
	' ********** ��� �������� �����ϰ� �������� �����´�.

	sql = "select startdate, enddate from dbo.wb_contact_mst where contidx = " & contidx

	call get_recordset(objrs, sql)

		startdate = objrs("startdate")
		enddate = objrs("enddate")

	objrs.close

	' ****************************************************************
	' ********** ��� ��ü�� ����Ѵ�.

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

	sidx = objrs("sidx")														' ��� ��ü�� �Ϸù�ȣ(��ະ ��ü��Ϲ�ȣ)

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
	' ********** ��� ��ü�� ����� �� ������� ������ ����Ѵ�.

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
	window.opener.location.href="/od/outdoor/pop_contact_view.asp?contidx=<%=contidx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
	this.close();
//-->
</script>