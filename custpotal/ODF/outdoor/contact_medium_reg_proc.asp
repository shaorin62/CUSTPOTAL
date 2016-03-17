<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
'	dim item
'	for each item in request.form
'		response.write item & " : "& request.form(item) & "<br>"
'	next
'	response.end

	dim mdidx : mdidx = request("mdidx")
	dim sidx : sidx = request("sidx")
	dim unitprice : unitprice = request("txtunitprice")
	dim monthprice : monthprice = request("txtmonthprice")
	dim contidx : contidx = request("contidx")
	dim title : title = request("txttitle")
	dim categoryidx : categoryidx = request("txtcategoryidx")
	dim expense : expense = request("txtexpense")
	dim side : side = request("selside")
	dim qty : qty = request("txtqty")
	dim jobidx : jobidx = request("selsubject")
	dim unit : unit = request("txtunit")
	dim standard : standard = request("txtstandard")
	dim quality : quality = request("selquality")
	dim locate : locate = request("txtlocate")
	dim custcode3 : custcode3 = request("selcustcode")
	dim trust : trust = request("rdotrust")
	dim startdate : startdate = request("startdate")
	dim enddate : enddate = request("enddate")
	dim regionmemo :  regionmemo = request("txtregionmemo")
	dim mediummemo :  mediummemo = request("txtmediummemo")
	dim map : map = request("txtmap")

	if side = "" then side = null
	if unit = "" then unit = null
	if unitprice = "" then unitprice = 0 else unitprice = replace(unitprice, ",","")
	if expense = "" then expense = 0 else expense = replace(expense, ",","")
	if monthprice = "" then monthprice = 0 else monthprice = replace(monthprice, ",","")
	if quality = "" then quality = null
	if qty = "" then qty = 0 else qty = replace(qty, ",","")
	if locate = "" then locate = null
	if jobidx = "" then jobidx = null


	dim objrs, sql
	sql = "select contidx, sidx, mdidx, title, locate, categoryidx, side, unit, unitprice, standard, quality, custcode, qty, trust, map, cuser, cdate, uuser, udate from dbo.WB_CONTACT_MD where contidx = "&contidx &" and sidx = " & sidx
	call set_recordset(objrs, sql)
	objrs.addnew			'입력할 데이터가 없는 경우에는 신규입력
		objrs.fields("contidx").value = contidx
		objrs.fields("mdidx").value = mdidx
		objrs.fields("sidx").value = sidx
		objrs.fields("title").value = title
		objrs.fields("locate").value = locate
		objrs.fields("categoryidx").value = categoryidx
		objrs.fields("side").value = side
		objrs.fields("unit").value = unit
		objrs.fields("unitprice").value = unitprice
		objrs.fields("standard").value = standard
		objrs.fields("quality").value = quality
		objrs.fields("custcode").value = custcode3
		objrs.fields("qty").value = qty
		objrs.fields("trust").value = trust
		objrs.fields("map").value = map
		objrs.fields("cuser").value = request.cookies("userid")
		objrs.fields("cdate").value = date
		objrs.fields("uuser").value = request.cookies("userid")
		objrs.fields("udate").value = date
		objrs.update
	objrs.close

	sql = "select contidx, sidx, cyear, cmonth, monthprice, expense, jobidx, photo_1, photo_2, photo_3, photo_4, perform from dbo.WB_CONTACT_MD_DTL where contidx = " & contidx & " and sidx = " &sidx
	call set_recordset(objrs, sql)

	dim intLoop, lastmonth, photo_1, photo_2, photo_3, photo_4
	lastmonth = DateDiff("m", startdate, enddate)

	if not objrs.eof then									'저장되어 있는 데이터가 없다면 신규로 입력
		do until objrs.eof
			objrs.delete()
			objrs.movenext
		loop
	end if

	for intLoop = 0 to lastmonth
		objrs.addnew
		objrs.fields("contidx").value = contidx
		objrs.fields("sidx").value = sidx
		objrs.fields("cyear").value = YEAR(startdate)
		objrs.fields("cmonth").value = MONTH(startdate)
		objrs.fields("jobidx").value = jobidx
		objrs.fields("photo_1").value = null
		objrs.fields("photo_2").value = null
		objrs.fields("photo_3").value = null
		objrs.fields("photo_4").value = null
		objrs.fields("perform").value = 0

		if intLoop = 0 then					' 계약 시작일자가 15일 이전인경우에는 청구금액을 입력하지 않는다.
			if day(startdate) > 1 and day(startdate) < 16 then
				objrs.fields("monthprice").value = 0
				objrs.fields("expense").value = 0
			else
				objrs.fields("monthprice").value = monthprice
				objrs.fields("expense").value = expense
			end if
		elseif intLoop = lastmonth then					' 계약 종료일자가 15일 이상인경우에는 청구금액을 입력하지 않는다.
			if day(enddate) <= 15 then
				objrs.fields("monthprice").value = 0
				objrs.fields("expense").value = 0
			else
				objrs.fields("monthprice").value = monthprice
				objrs.fields("expense").value = expense
			end if
		else
			objrs.fields("monthprice").value = monthprice
			objrs.fields("expense").value = expense
		end if

		objrs.update
		startdate = DateAdd("m", 1, startdate)
	next
	objrs.close

	sql = "select contidx, sidx, monthprice, expense, jobidx from dbo.wb_contact_tmp where contidx="&contidx&" and sidx="&sidx
	call set_recordset(objrs, sql)

	objrs.addnew
	objrs.fields("contidx").value = contidx
	objrs.fields("sidx").value = sidx
	objrs.fields("monthprice").value = monthprice
	objrs.fields("expense").value = expense
	objrs.fields("jobidx").value = jobidx
	objrs.update
	set objrs = nothing
%>
<script language="JavaScript">
<!--
	window.opener.location.reload();
	this.close();
//-->
</script>