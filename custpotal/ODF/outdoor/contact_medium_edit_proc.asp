<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
'	dim item
'	for each item in request.form
'		response.write item & " : "& request.form(item) & "<br>"
'	next
'	response.end

	dim sidx : sidx = request("sidx")
	dim cmonth : cmonth = request("cmonth")
	dim unitprice : unitprice = request("txtunitprice")
	dim monthprice : monthprice = request("txtmonthprice")
	dim contidx : contidx = request("contidx")
	dim locate : locate = request("txtlocate")
	dim cyear : cyear = request("cyear")
	dim quality : quality = request("selquality")
	dim title : title = request("txttitle")
	dim categoryidx : categoryidx = request("txtcategoryidx")
	dim expense : expense = request("txtexpense")
	dim side : side = request("selside")
	dim qty : qty = request("txtqty")
	dim jobidx : jobidx = request("selsubject")
	dim unit : unit = request("txtunit")
	dim standard : standard = request("txtstandard")
	dim trust : trust = request("rdotrust")
	dim custcode3 : custcode3 = request("selcustcode")
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
	sql = "select contidx, sidx, title, locate, categoryidx, side, unit, unitprice, standard, quality, custcode, qty, trust, map, cuser, cdate, uuser, udate from dbo.WB_CONTACT_MD where contidx = "&contidx &" and sidx = " & sidx
	response.write sql
	call set_recordset(objrs, sql)
		objrs.fields("side").value = side
		objrs.fields("quality").value = quality
		objrs.fields("qty").value = qty
		objrs.fields("trust").value = trust
		objrs.update
	objrs.close

	sql = "select contidx, sidx, cyear, cmonth, monthprice, expense, jobidx, photo_1, photo_2, photo_3, photo_4 from dbo.WB_CONTACT_MD_DTL where contidx = " & contidx & " and sidx = " &sidx&" and cyear = "&cyear&" and cmonth = "&cmonth
	response.write sql
	call set_recordset(objrs, sql)
		objrs.fields("sidx").value = sidx
		objrs.fields("jobidx").value = jobidx
		objrs.fields("monthprice").value = monthprice
		objrs.fields("expense").value = expense
	objrs.fields("jobidx").value = jobidx
	objrs.update

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