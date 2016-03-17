<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim tidx : tidx = request("tidx")
	dim validation : validation = request("txtvalidation")
	dim validation_class : validation_class = request("txtvalidationclass")
	dim panaprice : panaprice = request("sel_pana")
	dim panaqty : panaqty = request("txt_pana")
	dim neonprice : neonprice = request("sel_neon")
	dim neonqty : neonqty = request("txt_neon")
	dim contidx : contidx = request("contidx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	if panaprice = "" then panaprice = 0
	if panaqty = "" then panaqty = 0
	if neonprice = "" then neonprice = 0
	if neonqty = "" then neonqty = 0

	dim objrs, sql
	sql = "select tidx, neon, neonprice, pana, panaprice, board, boardprice, led, ledprice, paint, paintprice, validfee, validclass from dbo.wb_validation_tool where tidx = " & tidx
	call set_recordset(objrs, sql)

	if objrs.eof then
		objrs.addnew
		objrs.fields("tidx").value = tidx
		objrs.fields("neon").value = neonqty
		objrs.fields("neonprice").value = neonprice
		objrs.fields("pana").value = panaqty
		objrs.fields("panaprice").value = panaprice
		objrs.fields("board").value = 0
		objrs.fields("boardprice").value = 0
		objrs.fields("led").value = 0
		objrs.fields("ledprice").value = 0
		objrs.fields("paint").value = 0
		objrs.fields("paintprice").value = 0
		objrs.fields("validfee").value = validation
		objrs.fields("validclass").value = validation_class
		objrs.update
	else
		objrs.fields("neon").value = neonqty
		objrs.fields("neonprice").value = neonprice
		objrs.fields("pana").value = panaqty
		objrs.fields("panaprice").value = panaprice
		objrs.fields("validfee").value = validation
		objrs.fields("validclass").value = validation_class
		objrs.update
	end if

	objrs.close
	set objrs = nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	window.opener.location.href="pop_contact_view.asp?contidx=<%=contidx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
	this.close();
//-->
</SCRIPT>