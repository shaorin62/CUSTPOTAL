<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim tidx : tidx = request("tidx")
	dim validation : validation = request("txtvalidation")
	dim validation_class : validation_class = request("txtvalidationclass")
	dim led : led = request("sel_led")
	dim ledqty : ledqty = request("txt_led")
	dim contidx : contidx = request("contidx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	if led = "" then led = 0
	if ledqty = "" then ledqty = 0

	dim objrs, sql
	sql = "select tidx, neon, neonprice, pana, panaprice, board, boardprice, led, ledprice, paint, paintprice, validfee, validclass from dbo.wb_validation_tool where tidx = " & tidx
	call set_recordset(objrs, sql)

	if objrs.eof then
		objrs.addnew
		objrs.fields("tidx").value = tidx
		objrs.fields("neon").value = 0
		objrs.fields("neonprice").value = 0
		objrs.fields("pana").value = 0
		objrs.fields("panaprice").value = 0
		objrs.fields("board").value = 0
		objrs.fields("boardprice").value = 0
		objrs.fields("led").value = ledqty
		objrs.fields("ledprice").value = led
		objrs.fields("paint").value = 0
		objrs.fields("paintprice").value = 0
		objrs.fields("validfee").value = validation
		objrs.fields("validclass").value = validation_class
		objrs.update
	else
		objrs.fields("led").value = ledqty
		objrs.fields("ledprice").value = led
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