<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim tidx : tidx = request("tidx")
	dim validation : validation = request("txtvalidation")
	dim validation_class : validation_class = request("txtvalidationclass")
	dim board : board = request("sel_board")
	dim boardqty : boardqty = request("txt_board")
	if board = "" then board = 0
	if boardqty = "" then boardqty = 0


	dim objrs, sql
	sql = "select tidx, neon, neonprice, pana, panaprice, board, boardprice, led, ledprice, paint, paintprice, validfee, validclass from dbo.wb_validation_tool where tidx = " & tidx
	response.write sql
	call set_recordset(objrs, sql)

	if objrs.eof then
		objrs.addnew
		objrs.fields("tidx").value = tidx
		objrs.fields("neon").value = 0
		objrs.fields("neonprice").value = 0
		objrs.fields("pana").value = 0
		objrs.fields("panaprice").value = 0
		objrs.fields("board").value = boardqty
		objrs.fields("boardprice").value = board
		objrs.fields("led").value = 0
		objrs.fields("ledprice").value = 0
		objrs.fields("paint").value = 0
		objrs.fields("paintprice").value = 0
		objrs.fields("validfee").value = validation
		objrs.fields("validclass").value = validation_class
		objrs.update
	else
		objrs.fields("board").value = boardqty
		objrs.fields("boardprice").value = board
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