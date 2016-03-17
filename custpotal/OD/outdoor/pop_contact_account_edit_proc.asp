<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
'	dim item
'	for each item in request.form
'		response.write item & " : " & request.form(item) & "<br>"
'	next
'	response.end
	dim idx : idx = request("idx")
	dim objrs, sql, intLoop
	sql = "select cyear, cmonth, monthprice, expense from dbo.wb_contact_md_dtl_account where idx = " & idx
	call set_recordset(objrs, sql)

	intLoop = 1
	do until objrs.eof
		objrs("monthprice") = request.form("txtmonthprice")(intLoop)
		objrs("expense") = request.form("txtexpense")(intLoop)
		objrs.update
		intLoop = intLoop + 1
	objrs.movenext
	loop

	objrs.close
	set objrs = nothing
%>
<script language="JavaScript">
<!--
	window.opener.location.reload();
	this.close();
//-->
</script>