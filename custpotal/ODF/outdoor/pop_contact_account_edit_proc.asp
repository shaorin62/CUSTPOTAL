<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
'	dim item 
'	for each item in request.form
'		response.write item & " : " & request.form(item) & "<br>"
'	next
'	response.end
	dim contidx : contidx = request("contidx")
	dim sidx : sidx = request("sidx")
	dim objrs, sql, intLoop
	sql = "select contidx, sidx, cyear, cmonth, perform from dbo.wb_contact_md_dtl where contidx = " & contidx 
	call set_recordset(objrs, sql)

	for intLoop = 1 to request.form("txtchk").count
			objrs.filter = "contidx="&contidx&" and sidx="&sidx&" and cyear='" & request.form("txtcyear")(intLoop) &"' and cmonth='" & request.form("txtcmonth")(intLoop) &"'"
			objrs.fields("perform").value = cbool(request.form("txtchk")(intLoop))
			objrs.update
	next

	objrs.close
	set objrs = nothing
%>
<script language="JavaScript">
<!--
	window.opener.location.reload();
	this.close();
//-->
</script>