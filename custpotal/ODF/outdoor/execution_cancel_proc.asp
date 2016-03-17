<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
'	dim item
'	for each item in request.form
'		response.write item  & " : " & request.form(item) & "<br>"
'	next
'	response.end
	dim objrs, sql, intLoop
	dim contidx : contidx = request("txtcontidx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")

		sql = "select perform, canceldate, canceluser from dbo.wb_contact_md_dtl where contidx="&contidx&" and cyear='"&cyear&"' and cmonth='"&cmonth&"' "
		call set_recordset(objrs, sql)

		objrs.fields("perform").value = 0
		objrs.fields("canceldate").value = date
		objrs.fields("canceluser").value = request.cookies("userId")

		objrs.close


%>
<script language="JavaScript">
<!--
	location.href = "execution_list.asp?cyear=<%=cyear%>&cmonth=<%=cmonth%>";
//-->
</script>