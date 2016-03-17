<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
'	dim item
'	for each item in request.form
'		response.write item  & " : " & request.form(item) & "<br>"
'	next
'	response.end
	dim objrs, sql, intLoop
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	dim custcode : custcode = request("selcustcode")
	if len(cmonth) = 1 then cmonth = "0"&cmonth

	for intLoop = 1 to request.form("chkitem").count
		sql = "select IsPerform, performdate, performuser from dbo.wb_contact_mst m inner join dbo.wb_contact_md m2 on m.contidx = m2.contidx inner join dbo.wb_contact_md_dtl d on m2.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on d.idx = a.idx  where m.contidx = " & request.form("chkitem")(intLoop) &" and a.cyear = '"&cyear&"' and a.cmonth = '"&cmonth&"' "
		call set_recordset(objrs, sql)

		do until objrs.eof
				objrs.fields("IsPerform").value = 1
				objrs.fields("performdate").value = date
				objrs.fields("performuser").value = request.cookies("userId")
			objrs.movenext
		Loop
		objrs.close
	next

%>
<script language="JavaScript">
<!--
	location.href = "execution_list.asp?cyear=<%=cyear%>&cmonth=<%=cmonth%>&selcustcode=<%=custcode%>";
//-->
</script>