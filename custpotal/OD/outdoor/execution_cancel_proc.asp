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
	dim contidx : contidx = request("contidx")
	dim custcode : cuStCODe = requEst("cuStCoDe")
	if len(cmonth) = 1 then cmonth = "0"&cmonth

		sql = "select IsPerform, performdate, performuser, IsCancel, a.canceldate, a.canceluser from dbo.wb_contact_mst m inner join dbo.wb_contact_md m2 on m.contidx = m2.contidx inner join dbo.wb_contact_md_dtl d on m2.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on d.idx = a.idx  where m.contidx = " & contidx &" and a.cyear = '"&cyear&"' and a.cmonth = '"&cmonth&"' "
		call set_recordset(objrs, sql)

		do until objrs.eof
				objrs("IsPerform") = 0
				objrs("performdate") = null
				objrs("performuser") = null
				objrs("IsCancel") = 1
				objrs("performdate") = date
				objrs("performuser") = request.cookies("userid")
				objrs.update
			objrs.movenext
		Loop
		objrs.close

		set objrs = nothing

%>
<script language="JavaScript">
<!--
	location.href = "execution_list.asp?cyear=<%=cyear%>&cmonth=<%=cmonth%>&selcustcode=<%=custcode%>";
//-->
</script>