<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
'	dim item
'	for each item in request.form
'		response.write item  & " : " & request.form(item) & "<br>"
'	next
'	response.end
	dim objrs, sql, intLoop
	dim contidx : contidx = request("contidx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	
	for intLoop = 1 to request.form("chkitem").count
		sql = "select perform, performdate, performuser from dbo.wb_contact_md_dtl where contidx="&request.form("chkitem")(intLoop)&" and cyear='"&cyear&"' and cmonth='"&cmonth&"' "
		call set_recordset(objrs, sql)

		do until objrs.eof 
				objrs.fields("perform").value = 1
				objrs.fields("performdate").value = date
				objrs.fields("performuser").value = request.cookies("userId")
			objrs.movenext
		Loop
		objrs.close
	next


%>
<script language="JavaScript">
<!--
	location.href = "execution_list.asp?cyear=<%=cyear%>&cmonth=<%=cmonth%>";
//-->
</script>