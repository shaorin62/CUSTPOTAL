<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim item
	dim objrs, sql
	dim contidx : contidx = request("contidx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	dim cnt : cnt = request("txtcount")
	dim intLoop
	dim sidx, monthprice, expense

	for intLoop=0 To cnt-1
		for each item in request.form	
			if right(item,1) = cstr(intLoop) then 
				if Mid(item,1,Len(item)-1) = "sidx" then sidx = request.form(item) 
				if Mid(item,1,Len(item)-1) = "txtmonthprice" then monthprice = request.form(item) 
				if Mid(item,1,Len(item)-1) = "txtexpense" then expense = request.form(item) 
			end if
		next
		sql = "select monthprice, expense from dbo.wb_contact_md_dtl where contidx="&contidx&" and sidx="&sidx&" and cyear='"&cyear&"' and cmonth='" &cmonth &"' "
		call set_recordset(objrs, sql)
		objrs.fields("monthprice").value = monthprice
		objrs.fields("expense").value = expense
		objrs.update
		objrs.close
	next

%>
<script language="JavaScript">
<!--
	window.opener.location.reload();
	this.close();
//-->
</script>