<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%

	dim contidx : contidx = request("contidx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")

	dim objrs, sql
	sql = "select canceldate, uuser, udate from dbo.wb_contact_mst where contidx = " & contidx

	call set_recordset(objrs, sql)

	objrs.fields("canceldate").value = dateserial(cyear, cmonth, "01")
	objrs.fields("uuser").value = request.cookies("userid")
	objrs.fields("udate").value = date

	objrs.update
	objrs.close

	set objrs = nothing
%>
<script language="JavaScript">
<!--
	window.opener.location.reload();
	this.close();
	//location.href = "pop_contact_view.asp?contidx=<%=contidx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
//-->
</script>