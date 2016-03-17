<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim contidx : contidx = request("contidx")
	dim regionmemo : regionmemo = request("txtregionmemo")
	dim mediummemo : mediummemo = request("txtmediummemo")
	dim comment : comment = request("txtcomment")

	dim objrs, sql
	sql = "select regionmemo, mediummemo, comment from dbo.wb_contact_mst where contidx = " & contidx
	call set_recordset(objrs, sql)

	objrs.fields("regionmemo").value = regionmemo
	objrs.fields("mediummemo").value = mediummemo
	objrs.fields("comment").value = comment
	objrs.update

	objrs.close
	set objrs = nothing
%>
<script language="JavaScript">
<!--
	window.opener.location.reload();
	this.close();
//-->
</script>