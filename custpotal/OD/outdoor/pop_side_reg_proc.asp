<!--#include virtual="/inc/getdbcon.asp"-->
<!--#include virtual="/inc/func.asp"-->
<%
	dim item
	for each item in request.form
		response.write item  & " : " & request.form(item) & "<br>"
	next

	dim side : side = request("selside")
	dim standard : standard = request("txtstandard")
	dim quality : quality = request("selquality")
	dim unitprice : unitprice = request("txtunitprice")
	dim mdidx : mdidx = request("mdidx")
	if side = "" then side = null
	if quality = "" then quality = null

	dim objrs, sql
	sql = "select top 1 sidx, mdidx, side, standard, quality, unitprice from dbo.WB_MEDIUM_DTL "
	call set_recordset(objrs, sql)

	objrs.addnew
	objrs.fields("mdidx").value = mdidx
	objrs.fields("side").value = side
	objrs.fields("standard").value = replace(standard, """", chr(34))
	objrs.fields("quality").value = quality
	objrs.fields("unitprice").value = unitprice
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