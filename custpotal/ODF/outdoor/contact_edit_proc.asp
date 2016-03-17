<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
'	Dim item
'	For Each item In request.Form
'		response.write item & " : " & request.Form(item) & "<br>"
'	Next
'	response.end

	dim contidx : contidx = request("contidx")
	dim title : title = request("txttitle")
	dim startdate : startdate = request("txtstartdate")
	dim mediummemo : mediummemo = request("txtmediummemo")
	dim firstdate : firstdate = request("txtfirstdate")
	dim enddate : enddate = request("txtenddate")
	dim custcode : custcode = request("txtdeptcode")
	dim comment : comment = request("txtcomment")
	dim regionmemo : regionmemo = request("txtregionmemo")

	if mediummemo = "" then mediummemo = null
	if regionmemo = "" then regionmemo = null
	if comment = "" then comment = null


	dim objrs, sql
	sql = "select  custcode, title, firstdate, startdate, enddate,  regionmemo, mediummemo, comment,  uuser, udate from dbo.wb_contact_mst where contidx =" &contidx

	call set_recordset(objrs, sql)

	objrs.fields("custcode").value = custcode
	objrs.fields("title").value = title
	objrs.fields("firstdate").value = firstdate
	objrs.fields("startdate").value = startdate
	objrs.fields("enddate").value = enddate
	objrs.fields("mediummemo").value = mediummemo
	objrs.fields("regionmemo").value = regionmemo
	objrs.fields("comment").value = comment
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
//-->
</script>