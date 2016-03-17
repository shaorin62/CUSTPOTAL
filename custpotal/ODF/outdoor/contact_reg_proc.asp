<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
'	Dim item
'	For Each item In request.Form
'		response.write item & " : " & request.Form(item) & "<br>"
'	Next
'	response.end

	dim title : title = request("txttitle")
	dim startdate : startdate = request("txtstartdate")
	dim firstdate : firstdate = request("txtfirstdate")
	dim enddate : enddate = request("txtenddate")
	dim custcode : custcode = request("txtdeptcode")
	dim comment : comment = request("txtcomment")
	dim mediummemo : mediummemo = request("txtmediummemo")
	dim regionmemo : regionmemo = request("txtregionmemo")

	dim objrs, sql
	sql = "select top 1 contidx, custcode, title, firstdate, startdate, enddate,  regionmemo, mediummemo, comment, cuser, cdate, uuser, udate from dbo.wb_contact_mst "

	call set_recordset(objrs, sql)

	if comment =  "" then comment = null

	objrs.addnew
	objrs.fields("custcode").value = custcode
	objrs.fields("title").value = title
	objrs.fields("firstdate").value = firstdate
	objrs.fields("startdate").value = startdate
	objrs.fields("enddate").value = enddate
	objrs.fields("comment").value = comment
	objrs.fields("mediummemo").value = mediummemo
	objrs.fields("regionmemo").value = regionmemo
	objrs.fields("cuser").value = request.cookies("userid")
	objrs.fields("cdate").value = date
	objrs.fields("uuser").value = request.cookies("userid")
	objrs.fields("udate").value = date
	objrs.update

	dim contidx : contidx = objrs.fields("contidx").value
	
	objrs.close
	set objrs = nothing
%>
<script language="JavaScript">
<!--
	window.opener.location.reload();
	this.close();
//-->
</script>