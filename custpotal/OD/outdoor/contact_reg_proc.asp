<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
'	Dim item
'	For Each item In request.Form
'		response.write item & " : " & request.Form(item) & "<br>"
'	Next
'	response.end
	dim cyear : cyear= request("cyear")
	dim cmonth : cmonth = request("cmonth")
	dim tcustcode : tcustcode = request("custcode")
	dim title : title = request("txttitle")
	dim startdate : startdate = request("txtstartdate")
	dim firstdate : firstdate = request("txtfirstdate")
	dim enddate : enddate = request("txtenddate")
	dim custcode : custcode = request("txtdeptcode")
	dim comment : comment = request("txtcomment")
	dim mediummemo : mediummemo = request("txtmediummemo")
	dim regionmemo : regionmemo = request("txtregionmemo")

	if comment =  "" then comment = null
	if mediummemo = "" then mediummemo = null
	if regionmemo = "" then regionmemo = null

	dim objrs, sql
	sql = "select  contidx, custcode, title, firstdate, startdate, enddate,  regionmemo, mediummemo, comment,   cancel, canceldate, cuser, cdate, uuser, udate from dbo.wb_contact_mst where custcode = '" & custcode & "' "

	call set_recordset(objrs, sql)


	objrs.addnew
	objrs("custcode") = custcode
	objrs("title") = title
	objrs("firstdate") = firstdate
	objrs("startdate") = startdate
	objrs("enddate") = enddate
	objrs("mediummemo") = mediummemo
	objrs("regionmemo") = regionmemo
	objrs("comment") = comment
	objrs("cancel") = 0
	objrs("canceldate") = enddate
	objrs("cuser") = request.cookies("userid")
	objrs("cdate") = date
	objrs("uuser") = request.cookies("userid")
	objrs("udate") = date
	objrs.update

	dim contidx : contidx = objrs.fields("contidx").value

	objrs.close
	set objrs = nothing
%>
<script language="JavaScript">
<!--
	window.opener.location.href = "contact_list.asp?selcustcode=<%=tcustcode%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
	this.close();
//-->
</script>