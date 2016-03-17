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

	sql = "select cyear, cmonth, monthprice, expense, contactcancel from dbo.wb_contact_md_dtl where contidx="&contidx

	call set_recordset(objrs, sql)

	if not objrs.eof then 
		do until objrs.eof 
			if dateserial(objrs("cyear"), objrs("cmonth"), "01") > dateserial(cyear, cmonth, "01") then 
				objrs.fields("monthprice").value = 0
				objrs.fields("expense").value = 0
				objrs.fields("contactcancel").value = 1
				objrs.update
			end if
		objrs.movenext
		loop
	end if

	objrs.close
	set objrs = nothing
%>
<script language="JavaScript">
<!--
	window.opener.location.reload();
	location.href = "pop_contact_view.asp?contidx=<%=contidx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
//-->
</script>