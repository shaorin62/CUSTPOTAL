<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim idx : idx = request("idx")
	dim photoIdx : photoIdx = request("photoIdx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	dim note : note = request("txtnote")


	dim objrs, sql
	sql = "select note from dbo.wb_contact_photo_dtl where idx = " & photoIdx
	call set_recordset(objrs, sql)

	if note = "" then note = null

	if not objrs.eof then
		objrs("note") = note
		objrs.update
	end if

	objrs.close
	set objrs = nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	window.opener.location.href = "pop_contact_photo_reg.asp?idx=<%=idx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
	this.close();
//-->
</SCRIPT>