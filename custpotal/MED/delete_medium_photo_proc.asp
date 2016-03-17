<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim idx : idx = request("idx")
	dim photoIdx : photoIdx = request("photoIdx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	dim filename : filename = request("filename")
	dim fso : set fso = server.createobject("scripting.filesystemobject")
	dim defaultpath : defaultpath = "C:\pds\media"

	dim objrs, sql, objrs2
	sql = "select * from dbo.wb_contact_photo_dtl where idx = " & photoIdx
	call set_recordset(objrs, sql)

	dim mstIdx : mstIdx = objrs("mstIdx")

	if not objrs.eof then
		objrs.delete
		objrs.update


	end if

	dim file : file = "C:\pds\media\"&filename
	if fso.fileexists(file) then	fso.deletefile(file)

	objrs.close

	sql = "select * from dbo.wb_contact_photo_dtl where mstIdx = " & mstIdx
	call get_recordset(objrs, sql)

	if objrs.eof then
		sql = "select * from dbo.wb_contact_photo_mst where idx = " & mstIdx
		call set_recordset(objrs2, sql)

		objrs2.delete
		objrs2.update

		objrs2.close
		set objrs2 = nothing
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