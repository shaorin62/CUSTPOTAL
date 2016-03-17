<!--#Include virtual="/inc/getdbcon.asp"-->
<!--#Include virtual="/inc/func.asp"-->
<%
	dim cidx : cidx = request("cidx")
	dim ridx : ridx = request("ridx")
	dim midx : midx = request("midx")


	dim objrs, sql
	sql = "select cidx, ridx, comment, attachfile, cuser, cdate from dbo.wb_report_comment where cidx ="&cidx & " and ridx =" & ridx

	call set_recordset(objrs, sql)

	dim attachFile : attachFile = "C:\pds\file" & "\"& objrs("attachfile")
	dim fso : set fso = server.createobject("scripting.filesystemobject")
	if fso.fileexists(attachFile) then  fso.deletefile(attachFile)
	set fso = nothing

	objrs.delete()
	objrs.update

	objrs.close
	Set objrs = Nothing

%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	window.opener.location.reload();
	location.href="pop_report_view.asp?ridx=<%=ridx%>&midx=<%=midx%>";
//-->
</SCRIPT>