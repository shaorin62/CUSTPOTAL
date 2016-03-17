<!--#Include virtual="/inc/getdbcon.asp"-->
<!--#Include virtual="/inc/func.asp"-->
<%

	dim idx : idx = request("idx")
	dim ridx : ridx = request("ridx")
	dim midx : midx = request("midx")

	dim objrs, sql
	dim fso : set fso = server.createobject("scripting.filesystemobject")

	sql = "select ridx, attachfile  from dbo.wb_report_pds where ridx = " & ridx & " and  idx = " & idx


	call set_recordset(objrs, sql)

	dim attachFile

	if not objrs.eof then
		attachFile = "C:\pds\file" & "\"& objrs("attachfile")
		 if fso.fileexists(attachFile) then  fso.deletefile(attachFile)

		objrs.delete
		objrs.update
	end if

	Set objrs = Nothing
	set fso = nothing

%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	location.href="pop_report_edit.asp?ridx=<%=ridx%>&midx=<%=midx%>" ;
//-->
</SCRIPT>