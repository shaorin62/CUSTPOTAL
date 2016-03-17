<!--#Include virtual="/inc/getdbcon.asp"-->
<!--#Include virtual="/inc/func.asp"-->
<%
	dim ridx : ridx = request("ridx")
	dim objrs, sql
	dim fso : set fso = server.createobject("scripting.filesystemobject")

	sql = "select cidx, ridx, comment, attachfile, cuser, cdate from dbo.wb_report_comment where ridx ="&ridx
	call set_recordset(objrs, sql)

	dim attachFile

	if not objrs.eof then
		do until objrs.eof
			 attachFile = "C:\pds\file" & "\"& objrs("attachfile")
			if fso.fileexists(attachFile) then  fso.deletefile(attachFile)
			objrs.delete()
		objrs.movenext
		Loop
	end if
	objrs.close

'	sql = "select ridx, title, contents, mail, tomail, attachfile, attachfile2, attachfile3, midx, cuser, cdate, uuser, udate from dbo.wb_report where ridx ="&ridx
	sql = "select attachfile from dbo.wb_report_pds where ridx = " & ridx
	call set_recordset(objrs, sql)

	if not objrs.eof then
	do until objrs.eof
	attachFile  = "C:\pds\file" & "\"& objrs("attachfile")
		if fso.fileexists(attachFile) then  	fso.deletefile(attachFile)
		objrs.delete()
		objrs.update
	objrs.movenext
	Loop
	end if

	objrs.close

	sql = "select * from dbo.wb_report where ridx =" & ridx
	call set_recordset(objrs, sql)
	if not objrs.eof then
		objrs.delete
		objrs.update
	end if
	objrs.close
	Set objrs = Nothing
	set fso = nothing

%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	window.opener.location.reload();
	this.close();
//-->
</SCRIPT>