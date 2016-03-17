<!--#Include virtual="/inc/getdbcon.asp"-->
<!--#Include virtual="/inc/func.asp"-->
<%
	dim uploadform : set uploadform = server.createobject("DEXT.FileUpload")
	uploadform.defaultPath = "C:\pds\file"


	dim ridx : ridx = uploadform("ridx")
	dim comment : comment = uploadform("txtcomment")
	dim attachfile : attachfile = uploadform("txtfile")
	dim userid : userid = uploadform("userid")
	dim filename

	if attachfile = "" then attachfile = null

	dim objrs, sql

	dim cidx
	'cidx 유니크해야함
	sql  = "select isnull(max(cidx),0)+1 cidx from dbo.wb_report_comment"
	call set_recordset(objrs, sql)
	cidx = objrs("cidx")
	objrs.close


	sql = "select cidx, ridx, comment, attachfile, cuser, cdate from dbo.wb_report_comment where ridx ="&ridx
	call set_recordset(objrs, sql)

	objrs.addnew
	objrs.fields("cidx").value = cidx
	objrs.fields("ridx").value = ridx
	objrs.fields("comment").value = comment
	if uploadform("txtfile") = "" then
		objrs.fields("attachfile").value = null
	else
		attachFile = "C:\pds\file" & "\"& objrs("attachfile")
		dim fso : set fso = server.createobject("scripting.filesystemobject")
		if fso.fileexists(attachFile) then
			fso.deletefile(attachFile)
		end if
		filename = uploadform("txtfile").Save(, false)
		objrs.fields("attachfile").value = replace(filename, uploadform.defaultPath&"\", "")
	end if
	objrs.fields("cuser").value = request.cookies("userid")
	objrs.fields("cdate").value = date
	objrs.update

	objrs.close
	Set objrs = Nothing

%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	window.opener.opener.location.reload();
	window.opener.location.reload();
	this.close();
//-->
</SCRIPT>