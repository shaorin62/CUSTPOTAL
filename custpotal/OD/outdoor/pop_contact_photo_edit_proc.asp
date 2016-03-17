<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim uploadform : set uploadform = server.createobject("DEXT.FileUpload")
	uploadform.defaultpath = "C:\pds\media"

	dim contidx : contidx = uploadform("contidx")
	dim sidx : sidx = uploadform("sidx")
	dim cyear : cyear = uploadform("cyear")
	dim cmonth : cmonth = uploadform("cmonth")
	dim photo_1 : photo_1 = uploadform("txtphoto_1")
	dim photo_2 : photo_2 = uploadform("txtphoto_2")
	dim photo_3 : photo_3 = uploadform("txtphoto_3")
	dim photo_4 : photo_4 = uploadform("txtphoto_4")
	dim photo1 : photo1 = uploadform("txtphoto1")
	dim photo2 : photo2 = uploadform("txtphoto2")
	dim photo3 : photo3 = uploadform("txtphoto3")
	dim photo4 : photo4 = uploadform("txtphoto4")
	dim fso : set fso = server.createobject("scripting.filesystemobject")

	dim objrs, sql, attachFile, file
	sql = "select photo_1, photo_2, photo_3, photo_4 from dbo.wb_contact_md_dtl where contidx = " & contidx & " and sidx = " & sidx & " and cyear = " & cyear & " and cmonth = " & cmonth

	call set_recordset(objrs, sql)

	if photo_1 <> "" then
		attachFile = uploadform.defaultpath & "\" & photo1
		if fso.fileexists(attachFile) then	fso.deletefile(attachFile)
		file = uploadform("txtphoto_1").save( ,false)
		objrs.fields("photo_1").value =  uploadform("txtphoto_1").filename
	end if
	if photo_2 <> "" then
		attachFile = uploadform.defaultpath & "\" & photo2
		if fso.fileexists(attachFile) then	fso.deletefile(attachFile)
		file = uploadform("txtphoto_2").save( ,false)
		objrs.fields("photo_2").value =  uploadform("txtphoto_2").filename
	end if
	if photo_3 <> "" then
		attachFile = uploadform.defaultpath & "\" & photo3
		if fso.fileexists(attachFile) then	fso.deletefile(attachFile)
		file = uploadform("txtphoto_3").save( ,false)
		objrs.fields("photo_3").value =  uploadform("txtphoto_3").filename
	end if
	if photo_4 <> "" then
		attachFile = uploadform.defaultpath & "\" & photo4
		if fso.fileexists(attachFile) then	fso.deletefile(attachFile)
		file = uploadform("txtphoto_4").save( ,false)
		objrs.fields("photo_4").value =  uploadform("txtphoto_4").filename
	end if
	set fso = nothing
	objrs.update
	objrs.close
	set objrs = nothing
%>
<script language="JavaScript">
<!--
	window.opener.location.replace("/od/outdoor/pop_contact_view.asp?contidx=<%=contidx%>&sidx=<%=sidx%>");
	this.close();
//-->
</script>