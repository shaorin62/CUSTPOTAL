<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim fso : set fso = server.createobject("scripting.filesystemobject")
	dim defaultPath : defaultPath = "C:\pds\media"
	dim defaultPath2 : defaultPath2 = "C:\map"
	dim idx : idx = request("idx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	dim attachFile, totalprice, totalqty

	dim objrs, sql
	sql = "select * from dbo.wb_contact_md_dtl_account where  idx="&idx

	call set_recordset(objrs, sql)

	if not objrs.eof then
			attachFile = defaultpath & "\" & objrs("photo_1")
			if fso.fileexists(attachFile) then	fso.deletefile(attachFile) end if
			attachFile = defaultpath & "\" & objrs("photo_2")
			if fso.fileexists(attachFile) then	fso.deletefile(attachFile) end if
			attachFile = defaultpath & "\" & objrs("photo_3")
			if fso.fileexists(attachFile) then	fso.deletefile(attachFile) end if
			attachFile = defaultpath & "\" & objrs("photo_4")
			if fso.fileexists(attachFile) then	fso.deletefile(attachFile) end if

			objrs.delete()

			objrs.update
	end if
	objrs.close


	sql = "select * from dbo.wb_contact_md_dtl where idx = " & idx
	call set_recordset(objrs, sql)
	dim sidx : sidx = objrs("sidx")

	objrs.delete()
	objrs.update

	objrs.close

	dim bln : bln = false' 삭제 후 동일한 매체의 면 정보을 확인 한 후 아무것도 없다면 true
	sql = "select * from dbo.wb_contact_md_dtl where idx = " & idx
	call get_recordset(objrs, sql)

	if objrs.eof then
		bln = true
	end if
	objrs.close

	if bln then
		sql = "select * from dbo.wb_contact_md m where sidx="&sidx
		call set_recordset(objrs, sql)
		dim contidx : contidx = objrs("contidx")
		attachFile = defaultpath2 & "\" & objrs("map")
		if fso.fileexists(attachFile) then	fso.deletefile(attachFile) end if
		objrs.delete()
		objrs.update

		objrs.close
	end if

	set objrs = nothing

%>
<script language="JavaScript">
<!--
	document.location.replace("/od/outdoor/pop_contact_view.asp?contidx=<%=contidx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>");
//-->
</script>