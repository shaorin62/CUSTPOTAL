<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>
<!--#include virtual="/inc/head.asp"-->			<% '초기 설정 페이지(에러 메세지 미출력) %>

<%
	dim uploadform : set uploadform = server.createobject("DEXT.FileUpload")
	uploadform.defaultpath = "C:\pds\media"

	dim file : file = uploadform("txtfile")
	dim idx : idx = uploadform("idx")
	dim cyear : cyear = uploadform("cyear")
	dim cmonth : cmonth = uploadform("cmonth")
	dim intLoop , filename

	' 첨부파일에 등록가능 여부 판단
	Dim strFileChk1

	for intLoop = 1 to uploadform("txtfile").count
		if uploadform("txtfile")(intLoop) <> "" Then
			strFileChk1 = Check_Ext(uploadform("txtfile")(intLoop),"JPG,GIF,PNG")
			If strFileChk1  = "error" Then
				Response.write "<script>"
				Response.write "alert('등록할 수 없는 파일입니다.\n\n 이미지 파일(JPG,GIF,PNG)만 등록하십시오.');"
				Response.write " this.close();"
				Response.write "</script>"
				Response.End
			End if
		end if
	next

	dim objrs, sql

	sql = "select idx from dbo.wb_contact_photo_mst where dtlidx = " & idx
	call get_recordset(objrs, sql)

	dim mstidx

	if not objrs.eof then
		mstidx = objrs("idx")
	else
		mstidx = 0
	end if

	sql = "select d.idx, d.mstIdx, d.chk, d.filename, d.note from dbo.wb_contact_photo_dtl d  where d.mstidx = " & mstidx
	call set_recordset(objrs, sql)

	
	for intLoop = 1 to uploadform("txtfile").count
		if uploadform("txtfile")(intLoop) <> "" then
			filename = uploadform("txtfile")(intLoop).Save (, false)
			filename = right(filename, len(filename)-InStrRev(filename, "\"))

			objrs.addnew
			objrs("mstIdx") = mstIdx
			objrs("chk")  = 0
			objrs("filename") = filename
			objrs("note") = uploadform("txtcomment")(intLoop)
			objrs.update
		end if
	next

	objrs.close
	set objrs = nothing
%>
<script language="JavaScript">
<!--
	window.opener.location.href="pop_contact_photo_reg.asp?idx=<%=idx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
	this.close();
//-->
</script>