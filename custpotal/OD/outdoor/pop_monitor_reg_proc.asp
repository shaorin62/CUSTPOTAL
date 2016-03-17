<!--#include virtual="/inc/getdbcon.asp"-->
<!--#include virtual="/inc/func.asp"-->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>
<!--#include virtual="/inc/head.asp"-->			<% '초기 설정 페이지(에러 메세지 미출력) %>

<%
	dim uploadform : set uploadform = server.createobject("DEXT.FileUpload")
	uploadform.defaultpath = "C:\pds\monitor"

	dim contidx : contidx = uploadform("contidx")
	dim acceptdate : acceptdate = uploadform("txtacceptdate")
	dim acceptweek : acceptweek = uploadform("selweek")
	dim status : status = uploadform("rdostatus")
	dim nextacceptdate : nextacceptdate = uploadform("txtnextacceptdate")
	dim file : file = uploadform("txtfile")
	dim comment : comment = uploadform("txtcomment")
	dim userid : userid = uploadform("txtuserid")

	dim filecount : filecount = uploadform("txtfile").count

	dim intLoop , temp, flag

	' 첨부파일에 등록가능 여부 판단
	Dim strFileChk1

	for intLoop = 1 to filecount
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
	sql = "select pidx, contidx, acceptdate, status, acceptweek, nextacceptdate, comment, cuser, cdate, uuser, udate from dbo.wb_contact_monitor_mst where contidx = " & contidx
	call set_recordset(objrs, sql)

	objrs.addnew
		objrs.fields("contidx").value = contidx
		objrs.fields("acceptdate").value = acceptdate
		objrs.fields("status").value = status
		objrs.fields("acceptweek").value = acceptweek
		objrs.fields("nextacceptdate").value = nextacceptdate
		objrs.fields("comment").value = comment
		objrs.fields("cuser").value = userid
		objrs.fields("cdate").value = date
		objrs.fields("uuser").value = userid
		objrs.fields("udate").value = date
	objrs.update
	dim pidx : pidx = objrs.fields("pidx").value
	objrs.close

	sql = "select didx, pidx, typical, filename from dbo.WB_CONTACT_MONITOR_DTL where pidx = " & pidx
	call set_recordset(objrs, sql)

	
	flag = 1
	for intLoop = 1 to filecount
			if uploadform("txtfile")(intLoop) <> "" then
				objrs.addnew
				objrs.fields("pidx").value = pidx
				objrs.fields("typical").value = flag
				temp = uploadform("txtfile")(intLoop).Save (, false)
				objrs.fields("filename").value = uploadform("txtfile")(intLoop).filename
				objrs.update
				if flag then flag = 0
			end if
	next
	objrs.close
%>
<script language="JavaScript">
<!--
	window.opener.location.reload();
	this.close();
//-->
</script>