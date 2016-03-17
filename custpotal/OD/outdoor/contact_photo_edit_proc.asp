<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>
<%
	dim uploadform : set uploadform = server.createobject("DEXT.FileUpload")
	uploadform.defaultpath = "C:\pds\media"

	dim contidx : contidx = uploadform("contidx")
	dim idx : idx = uploadform("idx")
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

	' 첨부파일에 등록가능 여부 판단
	Dim strFileChk1, strFileChk2, strFileChk3, strFileChk4
	
	If photo_1  = ""  Then
		photo_1 =  Null
	Else
		strFileChk1 = Check_Ext(photo_1,"JPG,GIF,PNG")

		If strFileChk1  = "error" Then
			Response.write "<script>"
			Response.write "alert('등록할 수 없는 파일입니다.\n\n이미지 파일(JPG,GIF,PNG)만 등록하십시오.');"
			Response.write " this.close();"
			Response.write "</script>"
			Response.End
		End if
	End If
	
	If photo_2  = ""  Then
		photo_2 =  Null
	Else
		strFileChk2 = Check_Ext(photo_2,"JPG,GIF,PNG")

		If strFileChk2  = "error" Then
			Response.write "<script>"
			Response.write "alert('등록할 수 없는 파일입니다.\n\n이미지 파일(JPG,GIF,PNG)만 등록하십시오.');"
			Response.write " this.close();"
			Response.write "</script>"
			Response.End
		End if
	End If
	
	If photo_3  = ""  Then
		photo_3 =  Null
	Else
		strFileChk3 = Check_Ext(photo_3,"JPG,GIF,PNG")

		If strFileChk3  = "error" Then
			Response.write "<script>"
			Response.write "alert('등록할 수 없는 파일입니다.\n\n이미지 파일(JPG,GIF,PNG)만 등록하십시오.');"
			Response.write " this.close();"
			Response.write "</script>"
			Response.End
		End if
	End If
	
	If photo_4  = ""  Then
		photo_4 =  Null
	Else
		strFileChk4 = Check_Ext(photo_4,"JPG,GIF,PNG")

		If strFileChk4  = "error" Then
			Response.write "<script>"
			Response.write "alert('등록할 수 없는 파일입니다.\n\n이미지 파일(JPG,GIF,PNG)만 등록하십시오.');"
			Response.write " this.close();"
			Response.write "</script>"
			Response.End
		End if
	End if

'	response.write "contidx : " & contidx & "<br>"
'	response.write "idx : " & idx & "<br>"
'	response.write "cyear : " & cyear & "<br>"
'	response.write "cmonth : " & cmonth & "<br>"
'	response.end


	dim fso : set fso = server.createobject("scripting.filesystemobject")

	dim objrs, sql, attachFile, file
	sql = "select photo_1, photo_2, photo_3, photo_4 from dbo.wb_contact_md_dtl_account where idx = " & idx & " and cyear = " & cyear & " and cmonth = " & cmonth

	call set_recordset(objrs, sql)

	if photo_1 <> "" then
		attachFile = uploadform.defaultpath & "\" & photo1
		if fso.fileexists(attachFile) then	fso.deletefile(attachFile)
		file = uploadform("txtphoto_1").save( ,false)
		file = right(file, len(file)-InStrRev(file, "\"))
		objrs.fields("photo_1").value =  file
	end if
	if photo_2 <> "" then
		attachFile = uploadform.defaultpath & "\" & photo2
		if fso.fileexists(attachFile) then	fso.deletefile(attachFile)
		file = uploadform("txtphoto_2").save( ,false)
		file = right(file, len(file)-InStrRev(file, "\"))
		objrs.fields("photo_2").value =  file
	end if
	if photo_3 <> "" then
		attachFile = uploadform.defaultpath & "\" & photo3
		if fso.fileexists(attachFile) then	fso.deletefile(attachFile)
		file = uploadform("txtphoto_3").save( ,false)
		file = right(file, len(file)-InStrRev(file, "\"))
		objrs.fields("photo_3").value =  file
	end if
	if photo_4 <> "" then
		attachFile = uploadform.defaultpath & "\" & photo4
		if fso.fileexists(attachFile) then	fso.deletefile(attachFile)
		file = uploadform("txtphoto_4").save( ,false)
		file = right(file, len(file)-InStrRev(file, "\"))
		objrs.fields("photo_4").value =  file
	end if
	set fso = nothing
	objrs.update
	objrs.close
	set objrs = nothing
%>
<script language="JavaScript">
<!--
	window.opener.location.replace("/od/outdoor/pop_contact_view.asp?contidx=<%=contidx%>&idx=<%=idx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>");
	this.close();
//-->
</script>