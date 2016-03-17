<!--#Include virtual="/inc/getdbcon.asp"-->
<!--#Include virtual="/inc/func.asp"-->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>
<!--#include virtual="/inc/head.asp"-->				<% '초기 설정 페이지(에러 메세지 미출력) %>

<%
	dim uploadform : set uploadform = server.createobject("DEXT.FileUpload")
	uploadform.defaultPath = "C:\pds\file"


	dim ridx : ridx = uploadform("ridx")
	dim midx : midx = uploadform("midx")
	dim title : title = uploadform("txttitle")
	dim content : content = uploadform("txtcontents")
	dim mail : mail = uploadform("txtmail")
	dim tomail : tomail = uploadform("txttomail")
	dim attachfile : attachfile = uploadform("txtfile")
	dim userid : userid = uploadform("userid")
	dim gotopage : gotopage = uploadform("gotopage")
	dim searchstring : searchstring = uploadform("searchstring")
	dim password : password = uploadform("txtpassword")
	dim highcategory : highcategory = uploadform("cmbhighcategory")
	dim category : category = uploadform("cmbcategory")
	dim custcode : custcode = uploadform("cmbcustcode")
	dim cyear : cyear = uploadform("cyear")
	dim cmonth : cmonth = uploadform("cmonth")
	dim filename
	Dim atag
	dim intLoop
	dim idx


	' 첨부파일에 등록가능 여부 판단
	Dim strFileChk1

	for intLoop = 1 to uploadform("txtfile").count
		if uploadform("txtfile")(intLoop) <> "" Then
			strFileChk1 = Check_Ext(uploadform("txtfile")(intLoop),"DOC,PPT,PPTX,XLS,XLSX,TXT,JPG,GIF,PNG,PDF,AVI,SMI,WMV,MPEG,MPG,ASF,MKV,MP4,TP,TS,MOV,SKM,K3G,FLV,ZIP")
			If strFileChk1  = "error" Then
				Response.write "<script>"
				Response.write "alert('등록할 수 없는 파일입니다.\n\n 파일(DOC,PPT,PPTX,XLS,XLSX,TXT,JPG,GIF,PNG,PDF,AVI,SMI,WMV,MPEG,MPG,ASF,MKV,MP4,TP,TS,MOV,SKM,K3G,FLV,ZIP)만 등록하십시오.');"
				Response.write " this.close();"
				Response.write "</script>"
				Response.End
			End if
		end if
	next

	if mail = "" then mail = null
	if tomail = "" then tomail = null
	if attachfile = "" then attachfile = null

	dim objrs,objrs2, sql
	sql = "select ridx, title, contents, mail, tomail, midx, password, cuser, cdate, uuser, udate, highcategory, category, custcode, cyear, cmonth  from dbo.wb_report where ridx ="&ridx

	call set_recordset(objrs, sql)
	objrs.fields("title").value =clearXSS( title, atag)
	objrs.fields("contents").value = clearXSS(replace(content, "'", "''"), atag)
	objrs.fields("mail").value = clearXSS(mail, atag)
	objrs.fields("tomail").value = clearXSS(tomail, atag)
	objrs.fields("password").value = password
	objrs.fields("uuser").value = userid
	objrs.fields("udate").value = Date
	objrs.fields("highcategory").value = highcategory
	objrs.fields("category").value = category
	objrs.fields("custcode").value = custcode
	objrs.fields("cyear").value = cyear
	objrs.fields("cmonth").value = cmonth

	objrs.update

	if not isnull(tomail) then
		'call getSendMail(mail, tomail, title, content)
	end if

	objrs.close

'
'	sql = "select ridx, attachfile from dbo.wb_Report_pds where ridx = " & ridx
'
'	call set_recordset(objrs, sql)
'
'	for intLoop = 1 to uploadform("txtfile").count
'		if uploadform("txtfile")(intLoop) <> "" then
'			objrs.addnew
'				objrs("ridx") = ridx
'				filename = uploadform("txtfile")(intLoop).Save("C:\pds\file", false)
'				objrs.fields("attachfile").value = replace(filename, uploadform.defaultPath&"\", "")
'			objrs.update
'		end if
'	next
'	objrs.close
'	Set objrs = Nothing


	sql = "select idx, ridx, attachfile from dbo.wb_Report_pds"
	call set_recordset(objrs, sql)

	for intLoop = 1 to uploadform("txtfile").count
		if uploadform("txtfile")(intLoop) <> "" then

			sql  = "select isnull(max(idx),0)+1 idx from dbo.wb_Report_pds where ridx=" & ridx

			call set_recordset(objrs2, sql)
			idx = objrs2("idx")
			objrs2.close

			objrs.addnew
				objrs("idx") = idx
				objrs("ridx") = ridx
				filename = uploadform("txtfile")(intLoop).Save(, false)
				objrs.fields("attachfile").value = replace(filename, uploadform.defaultPath&"\", "")
			objrs.update
		end if
	next

	objrs.close
	Set objrs = Nothing



%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	//window.opener.location.href="list.asp?gotopage=<%=gotopage%>&searchstring=<%=searchstring%>&midx=<%=midx%>";
	//location.href = "pop_report_view.asp?ridx=<%=ridx%>&midx=<%=midx%>&flag=T&txtpassword=<%=password%>";
	window.opener.document.location.href = window.opener.document.URL;
	this.close();
//-->
</SCRIPT>