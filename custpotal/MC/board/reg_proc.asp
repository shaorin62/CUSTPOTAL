<!--#Include virtual="/inc/getdbcon.asp"-->
<%
	dim uploadform : set uploadform = server.createobject("DEXT.FileUpload")
	uploadform.defaultPath = server.mappath("..")&"\pds\file"

	Dim menuidx : menuidx = uploadform("menuidx")
	Dim subject : subject =  uploadform("txtsubject")
	Dim contents : contents = uploadform("txtcontents")
	Dim c_date : c_date = Date
	Dim c_user : c_user = Request.Cookies("log_user")
	Dim u_date : u_date = Date
	Dim u_user : u_user = Request.cookies("log_user")
	Dim email : email = uploadform("txtemail")
	Dim filename
	dim custcode : custcode = uploadform("selcustcode")
	dim deptcode : deptcode = uploadform("seldeptcode")

	Dim objrs : Set objrs = server.createobject("adodb.recordset")
	objrs.activeconnection = application("connectionstring")
	objrs.cursorlocation = aduseclient
	objrs.cursortype = adopenstatic
	objrs.locktype = adlockoptimistic
	objrs.source = "dbo.WEB_BOARD"
	objrs.open

	If uploadform("txtfile") <> "" Then
		uploadform("txtfile").save
		filename = uploadform("txtfile").filename
	Else
		filename = null
	End If

	If email = "" Then
		email = null
	else
		getSendMail(email)
	End if

	objrs.addnew
	objrs.fields("SUBJECT").value = Replace(subject, "'", Chr(39))
	objrs.fields("CONTENTS").value = Replace(contents, "'", Chr(39))
	objrs.fields("CDATE").value = c_date
	objrs.fields("CUSER").value = c_user
	objrs.fields("UDATE").value = u_date
	objrs.fields("UUSER").value = u_user
	objrs.fields("FILENAME").value = filename
	objrs.fields("EMAIL").value = email
	objrs.fields("MENUIDX").value = menuidx
	objrs.update

	objrs.close
	Set objrs = Nothing

	sub getSendMail(email)
	dim objMail
	Set objMail = Server.CreateObject("CDO.Message")
	' 메일을 보내는 사람의 이메일 주소
	objMail.From = request.cookies("email")
	' 메일을 받는 사람의 이메일주소(여러사람일 경우는 , 표시로 구분)
	objMail.To = email
	'메일 참조
	'objMail.Cc = ""
	'숨은 참조
	'objMail.Bcc = ""
	' 메일 제목
	objMail.Subject = subject
	dim mail_type
	mail_type = "1"

	'HTML 형식으로 보낼건지 결정
	If mail_type = "0" Then 'HTML 형식이면...
		objMail.HTMLBody = contents
	Else
		objMail.TextBody = contents
	End if

	' 메일 보내기 메소드(이부분이 보내는 부분)
	objMail.Send
	'response.write "<script type='text/javascript'> alert('메일이 발송되었습니다.'); </script>"
	Set objMail = Nothing
	end sub
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	location.href = "list.asp??menuidx=<%=menuidx%>&selcustcode=<%=custcode%>&seldeptcode=<%=deptcode%>"
//-->
</SCRIPT>