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
	' ������ ������ ����� �̸��� �ּ�
	objMail.From = request.cookies("email")
	' ������ �޴� ����� �̸����ּ�(��������� ���� , ǥ�÷� ����)
	objMail.To = email
	'���� ����
	'objMail.Cc = ""
	'���� ����
	'objMail.Bcc = ""
	' ���� ����
	objMail.Subject = subject
	dim mail_type
	mail_type = "1"

	'HTML �������� �������� ����
	If mail_type = "0" Then 'HTML �����̸�...
		objMail.HTMLBody = contents
	Else
		objMail.TextBody = contents
	End if

	' ���� ������ �޼ҵ�(�̺κ��� ������ �κ�)
	objMail.Send
	'response.write "<script type='text/javascript'> alert('������ �߼۵Ǿ����ϴ�.'); </script>"
	Set objMail = Nothing
	end sub
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	location.href = "list.asp??menuidx=<%=menuidx%>&selcustcode=<%=custcode%>&seldeptcode=<%=deptcode%>"
//-->
</SCRIPT>