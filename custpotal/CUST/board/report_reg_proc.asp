<!--#Include virtual="/inc/getdbcon.asp"-->
<!--#Include virtual="/inc/func.asp"-->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '���Ȱ��ö��̺귯�� %>
<!--#include virtual="/inc/head.asp"-->			<% '�ʱ� ���� ������(���� �޼��� �����) %>

<%
	dim uploadform : set uploadform = server.createobject("DEXT.FileUpload")
	uploadform.defaultPath = "c:\pds\file"

	dim midx : midx = uploadform("midx")
	dim title : title = uploadform("txttitle")
	dim content : content = uploadform("txtcontents")
	dim mail : mail = uploadform("txtmail")
	dim tomail : tomail = uploadform("txttomail")
	dim attachfile : attachfile = uploadform("txtfile")
	dim attachfile2 : attachfile2 = uploadform("txtfile2")
	dim attachfile3 : attachfile3 = uploadform("txtfile3")
	dim userid : userid = uploadform("userid")
	dim password : password = uploadform("txtpassword")
	dim filename
	dim intloop
	dim ridx , idx
	Dim atag
	dim highcategory : highcategory = uploadform("cmbhighcategory")
	dim category : category = uploadform("cmbcategory")
	dim custcode : custcode = uploadform("cmbcustcode")
	dim cyear : cyear = uploadform("cyear")
	dim cmonth : cmonth = uploadform("cmonth")

	

'	response.write "midx : "& midx & "<br>"
'	response.write "title : "& title & "<br>"
'	response.write "content : "& content & "<br>"
'	response.write "mail : "& mail & "<br>"
'	response.write "tomail : "& tomail & "<br>"
'	response.write "userid : "& userid & "<br>"
'	response.write "password : "& password & "<br>"
'	response.write "attachfile : "& attachfile & "<br>"
'	response.write "attachfile2 : "& attachfile2 & "<br>"
'	response.write "attachfile3 : "& attachfile3 & "<br>"
'

	' ÷�����Ͽ� ��ϰ��� ���� �Ǵ�
	Dim strFileChk1

	for intLoop = 1 to uploadform("txtfile").count
		if uploadform("txtfile")(intLoop) <> "" Then
			strFileChk1 = Check_Ext(uploadform("txtfile")(intLoop),"DOC,PPT,PPTX,XLS,XLSX,TXT,JPG,GIF,PNG,PDF,AVI,SMI,WMV,MPEG,MPG,ASF,MKV,MP4,TP,TS,MOV,SKM,K3G,FLV,ZIP")
			
			'AVI,SMI,WMV,MPEG,MPG,ASF,MKV,MP4,TP,TS,MOV,SKM,K3G,FLV ������ ���̴� Ȯ���� ����Ʈ
			If strFileChk1  = "error" Then
				Response.write "<script>"
				Response.write "alert('����� �� ���� �����Դϴ�.\n\n ����(DOC,PPT,XLS,XLSX,TXT,JPG,GIF,PNG,PDF,AVI,SMI,WMV,MPEG,MPG,ASF,MKV,MP4,TP,TS,MOV,SKM,K3G,FLV,ZIP)�� ����Ͻʽÿ�.');"
				Response.write " this.close();"
				Response.write "</script>"
				Response.End
			End if

		   If uploadform("txtfile")(intLoop).FileLen > 1073741824 then   '1073741824byte = 1GB
				Response.write "<script>"
				Response.write "alert('1GB���� ū�뷮�� ���ѵ˴ϴ�.');"
				Response.write " this.close();"
				Response.write "</script>"
				Response.End
		   end if

		end if
	next


	if mail = "" then mail = null
	if tomail = "" then tomail = null
	if trim(password) = "" then password = null

	dim objrs, objrs2, sql

	'ridx ����ũ�ؾ���
	sql  = "select isnull(max(ridx),0)+1 ridx from dbo.wb_report"
	call set_recordset(objrs, sql)
	ridx = objrs("ridx")
	objrs.close

	sql = "select ridx, title, contents, mail, tomail, midx, password, cuser, cdate, uuser, udate, highcategory, category, custcode, cyear, cmonth from dbo.wb_report"
	call set_recordset(objrs, sql)

	objrs.addnew
	'atag ����� �±� ����   clearXSS ( '����' , '������±�') db�� < ,  > �Ͱ��� ��ȣ��  &#41;�Ͱ��� ����
	objrs.fields("midx").value = midx
	objrs.fields("ridx").value = ridx
	objrs.fields("title").value =clearXSS( title, atag)
	objrs.fields("contents").value = clearXSS(replace(content, "'", "''"), atag)
	objrs.fields("mail").value = clearXSS(mail, atag)
	objrs.fields("tomail").value = clearXSS(tomail, atag)
	objrs.fields("midx").value = midx
	objrs.fields("password").value = password
	objrs.fields("cuser").value = userid
	objrs.fields("cdate").value = date
	objrs.fields("uuser").value = userid
	objrs.fields("udate").value = Date
	objrs.fields("highcategory").value = highcategory
	objrs.fields("category").value = category
	objrs.fields("custcode").value = custcode
	objrs.fields("cyear").value = cyear
	objrs.fields("cmonth").value = cmonth
	objrs.update
	objrs.close


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

	if not (isnull(tomail) or tomail = "") then
		call getSendMail(mail, tomail, title, content)
	end if

%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	window.opener.document.location.href = window.opener.document.URL;
	//window.opener.scriptFrame.location.href="list.asp?midx=<%=midx%>";
	this.close();
//-->
</SCRIPT>