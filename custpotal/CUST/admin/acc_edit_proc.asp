<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%

'	Dim item
'	For Each item In request.Form
'		response.write item &  " :" & request.Form(item) & "<br>"
'	Next
'	response.end
	Dim account : account = request.Form("txtaccount")
	Dim password : password = request.Form("txtpassword")
	Dim class_ : class_ = request.Form("rdoclass")
	Dim custcode : custcode = request.Form("txtcustcode")
	Dim isuse : isuse = request.Form("rdoisuse")

	Dim objrs, sql
	sql = "select USERID, PASSWORD, CUSTCODE, CLASS, ISUSE, uuser, udate from dbo.WB_ACCOUNT where userid = '" & account & "'"

	Call set_recordset(objrs, sql)
	
	If password = "" Then password = objrs("password").value


	objrs.fields("PASSWORD").value = password
	objrs.fields("CUSTCODE").value = custcode
	objrs.fields("CLASS").value = class_
	objrs.fields("ISUSE").value = isuse
	objrs.fields("uuser").value = Request.Cookies("userid")
	objrs.fields("udate").value = date
	objrs.update

	objrs.close
	Set objrs = Nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	window.opener.location.reload();
	this.close();
//-->
</SCRIPT>