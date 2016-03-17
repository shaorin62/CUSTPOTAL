<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	Dim userid : userid = request.Form("userid")
	Dim username : username = request.Form("username")
	Dim password : password = request.Form("txtpassword")
	Dim classflag : classflag = request.Form("rdoclass")
	Dim custcode : custcode = request.Form("txtcustcode")

	if custcode = "" then custcode = null

	dim objrs2, sql2

'	if classflag  = "G" or  classflag = "C" then
'		sql2 = "select top 1 custcode from dbo.sc_cust_dtl where use_flag = 1 and highcustcode ='" &custcode & "'"
'
'		call set_recordset(objrs2, sql2)
'
'		if not objrs2.eof then
'			custcode =  objrs2("custcode")
'		end if
'		objrs2.close
'
'	end if


	Dim objrs, sql

	'CUSTCODE,
	sql = "select USERID,USERNAME, PASSWORD,  CLASS, uuser, udate from dbo.WB_ACCOUNT where userid = '" & userid & "'"

	Call set_recordset(objrs, sql)

	If password = "" Then password = objrs("password").value

	objrs.fields("USERNAME").value = username
	objrs.fields("PASSWORD").value = password
	'objrs.fields("CUSTCODE").value = custcode
	objrs.fields("CLASS").value = classflag
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