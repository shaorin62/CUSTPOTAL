<!--#include virtual="/inc/getdbcon.asp" -->
<%
	Dim account : account = request.Form("txtaccount")
	Dim password : password = request.Form("txtpassword")
	Dim authority : authority = request.Form("rdoauthority")
	Dim deptcode : deptcode = request.Form("txtdeptcode")
	Dim custcode : custcode = request.Form("txtcustcode")
	dim empno : empno = request.form("txtempno")

	If deptcode <> "" Then 
		custcode = deptcode
	End If
	
	Dim objrs : Set objrs = server.CreateObject("adodb.recordset")
	objrs.activeconnection = application("connectionstring")
	objrs.cursorlocation = aduseclient
	objrs.cursortype = adopenstatic
	objrs.locktype = adlockoptimistic
	objrs.source = "dbo.WEB_ACCOUNT"
	objrs.open 

	objrs.addnew
	objrs.fields("USERID").value = account
	objrs.fields("PASSWORD").value = password
	objrs.fields("CUSTCODE").value = custcode
	objrs.fields("CLASS").value = authority
	objrs.fields("ISUSE").value = "Y"
	objrs.fields("CUSER").value = Request.Cookies("userid")
	objrs.fields("CDATE").value = date
	objrs.fields("EMPNO").value = empno
	objrs.update

	objrs.close
	Set objrs = Nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	location.href = "acc_list.asp";
//-->
</SCRIPT>