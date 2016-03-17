<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%


	Dim account : account = request.Form("txtaccount")
	Dim name : name = request.Form("txtname")
	Dim password : password = request.Form("txtpassword")
	Dim custcode : custcode = request.Form("txtcustcode")
	Dim classflag : classflag = request.Form("rdoclass")

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
'

	Dim objrs, sql

	sql = "select userid, username, password, custcode, class, isuse, initpwd, ispwdchange, clipinglevel, lastchangedate, cuser, cdate, uuser, udate from dbo.WB_ACCOUNT where userid = '" & account & "'"

	Call set_recordset(objrs, sql)

	objrs.addnew

		objrs.fields("userid").value = account
		objrs.fields("username").value = name
		objrs.fields("password").value = password
		'objrs.fields("custcode").value = custcode
		objrs.fields("class").value = classflag
		objrs.fields("isuse").value = "Y"
		objrs.fields("initpwd").value = password
		objrs.fields("ispwdchange").value = 0
		objrs.fields("clipinglevel").value = 0
		objrs.fields("lastchangedate").value = null
		objrs.fields("cuser").value = Request.Cookies("userid")
		objrs.fields("cdate").value = date
		objrs.fields("uuser").value = Request.Cookies("userid")
		objrs.fields("udate").value = date
	objrs.update

	objrs.close
	Set objrs = Nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	parent.opener.frmuser.location.href="account_fuserid.asp?strUserid=<%=account%>";
	parent.opener.frmcust.location.href="account_fcust.asp?strUserid=";
	parent.opener.frmtim.location.href="account_ftim.asp?strUserid=";
	this.close();
//-->
</SCRIPT>