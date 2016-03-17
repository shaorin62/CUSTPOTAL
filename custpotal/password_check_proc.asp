<!--#include virtual="/inc/getdbcon_first.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim userid : userid = request("userid")
	dim password : password = request("password")

	dim objrs, sql


	sql = "select c.custcode, c.custname, c.highcustcode, a.class, a.password, a.ispwdchange, a.clipinglevel, a.lastchangedate, a.cdate, a.uuser, a.udate from dbo.wb_account a left outer join dbo.sc_cust_dtl c on a.custcode = c.custcode where userid = '" & userid & "'"

	call set_recordset(objrs, sql)

	dim userclass : userclass = objrs("class")

	if not objrs.eof then
		objrs("password") = password
		objrs("ispwdchange") = 1
		objrs("clipinglevel") = 0
		objrs("lastchangedate") = date
		objrs("uuser") = userid
		objrs("udate") = date
		objrs.update

'		session("userid") = userid
'		response.cookies("userid") = userid
'		if not  isnull(custcode) then  response.cookies("custcode") = custcode else response.cookies("custcode") = "옥외 모니터링 업체"
'		If  IsNull(custcode) Then session("custcode") = "옥외 모니터링" Else session("custcode") = custcode
'		if not isnull(custcode2) then  response.cookies("custcode2") = custcode2 else response.cookies("custcode2") = ""
'		If  IsNull(custcode2) Then session("custcode2") = "" Else session("custcode2") = custcode
'		if not isnull(custname2) then  response.cookies("custname") = custname2 else response.cookies("custname") = ""
'		If  IsNull(custname2) Then session("custname") = "" Else session("custname") = custname2
'		response.cookies("class") = userClass
'		response.cookies("LogTime") = Now
'		session("class") = userClass
'		session("logtime") = now
	end if
	objrs.close
	set objrs = nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	var p = "<%=userclass%>";
	switch (p) {
		case "A":
			window.opener.parent.location.href = "/hq/" ;
			break;
		case "C":
		case "G":
			window.opener.parent.location.href = "/cust/" ;
			break;
		case "D":
		case "H":
			window.opener.parent.location.href = "/cust/" ;
			break;
		case "F":
			window.opener.parent.location.href = "/ODF/" ;
			break;
		case "M":
			window.opener.parent.location.href = "/med/" ;
			break;
		case "O":
			//window.opener.parent.location.href = "/od/outdoor/contact_list.asp" ;
			break;
	}
	this.close();
//-->
</SCRIPT>
