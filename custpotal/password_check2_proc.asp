<!--#include virtual="/inc/getdbcon_first.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim empid : empid = request("userid")
	dim emppwd : emppwd = request("password")

	dim objrs, sql
	sql = "select c.custname, c.highcustcode, a.class, a.emppwd, a.ispwdchange, a.clipinglevel, a.lastchangedate from dbo.wb_med_employee a left outer join dbo.sc_cust_hdr c on a.medcode = c.highcustcode where medflag='B' and empid = '" & empid & "'"

	call set_recordset(objrs, sql)

	dim userclass : userclass = "M"

	if not objrs.eof then
		sql = "update wb_med_employee set emppwd =?, ispwdchange=?, clipinglevel=?, lastchangedate=? where empid =?"
		dim cmd : set cmd = server.createobject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandType = adCmdText
		cmd.commandText = sql
		cmd.parameters.append cmd.createparameter("emppwd", adVarchar, adParamInput,12)
		cmd.parameters.append cmd.createparameter("ispwdchange", adBoolean, adParamInput)
		cmd.parameters.append cmd.createparameter("clipinglevel", adUnsignedTinyInt, adParamInput)
		cmd.parameters.append cmd.createparameter("lastchangedate", adDBTimeStamp, adParamInput)
		cmd.parameters.append cmd.createparameter("empid", advarchar, adParamInput, 12)
		cmd.parameters("emppwd").value = emppwd
		cmd.parameters("ispwdchange").value = 1
		cmd.parameters("clipinglevel").value = 0
		cmd.parameters("lastchangedate").value = date()
		cmd.parameters("empid").value = empid
		cmd.execute ,, adExecuteNoRecords
		set cmd = nothing


		session("userid") = empid
		session("class") ="m"
		session("custname") = objrs("custname")
		session("logtime") = now

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
