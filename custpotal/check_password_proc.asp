<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim userid : userid = request("userid")
	dim password : password = request("password")

	dim objrs, sql
	sql = "select password, pwdudate from dbo.wb_account where userid = '" & userid & "'"
'	response.write sql
'	response.end
	call set_recordset(objrs, sql)

	if not objrs.eof then
		objrs("password").value = password
		objrs("pwdudate").value = date
		objrs.update
	end if

	objrs.close
	set objrs = nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	window.opener.parent.location.href = "/hq/main.asp";
	this.close();
//-->
</SCRIPT>
