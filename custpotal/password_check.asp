<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim userid : userid = request("userid")
	dim password : password = request("password")

	dim objrs, sql
	sql = "select * from dbo.wb_account where userid = '" & userid & "'"
	call set_recordset(objrs, sql)

	dim clipinglevel : clipinglevel = objrs("clipinglevel")
	objrs.fields("clipinglevel").value = clipinglevel + 1
	objrs.update

	objrs.close

%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	parent.location.href = 'index.htm';
//-->
</SCRIPT>