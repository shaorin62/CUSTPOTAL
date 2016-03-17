<%
	session.abandon

	Response.cookies("userid") = ""
	response.cookies("custocode") = ""
	response.cookies("custocode2") = ""
	response.cookies("custname") = ""
	response.cookies("class") = ""
	response.cookies("logtime") = ""

%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	location.href = "/";
//-->
</SCRIPT>