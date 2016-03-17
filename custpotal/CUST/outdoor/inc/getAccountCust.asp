<%@CODEPAGE=65001%>
<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%

	' parameter
	' medcode : 매체사 코드 (필수)
	Dim atag_ : atag_ = ""
	Dim puserid : puserid = Trim(Request("userid"))

	Dim sql : sql = "select userid, custcode from wb_Account_Cust where userid=? "
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.parameters.append cmd.createparameter("userid", adChar, adParamInput, 6)
	cmd.commandType = adCmdText
	cmd.commandText = sql
	cmd.parameters("userid").value = puserid
	Dim rs : Set rs = cmd.Execute
	clearparameter(cmd)
	Set cmd = Nothing

	Response.write "<select id='cmbemp' name='cmbemp'  style='width:91px'>"&vbCrLf
	response.write "<option value=''></option>"
	Do Until rs.eof
		response.write "<option value='" & rs(0) & "' "
		If pempid = rs(0) Then Response.write "selected"
		response.write ">" & rs(1) & "</option>" & vbCrLf
		rs.movenext
	Loop
	Response.write "</select>"
%>
