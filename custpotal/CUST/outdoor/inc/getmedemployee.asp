<%@CODEPAGE=65001%>
<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%

	' parameter
	' medcode : 매체사 코드 (필수)
	Dim atag_ : atag_ = ""
	Dim pmedcode : pmedcode = Trim(Request("medcode"))
	Dim pempid : pempid = Trim(Request("empid"))

	Dim sql : sql = "select empid, empname from wb_med_employee where medcode=? and useflag=1"
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.parameters.append cmd.createparameter("custcode", adChar, adParamInput, 6)
	cmd.commandType = adCmdText
	cmd.commandText = sql
	cmd.parameters("custcode").value = pmedcode
	Dim rs : Set rs = cmd.Execute
	clearparameter(cmd)
	Set cmd = Nothing

	Response.write "<select id='cmbemp' name='cmbemp'  style='width:91px'>"&vbCrLf
	response.write "<option value=''> -- 매체담당 -- </option>"
	Do Until rs.eof
		response.write "<option value='" & rs(0) & "' "
		If pempid = rs(0) Then Response.write "selected"
		response.write ">" & rs(1) & "</option>" & vbCrLf
		rs.movenext
	Loop
	Response.write "</select>"
%>
