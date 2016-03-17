<%@CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
	' parameter
	' custcode : 광고주 코드 (필수)
	' teamcode : 운영팀 코드 -  선택할 운영팀이 없는 경우  null, 코드가 있으면 해당 운영팀을 선택
	Dim deptcode : deptcode = UCase(Trim(Request("deptcode")))
	Dim teamcode : teamcode = UCase(Trim(Request("teamcode")))
	Dim custcode : custcode = UCase(Trim(request("custcode")))

	sql = "select distinct a.custcode, a.custname from sc_cust_dtl a inner join sc_cust_dtl b on a.custcode=b.clientsubcode where b.highcustcode = '"&custcode&"' and  a.medflag = 'a' and b.use_flag=1"

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	Dim rs : Set rs = cmd.Execute 
	Set cmd = Nothing
	
	Response.write "<select id='cmbdeptcode' name='cmbdeptcode'  style='width:266px'>"&vbCrLf
	response.write "<option value=''> -- 사업부를 선택하세요 -- </option>"&vbCrLf
	Do Until rs.eof 
		response.write "<option value='" & rs(0) & "' "
		If deptcode = rs(0) Then Response.write "selected"
		response.write ">" & rs(1) & "</option>" & vbCrLf
		rs.movenext
	Loop
	Response.write "</select>"
%>
