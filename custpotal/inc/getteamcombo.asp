<%@CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
	' parameter
	' custcode : 광고주 코드 (필수)
	' teamcode : 운영팀 코드 -  선택할 운영팀이 없는 경우  null, 코드가 있으면 해당 운영팀을 선택
	Dim custcode : custcode = UCase(Trim(Request("custcode")))
	Dim deptcode : deptcode = UCase(Trim(request("deptcode")))
	Dim teamcode : teamcode = UCase(Trim(Request("teamcode")))

	Dim sql
	If deptcode = "" Then
	sql = "select custcode, custname from sc_cust_dtl where highcustcode = '"&custcode&"' and gbnflag=0 and use_flag=1 order by custname"
	Else
	sql ="select custcode, custname from sc_cust_dtl where clientsubcode = '" & deptcode & "' and use_flag=1 order by custname"
	End If
	'response.write sql
	'response.end

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	Dim rs : Set rs = cmd.Execute
	Set cmd = Nothing

	Response.write "<select id='cmbteamcode' name='cmbteamcode'  style='width:266px'>"&vbCrLf
	response.write "<option value=''> -- 운영팀를 선택하세요 --</option>"
	Do Until rs.eof
		response.write "<option value='" & rs(0) & "' "
		If teamcode = rs(0) Then Response.write "selected"
		response.write ">" & rs(1) & "</option>" & vbCrLf
		rs.movenext
	Loop
	Response.write "</select>"
%>
