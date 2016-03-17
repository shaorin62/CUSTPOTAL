<%@CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
	' parameter
	' custcode : 광고주 코드 (필수)
	' teamcode : 운영팀 코드 -  선택할 운영팀이 없는 경우  null, 코드가 있으면 해당 운영팀을 선택
	Dim custcode : custcode = UCase(Trim(Request("custcode")))	
	Dim teamcode : teamcode = UCase(Trim(Request("teamcode")))
	Dim userid : userid = request.cookies("userid")
	



	Dim sql, cmd, rs, rs2
	
	sql = "select count(*) "
	sql = sql & "  from wb_account_tim "
	sql = sql & "  where userid = '" & userid & "' and clientcode ='" & custcode & "' "
	
	Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	Set rs = cmd.Execute
	Set cmd = Nothing
		

	If  rs(0) = "0" Then
		sql = "select custcode, dbo.sc_get_custname_fun(custcode) timname "
		sql = sql & "  from sc_cust_dtl "
		sql = sql & "  where highcustcode ='" & custcode & "' "
	Else
		sql = "select timcode, dbo.sc_get_custname_fun(timcode) timname "
		sql = sql & "  from wb_account_tim "
		sql = sql & "  where userid = '" & userid & "' and clientcode ='" & custcode & "' "
		sql = sql & "  group by timcode ,dbo.sc_get_custname_fun(clientcode)  "
	End If 

	
	Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	Set rs2 = cmd.Execute
	Set cmd = Nothing


	Response.write "<select id='cmbteamcode' name='cmbteamcode'  style='width:266px'>"&vbCrLf
	response.write "<option value=''> -- 운영팀를 선택하세요 --</option>"
	Do Until rs2.eof
		response.write "<option value='" & rs2(0) & "' "
		If teamcode = rs2(0) Then Response.write "selected"
		response.write ">" & rs2(1) & "</option>" & vbCrLf
		rs2.movenext
	Loop
	Response.write "</select>"
%>
