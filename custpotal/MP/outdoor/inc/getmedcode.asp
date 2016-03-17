<%@CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
	On Error Resume Next

	Dim pmedname : pmedname = request("medname")
	Dim pmedcode : pmedcode  = request("medcode")
	Dim sql : sql = "select highcustcode, custname from sc_cust_hdr where custname like '%" & pmedname & "' order by custname"
	'response.write sql

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	Dim rs : Set rs = cmd.Execute
	Set cmd = Nothing

	response.write "<select id='cmbmed' name='cmbmed' >"
	if rs.eof then response.write "<option value=''> </option>"
		Do Until rs.eof
		response.write "<option value='" & rs(0) & "' "
		If rs(0) = highcustcode Then response.write " selected"
		response.write ">" & rs(1) & "</option>"
		rs.movenext
	Loop
	response.write "</select>"

	If Err.number <> 0 Then
		Call Debug
	End If

%>