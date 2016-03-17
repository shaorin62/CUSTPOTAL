<%@CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
	On Error Resume Next

	dim pmedcode : pmedcode = request("medcode")
	Dim pmedname : pmedname = request("medname")
	Dim sql : sql = "select distinct a.highcustcode, a.custname from sc_cust_hdr a inner join sc_cust_dtl b on a.highcustcode=b.highcustcode where a.medflag='B' and a.use_flag=1 and b.med_out = '1' and a.custname like '%" & pmedname & "%' order by a.custname"
'	response.write sql

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