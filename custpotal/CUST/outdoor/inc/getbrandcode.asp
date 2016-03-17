<%@CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
	Dim highcustcode : highcustcode = request("highcustcode")
	Dim pseqno : pseqno = request("seqno")
	Dim sql : sql = "select highseqno, highseqname from sc_subseq_hdr where custcode like  '" & highcustcode & "' order by highseqname"
	'response.write sql

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	Dim rs : Set rs = cmd.Execute
	Set cmd = Nothing

	response.write "<select size='20' id='cmbseqno' name='cmbseqno' style='width:225px;'>"
	Do Until rs.eof
		response.write "<option value='" & rs("highseqno") & "' "
		If rs("highseqno") = pseqno Then response.write " selected"
		response.write ">" & rs("highseqname") & "</option>"
		rs.movenext
	Loop
	response.write "</select>"
%>