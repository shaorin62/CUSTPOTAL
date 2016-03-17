<%@CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
	Dim pseqno : pseqno = request("seqno")
	Dim psubno : psubno = request("subno")
	Dim sql : sql = "select subno, subname from wb_subseq_mst where seqno = '" & pseqno & "' order by subname"
	'response.write sql
	
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	Dim rs : Set rs = cmd.Execute 
	Set cmd = Nothing

	response.write "<select size='20' id='cmbsubno' name='cmbsubno' style='width:225px;' class='subbrand'>"
	Do Until rs.eof 
		response.write "<option value='" & rs("subno") & "' "
		If rs("subno") = psubno Then response.write " selected"
		response.write ">" & rs("subname") & "</option>" 
		rs.movenext
	Loop
	response.write "</select>"
%>