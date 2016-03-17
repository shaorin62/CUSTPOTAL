<%@CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
	Dim phighseqno : phighseqno = request("highseqno")
	Dim psubno : psubno = request("subno")



	Dim sql : sql = "select subno, subname from wb_subseq_mst where seqno = '" & phighseqno & "' order by subname"
	



	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	Dim rs : Set rs = cmd.Execute
	Set cmd = Nothing

	response.write "<select id='cmbsubno' name='cmbsubno' style='width:120px;' class='subbrand'>"
	response.write "<option value=''> -- 서브브랜드 -- </option>"
	Do Until rs.eof
		response.write "<option value='" & rs("subno") & "' "
		If rs("subno") = psubno Then response.write " selected"
		response.write ">" & rs("subname") & "</option>"
		rs.movenext
	Loop
	response.write "</select>"
%>