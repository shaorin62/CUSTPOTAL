<%@CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
	On Error Resume Next 
	Dim psubno : psubno = request("subno")
	Dim pthmno : pthmno  = request("thmno")
	Dim sql : sql = "select thmno, thmname from wb_subseq_dtl where subno = '" & psubno & "' order by thmname"
'	response.write sql
	
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	Dim rs : Set rs = cmd.Execute 
	Set cmd = Nothing

	response.write "<select  id='cmbthmno' name='cmbthmno' style='width:120px;' class='theme'>"
	response.write "<option value=''> -- 소재명 -- </option>"
	Do Until rs.eof 
		response.write "<option value='" & rs("thmno") & "' "
		If rs("thmno") = pthmno Then response.write " selected"
		response.write ">" & rs("thmname") & "</option>" 
		rs.movenext
	Loop
	response.write "</select>"
	
	If Err.number <> 0 Then 
		Call Debug
	End If 

%>