<%@CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
	On Error Resume Next
	Dim pcustcode : pcustcode = request("custcode")
	Dim phighseqno : phighseqno = request("highseqno")
	Dim pthmno : pthmno  = request("thmno")
	Dim pCustcodesql : pCustcodesql = request("Custcodesql")



	Dim sql

	If phighseqno = ""  Then
		sql = " select thmno, thmname from wb_subseq_dtl  "
		sql = sql & " where subno = '' "

	Else
		sql = " select thmno, thmname from wb_subseq_dtl  "
		sql = sql & " where subno in  "
		sql = sql & " (  "
		sql = sql & " 	select subno from wb_subseq_mst "
		sql = sql & " 	where seqno = '" & phighseqno &"' "
		sql = sql & " ) "
		sql = sql & " order by thmname "
	End If

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