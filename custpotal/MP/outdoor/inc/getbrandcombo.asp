<%@CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%

	Dim pcustcode : pcustcode = request("custcode")
	Dim pseqno : pseqno = request("seqno")
	Dim pCustcodesql : pCustcodesql = request("Custcodesql")


	Dim sql

	If pcustcode = ""  Then
		sql = "select highseqno, highseqname from sc_subseq_hdr where custcode = '' "
	Else
		sql = "select highseqno, highseqname from sc_subseq_hdr where custcode ='" & pcustcode &"' order by highseqname"
	End If


	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	Dim rs : Set rs = cmd.Execute
	Set cmd = Nothing

	response.write "<select  id='cmbseqno' name='cmbseqno' style='width:120px;'>"
	response.write "<option value=''> -- 브랜드 -- </option>"
	Do Until rs.eof
		response.write "<option value='" & rs("highseqno") & "' "
		If rs("highseqno") = pseqno Then response.write " selected"
		response.write ">" & rs("highseqname") & "</option>"
		rs.movenext
	Loop
	response.write "</select>"
%>