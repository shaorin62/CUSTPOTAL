<%@CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
	On Error Resume Next 

	Dim pcontidx : pcontidx = request("contidx")
	Dim cyear : cyear = request("cyear")
	Dim cmonth : cmonth = request("cmonth")
	Dim sql : sql = "select distinct cmonth from  wb_contact_md a inner join wb_contact_exe b on a.mdidx=b.mdidx where a.contidx =" & pcontidx & " and cyear = '" & cyear & "' "
'	response.write sql
	
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	Dim rs : Set rs = cmd.Execute 
	Set cmd = Nothing

	response.write "<select id='cmonth' name='cmonth' >"
	If rs.eof Then response.write "<option value='" & cmonth & "'>" & cmonth & "</option>"
	Do Until rs.eof 
		response.write "<option value='" & rs(0) & "' "
		If CInt(rs(0)) = CInt(cmonth) Then response.write " selected"
		response.write ">" & rs(0) & "</option>" 
		rs.movenext
	Loop
	response.write "</select>"
	
	If Err.number <> 0 Then 
		Call Debug
	End If 

%>