<%@CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
	On Error Resume Next

	Dim pcontidx : pcontidx = request("contidx")
	Dim cyear : cyear = request("cyear")
	Dim sql : sql = "select startdate, enddate from wb_contact_mst where contidx =" & pcontidx

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	Dim rs : Set rs = cmd.Execute
	Set cmd = Nothing

	Dim syear : syear = Year(rs(0))
	Dim eyear : eyear = Year(rs(1))

	Dim intLoop
	response.write "<select id='cyear' name='cyear' >"
	For intLoop = syear To eyear
		response.write "<option value='" & intLoop & "' "
			If CStr(intLoop) = cyear Then response.write " selected "
		response.write ">" & intLoop
	Next
	response.write "</select>"

	If Err.number <> 0 Then
		Call Debug
	End If

%>