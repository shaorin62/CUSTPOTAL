<%@CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
	' parameter
	' custcode : 광고주 코드 -  선택할  광고주이 없는 경우  null, 코드가 있으면 해당 광고주을 선택
	Dim highcategory : highcategory = UCase(Trim(Request("highcategory")))
	If highcategory = "" Then highcategory = null

	Dim sql 

	sql = "SELECT CATEGORYIDX HIGHCATEGORYIDX, CATEGORYNAME HIGHCATEGORYNAME FROM WB_REPORT_CATEGORY WHERE CATEGORYLVL = 0 AND USE_YN = 1 "

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	Dim rs : Set rs = cmd.Execute 
	Set cmd = Nothing

	Response.write "<select id='cmbhighcategory' name='cmbhighcategory' style='width166px' onchange=getcategorycombo();>"&vbCrLf
	response.write "<option value=''> -- ALL -- </option>"&vbCrLf
	Do Until rs.eof 
		response.write "<option value='" & UCase(Trim(rs(0))) & "' "
		If highcategory = UCase(Trim(rs(0))) Then Response.write "selected"
		response.write ">" & rs(1) & "</option>" & vbCrLf
		rs.movenext
	Loop
	Response.write "</select>"
%>
