<%@CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
	Dim highcategory : highcategory = UCase(Trim(Request("highcategory")))
	Dim category : category = UCase(Trim(request("category")))

	Dim sql : sql = "SELECT CATEGORYIDX, CATEGORYNAME FROM WB_REPORT_CATEGORY"
	sql = sql  & " WHERE CATEGORYLVL = 1 "
	sql = sql  & " AND HIGHCATEGORYIDX = '" & highcategory & "'" 

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	Dim rs : Set rs = cmd.Execute
	Set cmd = Nothing


	Response.write "<select id='cmbcategory' name='cmbcategory'  style='width:106px' onchange=checkForSearch();>"&vbCrLf
	Response.write "<option value=''> -- ALL --</option>"
	Do Until rs.eof
		Response.write "<option value='" &  UCase(Trim(rs(0))) & "' "
		If category = UCase(Trim(rs(0))) Then Response.write "selected"
		Response.write ">" & rs(1) & "</option>" & vbCrLf
		rs.movenext
	Loop
	Response.write "</select>"


%>
