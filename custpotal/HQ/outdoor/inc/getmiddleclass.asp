<%@CODEPAGE=65001%>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
	On Error Resume Next

	Dim pcrud : pcrud = request("crud")
	Dim phighclasscode : phighclasscode = request("highclasscode")
	Dim pmiddleclasscode : pmiddleclasscode = request("middleclasscode")
	Dim pmiddleclassname : pmiddleclassname = request("middleclassname")
	If phighclasscode = "" Then phighclasscode = 0
	Dim sql 
	
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adCmdText


	Select Case UCase(pcrud)
		Case "D"' delete
			sql = "delete from dbo.wb_category where categoryidx = ?"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("categoryidx", adInteger, adParaminput, 4, pmiddleclasscode)
			cmd.Execute , , adExecuteNoRecords
		Case "C"	' insert
			sql = "insert into wb_category values (? ,? ,?)"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("categoryname", adVarChar, adParaminput, 50, pmiddleclassname)
			cmd.parameters.append cmd.createparameter("categorylvl", adUnsignedTinyInt, adParaminput, 1, 1)
			cmd.parameters.append cmd.createparameter("highcategoryidx", adInteger, adParaminput, 4, phighclasscode)
			cmd.Execute , , adExecuteNoRecords
		Case "U"	' update
			sql = "update wb_category set categoryname = ? where categoryidx = ?"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("categoryname", adVarChar, adParaminput, 50, pmiddleclassname)
			cmd.parameters.append cmd.createparameter("categoryidx", adInteger, adParaminput, 4, pmiddleclasscode)
			cmd.Execute , , adExecuteNoRecords
	End Select
	
	sql = "select categoryidx, categoryname from wb_category where categorylvl = 1 and highcategoryidx = " & phighclasscode & " order by categoryname"
	cmd.commandText = sql

	Dim rs : Set rs = cmd.Execute 
	Set cmd = Nothing 

	response.write "<select id='cmbmiddleclass' name='cmbmiddleclass' class='middleclass' style='width:250px;'  size='20'>" & VbCrLf
		Do Until rs.eof 
			response.write "<option value='" & rs("categoryidx") & "'> " & rs("categoryname") & "</option>" & VbCrLf
			rs.movenext
		Loop
	response.write "</select>"

	
	If Err.Number <> 0 Then 
	For Each item In Request.querystring
		Response.write item  & " : " & request.querystring(item) & "<br>"
	Next 
	response.write "number :  " & Err.Number & "<br>"
	response.write "description ( "  & pcrud & "): " & Err.Description & "<br>"
	response.write "source : " & Err.source
	End If 
%>