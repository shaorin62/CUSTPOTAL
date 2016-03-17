<%@CODEPAGE=65001%>
<!--#include virtual="/MP/outdoor/inc/Function.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
'	On Error Resume Next

	Dim pcrud : pcrud = request("crud")
	Dim psubno : psubno = request("subno")
	Dim pthmname : pthmname = request("thmname")
	Dim pthmno : pthmno = request("thmno")

	dim item
	for each item in request.querystring
		response.write request.querystring(item) & "<br>"
	next

	Dim sql
	Dim cmd :  Set  cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	Select Case UCase(pcrud)
		Case "C"
			If pthmno = "" Then
				pthmno = psubno+"01"
			Else
				Dim num : num = Int(Right(pthmno,2)) + 1
				If num < 10 Then num = "0"&num
				pthmno = psubno&num
			End If
			response.write pthmno
			sql = "insert into wb_subseq_dtl values (?, ?, ?)"
			cmd.parameters.append cmd.createparameter("thmno", adChar, adParamInput, 12, pthmno)
			cmd.parameters.append cmd.createparameter("thmname", adVarWChar, adParamInput, 255, pthmname)
			cmd.parameters.append cmd.createparameter("subno", adChar, adParamInput, 10, psubno)

		Case "U"
			sql = "update wb_subseq_dtl set thmname = ? where thmno = ?"
			cmd.parameters.append cmd.createparameter("thmname", adVarWChar, adParamInput, 255, pthmname)
			cmd.parameters.append cmd.createparameter("thmno", adChar, adParamInput, 12, pthmno)
		Case "D"
			sql = "delete from wb_subseq_dtl  where thmno = ?"
			cmd.parameters.append cmd.createparameter("thmno", adChar, adParamInput, 12, pthmno)
	End Select
	cmd.commandText = sql
	cmd.commandType = adcmdText
	cmd.Execute ,, adExecuteNoRecords
	Set cmd = Nothing
'
'	response.write sql & "<P>"
'	response.write "number : " & Err.number & "<br>"
'	response.write "description : " & Err.description & "<br>"
'	response.write "source : " & Err.source

%>