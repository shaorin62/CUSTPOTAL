<%@CODEPAGE=65001%>
<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
'	On Error Resume Next
	Dim pcrud : pcrud = request("crud")
	Dim psubno : psubno = request("subno")
	Dim psubname : psubname = request("subname")
	Dim pseqno : pseqno = request("seqno")
	Dim sql
	sql = "select max(subno) from wb_subseq_mst where seqno = ?"

	Dim cmd :  Set  cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	cmd.parameters.append cmd.createparameter("seqno", adChar, adParamInput, 8, pseqno)
	Dim rs : Set rs = cmd.execute
	clearparameter(cmd)
	
	Select Case UCase(pcrud)
		Case "C"
			If rs.eof  Then 
				psubno = pseqno+"01" 
			Else 
				Dim num : num = Int(Right(rs(0),2)) + 1
				If num < 10 Then num = "0"&num 
				psubno = pseqno&num
			End If 

			sql = "insert into wb_subseq_mst values (?, ?, ?)"
			cmd.parameters.append cmd.createparameter("subno", adChar, adParamInput, 10, psubno)
			cmd.parameters.append cmd.createparameter("subname", adVarWChar, adParamInput, 255, psubname)
			cmd.parameters.append cmd.createparameter("seqno", adChar, adParamInput, 8, pseqno)	
			cmd.commandText = sql
			cmd.commandType = adcmdText
			cmd.Execute ,, adExecuteNoRecords
		Case "U"
			sql = "update wb_subseq_mst set subname = ? where subno = ?"
			cmd.parameters.append cmd.createparameter("subname", adVarWChar, adParamInput, 255, psubname)
			cmd.parameters.append cmd.createparameter("subno", adChar, adParamInput, 10, psubno)	
			cmd.commandText = sql
			cmd.commandType = adcmdText
			cmd.Execute ,, adExecuteNoRecords
		Case "D"
			sql = "delete from wb_subseq_dtl where subno =?"
			cmd.parameters.append cmd.createparameter("subno", adChar, adParamInput, 10, psubno)
			cmd.commandText = sql
			cmd.commandType = adcmdText
			cmd.Execute ,, adExecuteNoRecords
	clearparameter(cmd)
			sql = "delete from wb_subseq_mst  where subno = ?"
			cmd.parameters.append cmd.createparameter("subno", adChar, adParamInput, 10, psubno)
			cmd.commandText = sql
			cmd.commandType = adcmdText
			cmd.Execute ,, adExecuteNoRecords
	End Select 

	Set cmd = Nothing 
	response.write sql
	response.End
	
'	response.write "number : " & Err.number & "<br>"
'	response.write "description : " & Err.description & "<br>"
'	response.write "source : " & Err.source 
'	Err.clear

%>