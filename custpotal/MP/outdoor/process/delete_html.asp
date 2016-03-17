<!--#include virtual="/mp/outdoor/inc/Function.asp" -->
<%
'	For Each item In request.form
'		response.write item & " : "& request.form(item) & "<br>"
'	Next
'	response.End
	Dim path : path = "\\11.0.12.201\adportal\print"
	Dim sql, rs , filename, fullpath
	Dim cyear : cyear = request("cyear")
	Dim cmonth : cmonth = request("cmonth")
	Dim fso : Set fso = CreateObject("scripting.filesystemobject")
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adCmdText
	cmd.parameters.append cmd.createparameter("contidx", adInteger, adParamInput)
	cmd.parameters.append cmd.createparameter("cyear", adChar, adParamInput, 4)
	cmd.parameters.append cmd.createparameter("cmonth", adChar, adParamInput, 2)

	For intLoop = 1 To request("contidx").count
		sql = "select report from wb_report_mst where contidx=? and cyear=? and cmonth=?"
		cmd.commandText = sql
		cmd.parameters("contidx").value = request("contidx")(intLoop)
		cmd.parameters("cyear").value = cyear
		cmd.parameters("cmonth").value = cmonth
		Set rs = cmd.execute

		If Not rs.eof Then
			fullpath = "\\11.0.12.201\adportal\print\"&rs(0)
			If fso.FileExists(fullpath) Then
				fso.DeleteFile(fullpath)
				sql = "delete from wb_report_mst where contidx=? and cyear=? and cmonth=?"
				cmd.commandText = sql
				cmd.parameters("contidx").value = request("contidx")(intLoop)
				cmd.parameters("cyear").value = cyear
				cmd.parameters("cmonth").value = cmonth
				cmd.execute ,, adexecutenorecords
			End If
		End If
		rs.close
	Next
	Set rs = Nothing
	Set fso = Nothing
	Set cmd = Nothing

		response.write "<script> alert('선택한 파일이 삭제되었습니다.'); parent.location.replace('/mp/outdoor/list_report.asp?cmbcustcode="&request("cmbcustcode")&"&cmbteamcode="&request("cmbteamcode")&"&cyear="&request("cyear")&"&cmonth="&request("cmonth")&"&menunum="&request("menunum")&"'); </script>"
%>

