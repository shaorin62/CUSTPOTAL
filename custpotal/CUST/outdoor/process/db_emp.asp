<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->
<%
'	On Error Resume Next
	Dim item
	For Each item In request.Form
		response.write item & " :" & request.Form(item) & "<br>"
	Next
'	response.End

	Dim atag : atag = ""
	Dim medcode : medcode = clearXSS(request("medcode"), atag)
	Dim emppwd : emppwd = clearXSS(request("emppwd"), atag)
	Dim empid : empid = request("empid")
	Dim empname : empname = clearXSS(request("empname"), atag)
	Dim useflag : useflag = clearXSS(request("useflag"), atag)
	Dim crud : crud = clearXSS(request("crud"), atag)

	Dim sql
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adcmdText


	Select Case UCase(crud)
		Case "U"
			sql = "update wb_med_employee set empname=?, emppwd=?, useflag=? where empid=?"
			cmd.parameters.append cmd.createparameter("empname", advarchar, adparaminput, 50)
			cmd.parameters.append cmd.createparameter("emppwd", advarchar, adparaminput, 50)
			cmd.parameters.append cmd.createparameter("useflag", adchar, adparaminput, 1)
			cmd.parameters.append cmd.createparameter("empid", adchar, adparaminput, 9)
			cmd.parameters("empid").value = empid
			cmd.parameters("empname").value = empname
			cmd.parameters("emppwd").value = emppwd
			cmd.parameters("useflag").value = useflag
			cmd.commandText = sql
			cmd.Execute ,, adExecuteNoRecords
		Case "C"
			sql = "insert into wb_med_employee (empid, medcode, empname, emppwd, useflag, ispwdchange, clipinglevel, class) values (?, ?, ?, ?, ?, ?, ?,?)"
			cmd.parameters.append cmd.createparameter("empid", adchar, adparaminput, 9)
			cmd.parameters.append cmd.createparameter("medcode", adchar, adparaminput,6)
			cmd.parameters.append cmd.createparameter("empname", advarchar, adparaminput, 50)
			cmd.parameters.append cmd.createparameter("emppwd", advarchar, adparaminput, 50)
			cmd.parameters.append cmd.createparameter("useflag", adchar, adparaminput, 1)
			cmd.parameters.append cmd.createparameter("ispwdchange", adBoolean, adparaminput)
			cmd.parameters.append cmd.createparameter("clipinglevel", adUnsignedTinyInt, adparaminput)
			cmd.parameters.append cmd.createparameter("class", adchar, adparaminput,1)
			cmd.parameters("empid").value = empid
			cmd.parameters("medcode").value = medcode
			cmd.parameters("empname").value = empname
			cmd.parameters("emppwd").value = emppwd
			cmd.parameters("useflag").value = useflag
			cmd.parameters("ispwdchange").value = false
			cmd.parameters("clipinglevel").value = 0
			cmd.parameters("class").value = "M"
			cmd.commandText = sql
			cmd.Execute ,, adExecuteNoRecords
			clearparameter(cmd)
		Case "D"
			sql  ="select count(*) from wb_contact_md where empid=?"
			cmd.parameters.append cmd.createparameter("empid", adchar, adparaminput, 9)
			cmd.parameters("empid").value = empid
			cmd.commandText = sql
			Dim rs : Set rs = cmd.Execute

			If rs(0) = 0 Then
				sql = "delete from wb_med_employee where empid=?"
				cmd.parameters("empid").value = empid
				cmd.commandText = sql
				cmd.Execute ,, adExecuteNoRecords
				clearparameter(cmd)
			Else
				response.write "<script> window.opener.msg('계약에 등록된 매체사 직원은 삭제할 수 없습니다.); window.close();</script>"
			End If
	End Select
	Set cmd = Nothing

'	If Err.number <> 0 Then
'		Call Debug
'	End If

'		response.write "mdidx : " & mdidx & "<br>"
'		response.write "side : " & side & "<br>"
'		response.write "thmno : " & thmno & "<br>"
'		response.write "startdate : " & startdate & "<br>"
'		response.write "no : " & no & "<br>"
'		response.write "enddate : " & enddate & "<br>"
'		response.write "cyear : " & cyear & "<br>"
'		response.write "cmonth : " & cmonth & "<br>"
%>
<script type="text/javascript">
<!--
	window.opener.getmedemployee();
	window.close();
//-->
</script>