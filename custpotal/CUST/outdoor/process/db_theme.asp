<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->
<%
'	On Error Resume Next
'	Dim item
'	For Each item In request.Form
'		response.write item & " :" & request.Form(item) & "<br>"
'	Next
'	response.End

	Dim atag : atag = ""
	Dim custcode : custcode = clearXSS(request("custcode"), atag)
	Dim seq : seq = clearXSS(request("seq"), atag)
	Dim mdidx : mdidx = clearXSS(request("mdidx"), atag)
	Dim side : side = clearXSS(request("side"), atag)
	Dim thmno : thmno = clearXSS(request("cmbthmno"), atag)
	Dim cyear : cyear = clearXSS(request("cyear"), atag)
	Dim cmonth : cmonth = clearXSS(request("cmonth"), atag)
	Dim no : no = clearXSS(request("txtno"), atag)
	Dim crud : crud = clearXSS(request("crud"), atag)
	If Not CBool(Len(thmno)) Then thmno = null

	Dim sql
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adcmdText

	Select Case UCase(crud)
		Case "U"
			sql = "update wb_subseq_exe set thmno=?,  no=?, cyear=?, cmonth=? where seq=?"
			cmd.parameters.append cmd.createparameter("thmno", adchar, adparaminput, 12, thmno)
			cmd.parameters.append cmd.createparameter("no", adUnsignedTinyInt, adparaminput, , no)
			cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4, cyear)
			cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2, cmonth)
			cmd.parameters.append cmd.createparameter("seq", adinteger, adparaminput, , seq)
			cmd.commandText = sql
			cmd.Execute ,, adExecuteNoRecords
			clearparameter(cmd)
		Case "C"
			sql = "insert into wb_subseq_exe (mdidx, side, no, thmno,  cyear, cmonth) values (?, ?, ?, ?, ?, ?)"
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput, , mdidx)
			cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1, side)
			cmd.parameters.append cmd.createparameter("no", adUnsignedTinyInt, adparaminput, , no)
			cmd.parameters.append cmd.createparameter("thmno", adchar, adparaminput, 12, thmno)
			cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4, cyear)
			cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2, cmonth)
			cmd.commandText = sql
			cmd.Execute ,, adExecuteNoRecords
			clearparameter(cmd)

		Case "D"
			sql = "delete from wb_subseq_exe where seq =? "
			cmd.parameters.append cmd.createparameter("seq", adinteger, adparaminput, , seq)
			cmd.commandText = sql
			cmd.Execute ,, adExecuteNoRecords
	End Select
	Set cmd = Nothing

	If Err.number <> 0 Then
		Call Debug
	End If

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
	//location.href="/cust/outdoor/popup/view_theme.asp?custcode=<%=custcode%>&mdidx=<%=mdidx%>&side=<%=side%>&lastdate=<%=lastdate%>";
				window.opener.getcontactdetail();
				window.close();
//-->
</script>