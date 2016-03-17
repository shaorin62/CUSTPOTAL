<!--#include virtual="/mp/outdoor/inc/Function.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>
<%
'	For Each item In request.Form
'		response.write item & " : " & request.Form(item) & "<br>"
'	Next
'	response.End

	Dim cyear : cyear = request("cyear")
	Dim cmonth : cmonth = request("cmonth")
	Dim side : side = request("side")
	Dim mdidx : mdidx = clearXSS(request("mdidx"), atag)
	Dim standard : standard = request("txtstandard")
	Dim quality : quality = request("cmbquality")

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adCmdText

	sql = "insert into wb_contact_md_dtl (mdidx, cyear, cmonth, side, standard, quality) values (?, ?, ?, ?,?,?)"
	cmd.commandText = sql
	cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput, , mdidx)
	cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4, cyear)
	cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2, cmonth)
	cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1, side)
	cmd.parameters.append cmd.createparameter("standard", advarchar, adparaminput, 200, standard)
	cmd.parameters.append cmd.createparameter("quality", advarchar, adparaminput, 200, quality)
	cmd.Execute ,, adExecuteNoRecords

	set cmd = nothing

%>
<script type="text/javascript">
<!--
	window.opener.getcontact();
	window.close();
//-->
</script>