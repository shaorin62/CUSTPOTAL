<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>
<%
	On Error Resume Next
'
'	Dim item
'	For Each item  In request.Form
'		response.write item & " :"  & request.Form(item) & "<br>"
'	Next

	Dim atag : atag = ""
	Dim pcontidx : pcontidx = clearXSS(request("contidx"), atag)
	Dim pmdidx : pmdidx = clearXSS(request("mdidx"), atag)
	Dim pside : pside = clearXSS(request("side"), atag)
	Dim pflag : pflag = clearXSS(request("flag"), atag)
	
	Dim intLoop 
	Dim sql : sql ="update wb_contact_exe set monthly=?, expense=? where mdidx=? and side=? and cyear=? and cmonth=?"
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adCmdText
	cmd.parameters.append cmd.createparameter("monthly", adCurrency, adParamInput)
	cmd.parameters.append cmd.createparameter("expense", adCurrency, adParamInput)
	cmd.parameters.append cmd.createparameter("mdidx", adInteger, adParamInput)
	cmd.parameters.append cmd.createparameter("side", adChar, adParamInput, 1)
	cmd.parameters.append cmd.createparameter("cyear", adChar, adParamInput, 4)
	cmd.parameters.append cmd.createparameter("cmonth", adChar, adParamInput, 2)

	For intLoop = 1 To request.Form("monthly").count
		cmd.commandText = sql
		cmd.parameters("monthly").value = clearXSS(request.Form("monthly")(intLoop), atag)
		cmd.parameters("expense").value = clearXSS(request.Form("expense")(intLoop), atag)
		cmd.parameters("mdidx").value = pmdidx
		cmd.parameters("side").value = pside
		cmd.parameters("cyear").value = clearXSS(request.Form("cyear")(intLoop), atag)
		cmd.parameters("cmonth").value = clearXSS(request.Form("cmonth")(intLoop), atag)
		cmd.Execute ,, adExecuteNoRecords
	Next
	clearparameter(cmd)
	
	If UCase(Trim(pflag)) = "B" Then 
		sql= "update wb_contact_mst set totalprice = (select sum(monthly) from wb_contact_exe where mdidx = "&pmdidx&") from wb_contact_mst m inner join wb_contact_md m2 on m.contidx=m2.contidx where m2.mdidx = " & pmdidx
		cmd.commandText = sql
	Else 
		sql = "update wb_contact_mst set totalprice=(select sum(monthly)  from wb_contact_exe a inner join wb_contact_md b on a.mdidx=b.mdidx inner join wb_contact_mst c on c.contidx = b.contidx where c.contidx="&pcontidx&") where contidx="&pcontidx
		response.write sql
		cmd.commandText = sql
	End If 
	cmd.Execute ,, adExecuteNoRecords
	Set cmd = Nothing 

	If Err.number <> 0 Then 
		Call Debug
	End If 
%>
<script language="JavaScript">
<!--
	window.opener.getcontact();
	window.close();
//-->
</script>