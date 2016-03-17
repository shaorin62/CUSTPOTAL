<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>
<%
'	For Each item In request.Form
'		response.write item & " : " & request.Form(item) & "<br>"
'	Next
'	response.End

	Dim atag : atag = ""
	Dim crud : crud = clearXSS(request("crud"), atag)
	Dim cmonth : cmonth = clearXSS(request("cmonth"), atag)
	Dim monthly : monthly = Replace(request("txtmonthly"), ",", "")
	Dim cyear : cyear = clearXSS(request("cyear"), atag)
	Dim mdidx : mdidx = clearXSS(request("mdidx"), atag)
	Dim region : region = clearXSS(request("cmbregion"), atag)
	Dim locate : locate = clearXSS(request("txtlocate"), atag)
	Dim contdix : contidx = request("contidx")
	Dim qty : qty = clearXSS(request("txtqty"), atag)
	Dim unit : unit = clearXSS(request("txtunit"), atag)
	Dim standard : standard = clearXSS(request("txtstandard"), atag)
	Dim quality : quality = clearXSS(request("cmbquality"), atag)
	Dim trust : trust = clearXSS(request("rdotrust"), atag)
	Dim empid : empid = clearXSS(request("cmbemp"), atag)
	Dim thmno : thmno = request("hdnthmno")
	Dim medcode : medcode = clearXSS(request("cmbmed"), atag)
	Dim expense : expense = Replace(request("txtexpense"), ",", "")
	Dim categoryidx : categoryidx = request("hdncategoryidx")
	Dim subseq : subseq = request("subseq")
	Dim intLoop, pcyear, pcmonth, attachFile
	dim no : no = request("no")
	If mdidx = "" Then mdidx = 0
	If empid = "" Then empid = Null


	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adCmdText

	dim sql : sql = "insert into wb_contact_md_dtl (mdidx, side, cyear, cmonth, standard, quality) values (?, ?, ?, ?, ?, ?)"
	cmd.commandText = sql
	cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput)
	cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1)
	cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4)
	cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2)
	cmd.parameters.append cmd.createparameter("standard", advarchar, adparaminput, 200)
	cmd.parameters.append cmd.createparameter("quality", advarchar, adparaminput, 200)
	cmd.parameters("mdidx").value = mdidx
	cmd.parameters("side").value = "F"
	cmd.parameters("cyear").value = cyear
	cmd.parameters("cmonth").value = cmonth
	cmd.parameters("standard").value = standard
	cmd.parameters("quality").value = quality
	cmd.Execute ,, adExecuteNoRecords

	Set cmd = Nothing

%>
<script type="text/javascript">
<!--
	window.opener.getcontact();
	window.close();
//-->
</script>