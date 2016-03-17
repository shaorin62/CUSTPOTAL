<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->
<%
	On Error Resume Next
	
'	Dim item 
'	For Each item In request.Form
'		Response.write item & " : " & request.Form(item) & "<br>"
'	next
	Dim intLoop 
	Dim atag : atag = ""	
	Dim crud : crud = clearXSS(request.form("crud"), atag)
	Dim cyear : cyear = clearXSS(request.form("cyear"), atag)
	Dim cmonth : cmonth = clearXSS(request.form("cmonth"), atag)
	Dim custcode : custcode =  clearXSS(request.form("cmbcustcode"), atag)
	Dim teamcode : teamcode = clearXSS(request.form("cmbteamcode"), atag)
	Dim menunum : menunum = request.form("menunum")
	Dim sql
	Dim pk
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adcmdText

	Select Case UCase(crud)
		Case "U"
			sql = "insert into wb_contact_trans(cyear, cmonth, contidx, medcode, isHold, uuser, udate) values (?,?,?,?,?,?,?)"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("cyear", adChar, adparaminput, 4)
			cmd.parameters.append cmd.createparameter("cmonth", adChar, adparaminput, 2)
			cmd.parameters.append cmd.createparameter("contidx", adinteger, adparaminput)
			cmd.parameters.append cmd.createparameter("medcode", adChar, adparaminput, 6)
			cmd.parameters.append cmd.createparameter("isHold", adChar, adparaminput, 1)
			cmd.parameters.append cmd.createparameter("uuser", adVarChar, adparaminput, 12)
			cmd.parameters.append cmd.createparameter("udate", adDBTimeStamp, adParamInput)
			
			For intLoop = 1 To Request.Form("contidx").count	
				pk = Split(Request.Form("contidx")(intLoop), ",")
				cmd.parameters("cyear").value = cyear
				cmd.parameters("cmonth").value = cmonth
				cmd.parameters("contidx").value = pk(0)
				cmd.parameters("medcode").value = pk(1)
				cmd.parameters("isHold").value = "N"
				cmd.parameters("uuser").value = session("userid")
				cmd.parameters("udate").value = Date()

				cmd.execute ,, adExecuteNoRecords
			Next
		Case "C"
			sql = "delete from wb_contact_trans where cyear=? and cmonth=? and contidx=? and medcode=?"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("cyear", adChar, adparaminput, 4)
			cmd.parameters.append cmd.createparameter("cmonth", adChar, adparaminput, 2)
			cmd.parameters.append cmd.createparameter("contidx", adinteger, adparaminput)
			cmd.parameters.append cmd.createparameter("medcode", adChar, adparaminput, 6)
			
			For intLoop = 1 To Request.Form("contidx").count		
				pk = Split(Request.Form("contidx")(intLoop), ",")
				cmd.parameters("cyear").value = cyear
				cmd.parameters("cmonth").value = cmonth
				cmd.parameters("contidx").value = pk(0)
				cmd.parameters("medcode").value = pk(1)

				cmd.execute ,, adExecuteNoRecords
			Next
	End Select 
	Set cmd = Nothing 

	If Err.number <> 0 Then 
		Call Debug
	End If 

%>
<script type="text/javascript">
<!--
	location.href='/cust/outdoor/list_transaction.asp?cyear=<%=cyear%>&cmonth=<%=cmonth%>&cmbcustcode=<%=custcode%>&cmbteamcode=<%=teamcode%>&menunum=<%=menunum%>';
//-->
</script>