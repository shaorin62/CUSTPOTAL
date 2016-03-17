<%@CODEPAGE=65001%>
<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
		Dim pcrud : pcrud = request("crud")
		Dim pname : pname = request("name")
		Dim pvalue : pvalue = request("value")
		Dim msg 

		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandType = adCmdText

		Select Case UCase(pcrud)
			Case "C"
				sql = "select max(codevalue) from wb_code_library where category = 'quality'"
				cmd.commandText = sql
				Set rs = cmd.Execute
				Dim value : value = CInt(rs(0))+1
				If Len(value) = 1 Then value = "0"&value		
				
				sql = "insert into wb_code_library(category, codename, codevalue) values (?, ?, ?)"
				cmd.commandText = sql
				cmd.parameters.append cmd.createparameter("category", advarchar, adparaminput, 50)
				cmd.parameters.append cmd.createparameter("codename", advarchar, adparaminput, 100)
				cmd.parameters.append cmd.createparameter("codevalue", advarchar, adparaminput, 10)
				cmd.parameters("category").value = "quality"
				cmd.parameters("codename").value = pname
				cmd.parameters("codevalue").value = value
				cmd.execute ,, adexecutenorecords
			Case "U"
				sql = "update wb_code_library set codename=? where category='quality' and codevalue=?"
				cmd.commandText = sql
				cmd.parameters.append cmd.createparameter("codename", advarchar, adparaminput, 100)
				cmd.parameters.append cmd.createparameter("codevalue", advarchar, adparaminput, 10)
				cmd.parameters("codename").value = pname
				cmd.parameters("codevalue").value = pvalue
				cmd.execute ,, adexecutenorecords
			Case "D"
				sql = "select count(*) from wb_contact_md_dtl where quality=?"
				cmd.commandText = sql 
				cmd.parameters.append cmd.createparameter("codename", advarchar, adparaminput, 100)
				cmd.parameters("codename").value = pname
				Set rs = cmd.Execute
				If rs(0) = 0 Then 
					sql = "delete from wb_code_library where category='quality' and codename=?"
					cmd.commandText = sql 
					cmd.parameters("codename").value = pname
					cmd.execute ,, adexecutenorecords
				Else 
					msg="&nbsp;<script defer> alert('시스템에 사용중인 데이터는 삭제할 수 없습니다.')</script>"
				End If 				
		End Select
		clearparameter(cmd)
		sql = "select codevalue, codename from wb_code_library where category='quality'  order by codename"
		cmd.commandText = sql
		Dim rs : Set rs = cmd.execute 
		Set cmd = Nothing

		Dim codevalue : Set codevalue = rs(0)
		Dim codename : Set codename = rs(1)
			
		response.write "<select id='cmbquality' name='cmbquality' style='width:265px;' size='18'>"
		Do Until rs.eof 
		response.write "<option value='"& codevalue & "'> " & codename & "</option>"
			rs.movenext
		Loop
		response.write "</select>"
		rs.close
		Set rs = Nothing 
		If Not IsEmpty(msg) Then 
			response.write msg
		End If 

%>