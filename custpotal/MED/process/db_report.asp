<!--#include virtual="/hq/outdoor/inc/function.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>
<%

	If request.cookies("userid") = "" Then
		response.write "<script> try {this.close();} catch(e) {window.close();} </script>"
		response.end
	end if
	Dim uploadform : Set uploadform = Server.CreateObject ("DEXT.FileUpload")
	uploadform.defaultpath = "\\11.0.12.201\adportal\report"

	Dim atag : atag = ""
	Dim mdidx : mdidx = clearXSS(uploadform("mdidx"), atag)
	Dim cyear : cyear = clearXSS(uploadform("cyear"), atag)
	Dim cmonth : cmonth = clearXSS(uploadform("cmonth"), atag)
	Dim crud : crud = clearXSS(uploadform("crud"), atag)
	Dim orgfile : orgfile = uploadform("orgfile")
	Dim msg : msg = "등록"

'	response.write crud
'	response.end


	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adCmdText

	Select Case UCase(crud)
		Case "C"
			Dim sql : sql = "select a.custcode, categoryidx, title, medcode from wb_contact_mst a  inner join wb_contact_md b on a.contidx=b.contidx where mdidx=?"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput)
			cmd.parameters("mdidx").value = mdidx
			Dim rs : Set rs = cmd.Execute
			clearparameter(cmd)
			Dim custname : custname = getcustname(rs(0))
			Dim teamname : teamname = getteamname(rs(0))
			Dim categoryname : categoryname = getmediumname(rs(1))
			Dim title : title = rs(2)
			Dim medname : medname = getmedname(rs(3))

			rs.close
			sql = "select filename from wb_report_dtl where mdidx=? and cyear=? and cmonth=?"
			cmd.commandText = sql
			cmd.commandType = adcmdtext
			cmd.parameters.append cmd.createparameter("mdidx", adInteger, adparaminput)
			cmd.parameters.append cmd.createparameter("cyear", adChar, adparaminput, 4)
			cmd.parameters.append cmd.createparameter("cmonth", adChar, adparaminput, 2)
			cmd.parameters("mdidx").value = mdidx
			cmd.parameters("cyear").value = cyear
			cmd.parameters("cmonth").value = cmonth


			Set rs = cmd.execute
			clearparameter(cmd)
			If rs.eof Then
				Dim ext : ext = uploadform("file").FileExtension
				filename = cyear &cmonth & "_" & custname & "_" & teamname &"_"& categoryname & "_" & title & "_" & medname &"_" & request.cookies("custname")&"_"&mdidx&"."&ext

			filename = replace(filename, "\", "")
			filename = replace(filename, "/", "")
			filename = replace(filename, ":", "")
			filename = replace(filename, "*", "")
			filename = replace(filename, """", "")
			filename = replace(filename, "<", "")
			filename = replace(filename, ">", "")
			filename = replace(filename, "|", "")
			filename = replace(filename, "(", "")
			filename = replace(filename, ")", "")
				sql = "insert into wb_report_dtl(mdidx, cyear, cmonth, empid, reportname, filename, cuser, cdate) values (?, ?, ?, ?, ?, ?, ?, ?)"
				cmd.commandText = sql
				cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput)
				cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4)
				cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2)
				cmd.parameters.append cmd.createparameter("empid", adchar, adparaminput, 9)
				cmd.parameters.append cmd.createparameter("reportname", advarchar, adparaminput, 200)
				cmd.parameters.append cmd.createparameter("filename", advarchar, adparaminput, 200)
				cmd.parameters.append cmd.createparameter("cuser", advarchar, adparaminput, 12)
				cmd.parameters.append cmd.createparameter("cdate", addbtimestamp, adparaminput)
				cmd.parameters("mdidx").value = mdidx
				cmd.parameters("cyear").value = cyear
				cmd.parameters("cmonth").value = cmonth
				cmd.parameters("empid").value = request.cookies("userid")
				cmd.parameters("reportname").value = uploadform("file").filename
				cmd.parameters("filename").value = filename
				cmd.parameters("cuser").value = request.cookies("userid")
				cmd.parameters("cdate").value = date
			Else
				If uploadform.FileExists(uploadform.defaultpath&"\"&orgfile) Then uploadform.DeleteFile(uploadform.defaultpath&"\"&orgfile)
				ext = uploadform("file").FileExtension
				filename = cyear &cmonth & "_" & custname & "_" & teamname &"_"& categoryname & "_" & title & "_" & medname &"_" & request.cookies("custname")&"_"&mdidx&"."&ext

			filename = replace(filename, "\", "")
			filename = replace(filename, "/", "")
			filename = replace(filename, ":", "")
			filename = replace(filename, "*", "")
			filename = replace(filename, """", "")
			filename = replace(filename, "<", "")
			filename = replace(filename, ">", "")
			filename = replace(filename, "|", "")
			filename = replace(filename, "(", "")
			filename = replace(filename, ")", "")

			sql = "update wb_report_dtl set empid=?, reportname=?, filename=?, cuser=?, cdate=? where mdidx=? and cyear=? and cmonth=? and empid=?"
				cmd.commandText = sql
				cmd.parameters.append cmd.createparameter("empid", adchar, adparaminput, 9)
				cmd.parameters.append cmd.createparameter("reportname", advarchar, adparaminput, 200)
				cmd.parameters.append cmd.createparameter("filename", advarchar, adparaminput, 200)
				cmd.parameters.append cmd.createparameter("cuser", advarchar, adparaminput, 12)
				cmd.parameters.append cmd.createparameter("cdate", addbtimestamp, adparaminput)
				cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput)
				cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4)
				cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2)
				cmd.parameters.append cmd.createparameter("empid2", adchar, adparaminput, 9)
				cmd.parameters("empid").value = request.cookies("userid")
				cmd.parameters("reportname").value = uploadform("file").filename
				cmd.parameters("filename").value = filename
				cmd.parameters("cuser").value = request.cookies("userid")
				cmd.parameters("cdate").value = date
				cmd.parameters("mdidx").value = mdidx
				cmd.parameters("cyear").value = cyear
				cmd.parameters("cmonth").value = cmonth
				cmd.parameters("empid2").value = request.cookies("userid")
			End If
			cmd.execute ,, adexecutenoreords
			uploadform.SaveAs(uploadform.defaultpath&"\"&filename)
		case "D"
			If uploadform.FileExists(uploadform.defaultpath&"\"&orgfile) Then uploadform.DeleteFile(uploadform.defaultpath&"\"&orgfile)
			sql = "delete from wb_report_dtl where mdidx=? and cyear=? and cmonth=? and empid=?"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput)
			cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4)
			cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2)
			cmd.parameters.append cmd.createparameter("empid2", adchar, adparaminput, 9)
			cmd.parameters("mdidx").value = mdidx
			cmd.parameters("cyear").value = cyear
			cmd.parameters("cmonth").value = cmonth
			cmd.parameters("empid2").value = request.cookies("userid")

			cmd.execute ,, adexecutenoreords
			msg = "삭제"
	End Select
	Set cmd = Nothing
	Set uploadform = Nothing

%>
<script type="text/javascript">
<!--
	alert('보고서 파일이 <%=msg%>되었습니다.');
	window.opener.location.replace("/med/sk_med_def.asp?cyear=<%=cyear%>&cmonth=<%=cmonth%>");
	window.close();
//-->
</script>