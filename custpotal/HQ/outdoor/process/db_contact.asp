<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>
<%
'	dim item
'	for each item in request.form
'		Response.write item &  " :" & request.form(item) & "<br>"
'	next


	Dim atag : atag = ""
	Dim crud : crud = request("crud")
	Dim contidx : contidx = request("contidx")
	Dim cyear : cyear = clearXSS(request("cyear"), atag)
	Dim cmonth : cmonth = clearXSS(request("cmonth"), atag)
	Dim custcode : custcode = clearXSS(request("cmbcustcode"), atag)
	Dim teamcode : teamcode = clearXSS(request("cmbteamcode"), atag)
	Dim title : title = request("txttitle")
	Dim firstdate : firstdate = request("txtfirstdate")
	Dim startdate : startdate = request("txtstartdate")


	Dim enddate : enddate = request("txtenddate")
	firstdate = Left(firstdate,4)&"-"&Mid(firstdate,5,2)&"-"&Right(firstdate,2)
	startdate = Left(startdate,4)&"-"&Mid(startdate,5,2)&"-"&Right(startdate,2)
	enddate = Left(enddate,4)&"-"&Mid(enddate,5,2)&"-"&Right(enddate,2)
	If UCase(crud) <> "D" Then
	If Not IsDate(firstdate)  Then response.write "<script> alert('날짜형식이 올바르지 않습니다.'); parent.document.forms[0].txtfirstdate.select(); </script>"
	If Not IsDate(startdate)  Then response.write "<script> alert('날짜형식이 올바르지 않습니다.'); parent.document.forms[0].txtstartdate.select(); </script>"
	If Not IsDate(enddate)  Then response.write "<script> alert('날짜형식이 올바르지 않습니다.'); parent.document.forms[0].txtenddate.select(); </script>"
	End If
	Dim comment : comment = request("txtcomment")
	Dim regionmemo : regionmemo = request("txtregionmemo")
	Dim mediummemo : mediummemo = request("txtmediummemo")
	Dim flag : flag = clearXSS(request("rdoflag"), atag)
	Dim cuser : cuser = request.cookies("userid")
	Dim orgcustcode : orgcustcode = request("orgcustcode")
	Dim orgteamcode : orgteamcode = request("orgteamcode")
	Dim uuser : uuser = Null
	Dim udate : udate = Null
	Dim sql2
	Dim rs, rs2, rs3, rs4
	Dim tempstartdate
	Dim nmdidx
	Set rs3 = server.CreateObject("adodb.recordset")
	rs3.activeconnection =application("connectionstring")
	rs3.cursorlocation = adUseClient
	rs3.cursorType = adOpenStatic
	rs3.lockType = adLockOptimistic

	Dim dateComp
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandType = adcmdText
	Dim cmd2 : Set cmd2 = server.CreateObject("adodb.command")
	cmd2.activeconnection = application("connectionstring")
	cmd2.commandType = adCmdText

	Select Case UCase(crud)
		Case "C"
		 dateComp = DateDiff("m", startdate, enddate)+1
			Dim sql : sql = "insert into dbo.wb_contact_mst values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)"

			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("custcode", adchar, adparaminput, 6, teamcode)
			cmd.parameters.append cmd.createparameter("title", advarchar, adparaminput, 200, title)
			cmd.parameters.append cmd.createparameter("firstdate", adDBTimeStamp, adparaminput, 4, firstdate)
			cmd.parameters.append cmd.createparameter("startdate", adDBTimeStamp, adparaminput, 4, startdate)
			cmd.parameters.append cmd.createparameter("enddate", adDBTimeStamp, adparaminput, 4, enddate)
			cmd.parameters.append cmd.createparameter("comment", adLongVarChar, adparaminput, 2147483647, comment)
			cmd.parameters.append cmd.createparameter("regionmemo", adLongVarChar, adparaminput, 2147483647, regionmemo)
			cmd.parameters.append cmd.createparameter("mediummemo", adLongVarChar, adparaminput, 2147483647, mediummemo)
			cmd.parameters.append cmd.createparameter("flag", adChar, adparaminput, 1, flag)
			cmd.parameters.append cmd.createparameter("totalprice", adCurrency, adparaminput, 8, 0)
			cmd.parameters.append cmd.createparameter("cuser", adVarChar, adparaminput, 12, cuser)
			cmd.parameters.append cmd.createparameter("cdate", adDBTimeStamp, adparaminput, 4, Date)
			cmd.parameters.append cmd.createparameter("uuser", adVarChar, adparaminput, 12, uuser)
			cmd.parameters.append cmd.createparameter("udate", adDBTimeStamp, adparaminput, 4, udate)
			cmd.parameters.append cmd.createparameter("oldcustcode", adchar, adparaminput, 6, null)
			cmd.parameters.append cmd.createparameter("type", adWchar, adparaminput, 2, "신규")
			cmd.Execute ,, adExecuteNoRecords
		Case "U"
			' 저장된 종료일자
			 dateComp = DateDiff("m", startdate, enddate)+1
			sql ="select enddate, flag from wb_contact_mst where contidx=?"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("contidx", adInteger, adParamInput)
			cmd.parameters("contidx").value = contidx
			Set rs = cmd.Execute
			Dim realenddate : realenddate = rs(0)
			flag = rs(1)
			rs.close
			Set rs = Nothing
			Dim nextDate
			sql = "select a.mdidx, a.side, a.monthly, a.expense, a.qty from wb_contact_account a inner join (select distinct c.mdidx, c.side from wb_contact_md b inner join wb_contact_md_dtl c on b.mdidx=c.mdidx where contidx = ?) as d on a.mdidx=d.mdidx and a.side=d.side"
			cmd.commandText = sql
			cmd.parameters("contidx").value = contidx
			Set rs = cmd.execute
			clearparameter(cmd)

			If Not rs.eof Then
			Dim datComp : datComp = DateDiff("m", realenddate, enddate)
				If datComp > 0 Then ' 종료일이 연장
					sql = "insert into wb_contact_exe(cyear, cmonth, mdidx, side, monthly, expense, qty, isHold, uuser, udate) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
					cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4)
					cmd.parameters.append cmd.createparameter("cmonth", adchar, adParaminput, 2)
					cmd.parameters.append cmd.createparameter("mdidx", adInteger, adParamInput)
					cmd.parameters.append cmd.createparameter("side", adChar, adParamInput, 1)
					cmd.parameters.append cmd.createParameter("monthly", adCurrency, adParaminput)
					cmd.parameters.append cmd.createParameter("expense", adCurrency, adParaminput)
					cmd.parameters.append cmd.createParameter("qty", adInteger, adParaminput)
					cmd.parameters.append cmd.createParameter("isHold", adChar, adParaminput, 1)
					cmd.parameters.append cmd.createParameter("uuser", adVarChar, adParaminput, 12)
					cmd.parameters.append cmd.createParameter("udate", adDBTimeStamp, adParaminput)

					Do Until rs.eof
						For intLoop = 1 To datComp
							nextDate = DateAdd("m", intLoop, realenddate)
							cmd.parameters("cyear").value = Year(nextDate)
							cmd.parameters("cmonth").value = setmonth(Month(nextDate))
							cmd.parameters("mdidx").value = rs("mdidx")
							cmd.parameters("side").value = rs("side")
							cmd.parameters("monthly").value = rs("monthly")
							cmd.parameters("expense").value = rs("expense")
							cmd.parameters("qty").value = rs("qty")
							cmd.parameters("isHold").value = Null
							cmd.parameters("uuser").value = session("userid")
							cmd.parameters("udate").value = Date
							cmd.commandText = sql
							cmd.execute ,, adexecutenorecords
						Next
						rs.movenext
					Loop
					clearparameter(cmd)
				ElseIf datComp < 0 Then '종료일 축소

					Dim yearmon :
					if Len(month(enddate)) = 1 then
						yearmon = Year(enddate)&"0"&Month(enddate)
					else
						yearmon = Year(enddate)&Month(enddate)
					end if
					cmd.parameters.append cmd.createparameter("contidx", adInteger, adParamInput)
					cmd.parameters.append cmd.createparameter("yearmon", adChar, adParamInput, 6)
					sql = "select count(isHold) from wb_contact_md a inner join wb_contact_exe b on a.mdidx=b.mdidx where a.contidx = ? and cyear+cmonth > ? "
					cmd.commandText = sql
					cmd.parameters("contidx").value = contidx
					cmd.parameters("yearmon").value = yearmon
					Set rs2 = cmd.execute
					If rs2(0)  > 0 Then
						response.write "<script> alert('정산된 년월이 존재합니다.\n\n계약을 수정할 수 없습니다.'); </script>"
						response.End
					End If
					rs2.close
					Set rs2 = nothing
					clearparameter(cmd)
					cmd.parameters.append cmd.createparameter("mdidx", adInteger, adParamInput)
					cmd.parameters.append cmd.createparameter("yearmon", adChar, adParamInput, 6)
					do until rs.eof
						sql = "delete from wb_contact_exe where mdidx =? and cyear+cmonth > ?"
						cmd.parameters("mdidx").value = rs("mdidx")
						cmd.parameters("yearmon").value = cstr(yearmon)
						cmd.commandText = sql
						cmd.execute ,,adExecuteNoRecords
						sql = "delete from wb_subseq_exe where mdidx=? and cyear+cmonth > ?"
						cmd.commandText = sql
						cmd.execute ,,adExecuteNoRecords
						sql = "delete from wb_contact_md_dtl where mdidx=? and cyear+cmonth > ?"
						cmd.commandText = sql
						cmd.execute ,,adExecuteNoRecords

						rs.movenext
					Loop
				Else	'변화 없음

				End If
			End If
			rs.close
			clearparameter(cmd)
			sql = "select sum(monthly) from wb_contact_exe a inner join wb_contact_md b on a.mdidx=b.mdidx where b.contidx=?"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("contidx", adInteger, adParamInput)
			cmd.parameters("contidx").value = contidx
			Set rs = cmd.execute
			Dim totalprice : totalprice = rs(0)
			clearparameter(cmd)

			sql = "update wb_contact_mst set custcode=?, title=?, firstdate=?, startdate=?,  enddate=?, comment=?, regionmemo=?, mediummemo=?, totalprice=?, uuser=?, udate=?, flag=? where contidx=?"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("custcode", adchar, adparaminput, 6, teamcode)
			cmd.parameters.append cmd.createparameter("title", advarchar, adparaminput, 200, title)
			cmd.parameters.append cmd.createparameter("firstdate", adDBTimeStamp, adparaminput, 4, firstdate)
			cmd.parameters.append cmd.createparameter("startdate", adDBTimeStamp, adparaminput, 4, startdate)
			cmd.parameters.append cmd.createparameter("enddate", adDBTimeStamp, adparaminput, 4, enddate)
			cmd.parameters.append cmd.createparameter("comment", adLongVarChar, adparaminput, 214748, comment)
			cmd.parameters.append cmd.createparameter("regionmemo", adLongVarChar, adparaminput, 214748, regionmemo)
			cmd.parameters.append cmd.createparameter("mediummemo", adLongVarChar, adparaminput, 214748, mediummemo)
			cmd.parameters.append cmd.createparameter("totalprice", adCurrency, adparaminput, 8, totalprice)
			cmd.parameters.append cmd.createparameter("uuser", adVarChar, adparaminput, 12, session("userid"))
			cmd.parameters.append cmd.createparameter("udate", adDBTimeStamp, adparaminput, 4, date)
			cmd.parameters.append cmd.createparameter("flag", adChar, adparaminput, 1, flag)
			cmd.parameters.append cmd.createparameter("contidx", adInteger, adParamInput,,contidx)

			cmd.Execute ,, adExecuteNoRecords
		Case "E"
			' 재 계약 입력

			dateComp = DateDiff("m", startdate, enddate)+1
			sql ="select flag from wb_contact_mst where contidx=?"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("contidx", adInteger, adParamInput)
			cmd.parameters("contidx").value = contidx
			Set rs = cmd.Execute
			flag = rs(0)
			clearparameter(cmd)

'			response.write flag
'			response.end
			rs.close

			sql = "insert into wb_contact_mst(custcode, title, firstdate, startdate, enddate, comment, regionmemo, mediummemo, flag, totalprice, cuser, cdate, uuser, udate,type) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) "
			cmd.commandText = sql

			cmd.parameters.append cmd.createparameter("custcode", adchar, adparaminput, 6)
			cmd.parameters.append cmd.createparameter("title", advarchar, adparaminput, 200)
			cmd.parameters.append cmd.createparameter("firstdate", adDBTimeStamp, adparaminput)
			cmd.parameters.append cmd.createparameter("startdate", adDBTimeStamp, adparaminput)
			cmd.parameters.append cmd.createparameter("enddate", adDBTimeStamp, adparaminput)
			cmd.parameters.append cmd.createparameter("comment", adLongVarChar, adparaminput, 2147483647)
			cmd.parameters.append cmd.createparameter("regionmemo", adLongVarChar, adparaminput, 2147483647)
			cmd.parameters.append cmd.createparameter("mediummemo", adLongVarChar, adparaminput, 2147483647)
			cmd.parameters.append cmd.createparameter("flag", adChar, adparaminput, 1)
			cmd.parameters.append cmd.createparameter("totalprice", adCurrency, adparaminput)
			cmd.parameters.append cmd.createparameter("cuser", adVarChar, adparaminput, 12)
			cmd.parameters.append cmd.createparameter("cdate", adDBTimeStamp, adparaminput)
			cmd.parameters.append cmd.createparameter("uuser", adVarChar, adparaminput, 12)
			cmd.parameters.append cmd.createparameter("udate", adDBTimeStamp, adparaminput)
			cmd.parameters.append cmd.createparameter("type", adWChar, adparaminput, 2)


			cmd.parameters("custcode").value = teamcode
			cmd.parameters("title").value = title
			cmd.parameters("firstdate").value = firstdate
			cmd.parameters("startdate").value = startdate
			cmd.parameters("enddate").value = enddate
			cmd.parameters("comment").value = comment
			cmd.parameters("regionmemo").value = regionmemo
			cmd.parameters("mediummemo").value = mediummemo
			cmd.parameters("flag").value = flag
			cmd.parameters("totalprice").value = 0
			cmd.parameters("cuser").value = Request.cookies("userid")
			cmd.parameters("cdate").value = Date
			cmd.parameters("uuser").value = uuser
			cmd.parameters("udate").value = udate
			cmd.parameters("type").value = "연장"

			'신규 계약 번호
			cmd.execute ,, adExecuteNoRecords
			clearparameter(cmd)

			sql = "select @@identity from wb_contact_mst"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("contidx", adInteger, adParamInput)
			cmd.parameters("contidx").value = contidx
			set rs = cmd.execute
			ncontidx =  rs(0)
			clearparameter(cmd)


			sql = "select * from wb_contact_md where contidx=?"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("contidx", adInteger, adParamInput)
			cmd.parameters("contidx").value = contidx
			set rs = cmd.execute
			clearparameter(cmd)

			if not rs.eof then
				do until rs.eof
					'기존에 등록된 매체 정보를 가져와 등록한다
					sql = "insert into wb_contact_md values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
					cmd.commandText = sql
					cmd.parameters.append cmd.createparameter("categoryidx", adInteger, adParamInput)
					cmd.parameters.append cmd.createparameter("region", advarchar, adParamInput, 20)
					cmd.parameters.append cmd.createparameter("locate", advarchar, adParamInput, 200)
					cmd.parameters.append cmd.createparameter("unit", advarchar, adParamInput, 10)
					cmd.parameters.append cmd.createparameter("map", advarchar, adParamInput,100)
					cmd.parameters.append cmd.createparameter("medcode", adchar, adParamInput, 6)
					cmd.parameters.append cmd.createparameter("trust", advarchar, adParamInput, 10)
					cmd.parameters.append cmd.createparameter("medclass", advarchar, adParamInput, 50)
					cmd.parameters.append cmd.createparameter("validclass", advarchar, adParamInput, 50)
					cmd.parameters.append cmd.createparameter("contidx", adInteger, adParamInput)
					cmd.parameters.append cmd.createparameter("empid", adchar, adParamInput,9)
					cmd.parameters.append cmd.createparameter("cuser", advarchar, adParamInput, 12)
					cmd.parameters.append cmd.createparameter("cdate", adDBTimeStamp, adParamInput)
					cmd.parameters.append cmd.createparameter("uuser", advarchar, adParamInput, 12)
					cmd.parameters.append cmd.createparameter("udate", adDBTimeStamp, adParamInput)
					cmd.parameters.append cmd.createparameter("oldmedcode", adchar, adParamInput, 6)
					cmd.parameters("categoryidx").value = rs("categoryidx")
					cmd.parameters("region").value = rs("region")
					cmd.parameters("locate").value =  rs("locate")
					cmd.parameters("unit").value =  rs("unit")
					cmd.parameters("map").value =  rs("map")
					cmd.parameters("medcode").value =  rs("medcode")
					cmd.parameters("trust").value =  rs("trust")
					cmd.parameters("medclass").value = rs("medclass")
					cmd.parameters("validclass").value =  rs("validclass")
					cmd.parameters("contidx").value = ncontidx
					cmd.parameters("empid").value = rs("empid")
					cmd.parameters("cuser").value = request.cookies("userid")
					cmd.parameters("cdate").value = date
					cmd.parameters("uuser").value = uuser
					cmd.parameters("udate").value = udate
					cmd.parameters("oldmedcode").value = null
					cmd.execute ,, adExecuteNoRecords
					clearparameter(cmd)

					sql ="select max(mdidx) from wb_contact_md where contidx=" & ncontidx
					cmd.commandText = sql
					set rs2 = cmd.execute
					nmdidx = rs2(0)
					clearparameter(cmd)
					rs2.close

					dim nyear
					dim nmonth
					nyear = Year(startdate)

					if Len(month(startdate)) = 1 then
						nmonth = "0"&month(startdate)
					else
						nmonth = cstr(month(startdate))
					end if

					sql = "select * from wb_contact_md_dtl where mdidx=?"
					cmd.commandText = sql
					cmd.parameters.append cmd.createparameter("mdidx", adInteger, adParamInput)
					cmd.parameters("mdidx").value = rs("mdidx")
					set rs4 = cmd.execute
					clearparameter(cmd)

					Dim strmdidx : strmdidx = rs("mdidx")


					if not rs4.eof then
						do until rs4.eof
							sql = "insert into wb_contact_md_dtl values (?, ?, ?,?, ?, ?)"
							cmd.commandText = sql
							cmd.parameters.append cmd.createparameter("mdidx", adInteger, adParamInput)
							cmd.parameters.append cmd.createparameter("side", adchar, adParamInput, 1)
							cmd.parameters.append cmd.createparameter("cyear", adchar, adParamInput, 4)
							cmd.parameters.append cmd.createparameter("cmonth", adchar, adParamInput, 2)
							cmd.parameters.append cmd.createparameter("standard", advarchar, adParamInput,400)
							cmd.parameters.append cmd.createparameter("quality", advarchar, adParamInput, 400)

							cmd.parameters("mdidx").value = nmdidx
							cmd.parameters("side").value = rs4("side")
							cmd.parameters("cyear").value =  nyear
							cmd.parameters("cmonth").value = nmonth
							cmd.parameters("standard").value =  rs4("standard")
							cmd.parameters("quality").value =  rs4("quality")
							cmd.execute ,, adExecuteNoRecords
							clearparameter(cmd)

							rs4.movenext
						Loop
					end if
					rs4.close

					sql="insert wb_subseq_exe "
					sql = sql & " select "&nmdidx&", side, thmno, null, 1, null, '"&nyear&"', '"&nmonth&"' from wb_subseq_exe where seq in (select max(seq) from wb_subseq_exe where mdidx="&rs("mdidx")&" group by side, cyear, cmonth)"
					cmd.commandText = sql
					cmd.execute ,, adExecuteNoRecords

					sql = "insert wb_contact_account "
					sql = sql & " select "&nmdidx&", side, monthly, expense, qty from wb_contact_account where mdidx =" & rs("mdidx")
					cmd.commandText = sql
					cmd.execute ,, adExecuteNoRecords

					sql="insert wb_contact_photo "
					sql = sql & " select  '" & nyear & "', '" & nmonth &"' , "&nmdidx&", side, pht_01, pht_02, pht_03, pht_04, desc_01, desc_02, desc_03, desc_04, null, null, null from wb_contact_photo where seq in (select max(seq) from wb_contact_photo where mdidx="&rs("mdidx")&" group by side, cyear, cmonth)"
					cmd.commandText = sql
					cmd.execute ,, adExecuteNoRecords

					sql="insert wb_report_photo "
					sql = sql & " select "&ncontidx&",  '" & nyear & "', '" & nmonth &"', photo1, photo2, photo3, photo4  from wb_report_photo where seq in (select max(seq) from wb_report_photo where contidx="&contidx&" group by cyear, cmonth)"
					cmd.commandText = sql
					cmd.execute ,, adExecuteNoRecords

					dim tmp_date
					tmp_date = startdate


					if dateComp = 1 then
						nyear = year(startdate)
						if Len(month(startdate)) = 1 then
							nmonth = "0"&month(startdate)
						else

							nmonth = month(stratdate)
						end if

						sql = "insert into wb_contact_exe "
						sql = sql & " select '" & nyear & "', '" & nmonth &"' , "&nmdidx&", side, monthly, expense, qty, null, '" & request.cookies("userid") & "', '" & date & "' from wb_contact_account where mdidx = " & nmdidx
						response.write sql
						cmd.commandText = sql
						cmd.execute ,, adExecuteNoRecords
					else
						for intLoop = 1 to dateComp
							nyear = year(tmp_date)
							if Len(month(tmp_date)) = 1 then
								nmonth = "0"&month(tmp_date)
							else
								nmonth = month(tmp_date)
							end if


							if tmp_date > enddate then
								sql = "insert into wb_contact_exe "
								sql = sql & " select '" & nyear & "', '" & nmonth &"' , "&nmdidx&", side, 0, 0, 0, null, '" & request.cookies("userid") & "', '" & date & "' from wb_contact_account where mdidx = " & nmdidx
							else
								sql = "insert into wb_contact_exe "
								sql = sql & " select '" & nyear & "', '" & nmonth &"' , "&nmdidx&", side, monthly, expense, qty, null, '" & request.cookies("userid") & "', '" & date & "' from wb_contact_account where mdidx = " & nmdidx
							end if
							response.write sql
							cmd.commandText = sql
							cmd.execute ,, adExecuteNoRecords
							tmp_date = dateadd("m", 1, tmp_date)
						next
					end if

				rs.movenext
				Loop
			end if
			rs.close
			clearParameter(cmd)

			cmd.parameters.append cmd.createparameter("contidx", adinteger, adparaminput)
			cmd.parameters.append cmd.createparameter("contidx2", adinteger, adparaminput)
			sql= "update wb_contact_mst set totalprice = (select sum(monthly) from wb_contact_exe a inner join wb_contact_md b on a.mdidx=b.mdidx where b.contidx=?) where contidx=?"

			cmd.commandText = sql
			cmd.parameters("contidx").value = ncontidx
			cmd.parameters("contidx2").value = ncontidx
			cmd.Execute ,, adExecuteNoRecords
			clearParameter(cmd)
			Set rs = Nothing
			response.write "<script>parent.opener.document.location.reload();</script>"
'	dim item
'	for each item in request.form
'		Response.write request.form(item) & "<br>"
'	next

		Case "D"
			sql = "select mdidx from wb_contact_md where contidx = ?"
			cmd.parameters.append cmd.createparameter("contidx", adinteger, adparaminput)
			cmd.parameters("contidx").value = contidx
			cmd.commandText = sql
			Set rs = cmd.execute

			If rs.eof Then
				sql = "select count(mdidx) from wb_contact_exe where mdidx in (select mdidx from wb_contact_md where contidx=?) and isHold is not null"
				cmd.commandText = sql
				cmd.parameters("contidx").value = contidx
				Set rs2 = cmd.execute

				If rs2(0) = 0 Then
					sql = "delete from wb_contact_mst where contidx=?"
					cmd.parameters("contidx").value = contidx
					cmd.commandText = sql
					cmd.Execute ,, adExecuteNoRecords
					response.write "<script> parent.location.replace('/hq/outdoor/list_contact.asp?cmbcustcode="&orgcustcude&"&cmbteamcode="&orgteamcode&"&cyear="&cyear&"&cmonth="&cmonth&"');</script>"
				Else
					response.write "<script> alert('정산완료 또는 정산중인 계약은 삭제할 수 없습니다.');</script>"
					response.end
				End If
			Else
				response.write "<script> alert('매체정보가 등록된 계약은 삭제할 수 없습니다.\n\n계약된 매체정보를 먼저 삭제하세요');</script>"
				response.end
			End If
			rs.close
			Set rs = Nothing
			Set cmd = Nothing


	End Select
%>
<script language="JavaScript">
<!--
	window.parent.opener.location.replace("/hq/outdoor/list_contact.asp?cmbcustcode=<%=orgcustcode%>&cmbteamcode=<%=orgteamcode%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>");
	var crud = "<%=lcase(crud)%>";
	if (crud != 'c') parent.close();
	self.focus();
//-->
</script>