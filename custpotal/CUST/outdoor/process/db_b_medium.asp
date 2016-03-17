<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '보안관련라이브러리 %>
<%
'	For Each item In request.Form
'		response.write item & " : " & request.Form(item) & "<br>"
'	Next
'	response.End

	Dim atag : atag = ""
	Dim crud : crud = clearXSS(request("crud"), atag)
	Dim monthly : monthly = Replace(request("txtmonthly"), ",", "")
	Dim theme : theme = clearXSS(request("txttheme"), atag)
	Dim cyear : cyear = clearXSS(request("cyear"), atag)
	Dim cmonth : cmonth = clearXSS(request("cmonth"), atag)
	Dim side : side = clearXSS(request("rdoside"), atag)
	Dim orgside : orgside = request("side")
	Dim mdidx : mdidx = clearXSS(request("mdidx"), atag)
	Dim contdix : contidx = request("contidx")
	Dim qty : qty = clearXSS(request("txtqty"), atag)
	Dim standard : standard = request("txtstandard")
	Dim quality : quality = request("cmbquality")
	Dim thmno : thmno = request("hdnthmno")
	Dim expense : expense = Replace(request("txtexpense"), ",", "")
	Dim subexeseq : subexeseq = clearXSS(request("subexeseq"), atag)
	dim no : no = clearXSS(Request("no"), atag)
	if no = "" then no = 1
	Dim intLoop, pcyear, pcmonth, attachFile
	If side = "" Then side=orgside

'
'			response.write "standard : " & standard & "<br>"
'			response.write "quality : " & quality & "<br>"
'			response.write "mdidx : " & mdidx & "<br>"
'			response.write "side : " & side & "<br>"
'			response.write "thmno : " & thmno & "<br>"
'			response.write "subexeseq : " & subexeseq & "<br>"
'			response.write "subexeseqisnull : " & IsNull(subexeseq) & "<br>"
'			response.write "monthly : " & monthly & "<br>"
'			response.write "expense : " & expense & "<br>"
'			response.write "cyear : " & cyear & "<br>"
'			response.write "cmonth : " & cmonth & "<br>"
'
'response.end

'dim item
'for each item in request.form
'	Response.write item & " : " & request.form(item) & "<br>"
'next

	Dim con : Set con = server.CreateObject("adodb.connection")
	con.connectionstring = application("connectionstring")
	con.open

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	Set cmd.activeconnection = con
	cmd.commandType = adCmdText

	Select Case UCase(crud)
		Case "C"

			Dim sql : sql = "select startdate, enddate from wb_contact_mst where contidx = " & request("contidx")
			cmd.commandText = sql
			Dim rs : Set rs = cmd.execute

			' 광고 시작 기점 (시작년도, 시작월)
			Dim startdate : startdate = rs(0)
			Dim enddate : enddate = rs(1)
			Dim eyear : eyear = CStr(Year(enddate))
			Dim emonth
			If Len(Month(enddate)) =  1 Then emonth = "0"&Month(enddate) Else emonth = CStr(Month(enddate))

			Dim term

			' 광고 기간(월) 수
			pstartdate = startdate
			term = DateDiff("m", startdate, enddate)+1

			pcyear = CStr(Year(startdate))
			If Len(Month(startdate)) =  1 Then pcmonth = "0"&Month(startdate) Else pcmonth = CStr(Month(startdate))


			' 1. 광고 매체 면 정보 추가
			sql = "insert into wb_contact_md_dtl (mdidx, cyear, cmonth, side, standard, quality) values (?, ?, ?, ?,?,?)"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput, , mdidx)
			cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4, pcyear)
			cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2, pcmonth)
			cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1, side)
			cmd.parameters.append cmd.createparameter("standard", advarchar, adparaminput, 200, standard)
			cmd.parameters.append cmd.createparameter("quality", advarchar, adparaminput, 200, quality)
			cmd.Execute ,, adExecuteNoRecords
			clearparameter(cmd)
			' 2. 매체 광고비 정보 임시 추가
			' 이 데이터를 기준으로
			sql = "insert into wb_contact_account (mdidx, side, monthly, expense, qty) values (?, ?, ?, ?, ?)"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput, , mdidx)
			cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1, side)
			cmd.parameters.append cmd.createparameter("monthly", adCurrency, adparaminput,,monthly)
			cmd.parameters.append cmd.createparameter("expense", adCurrency, adparaminput,, expense)
			cmd.parameters.append cmd.createparameter("qty", adUnsignedTinyInt, adparaminput,,qty)
			cmd.Execute ,, adExecuteNoRecords
			clearparameter(cmd)
			sql = "insert into wb_contact_exe (mdidx, side, cyear, cmonth , monthly, expense, qty) values (?, ?, ?, ?, ?, ?, ?)"
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput, , mdidx)
			cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1, side)
			cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4, pcyear)
			cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2, pcmonth)
			cmd.parameters.append cmd.createparameter("monthly", adCurrency, adparaminput,,monthly)
			cmd.parameters.append cmd.createparameter("expense", adCurrency, adparaminput,, expense)
			cmd.parameters.append cmd.createparameter("qty", adUnsignedTinyInt, adparaminput,,qty)
			For intLoop = 1 To term
				cyear = CStr(Year(startdate))
				If Len(Month(startdate)) =  1 Then cmonth = "0"&Month(startdate) Else cmonth = CStr(Month(startdate))

				cmd.parameters("mdidx").value = mdidx
				cmd.parameters("side").value = side
				cmd.parameters("cyear").value = cyear
				cmd.parameters("cmonth").value = cmonth
				if term = 1 then
					cmd.parameters("monthly").value = monthly
					cmd.parameters("expense").value = expense
				else
					If (eyear = cyear And emonth = cmonth) And Day(startdate) > 1 Then
						cmd.parameters("monthly").value = 0
						cmd.parameters("expense").value = 0
					Else
						cmd.parameters("monthly").value = monthly
						cmd.parameters("expense").value = expense
					End If
				end if
				cmd.parameters("qty").value = qty

				cmd.commandText = sql
				cmd.Execute ,, adExecuteNoRecords
				startdate = DateAdd("m", 1, startdate)
			Next
			clearparameter(cmd)
			' 3.소재 정보 추가
			sql= "insert into wb_subseq_exe (mdidx, side, cyear, cmonth, thmno, no) values (?, ?, ?, ?, ?, ?)"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput, , mdidx)
			cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1, side)
			cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4, pcyear)
			cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2, pcmonth)
			cmd.parameters.append cmd.createparameter("thmno", adchar, adparaminput, 12, thmno)
			cmd.parameters.append cmd.createparameter("no", adUnsignedTinyInt, adparaminput,,no)
			cmd.Execute ,, adExecuteNoRecords
			clearParameter(cmd)

		Case "U"
			sql = "select startdate, enddate from wb_contact_mst where contidx = " & request("contidx")
			cmd.commandText = sql
			Set rs = cmd.execute

			' 광고 시작 기점 (시작년도, 시작월)
			startdate = rs(0)
			enddate = rs(1)

			' 매체 면 상세 정보 변경 내역 추가
			sql = "insert into wb_contact_md_dtl (mdidx, side, cyear, cmonth, standard, quality) values (?, ?, ? ,?, ?, ?)"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput, ,mdidx)
			cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1, side)
			cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4, cyear)
			cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2, cmonth)
			cmd.parameters.append cmd.createparameter("standard", advarchar, adparaminput, 200, standard)
			cmd.parameters.append cmd.createparameter("quality", advarchar, adparaminput, 200, quality)
			cmd.Execute ,, adExecuteNoRecords
			clearParameter(cmd)
			'메체 면별 소재 정보 등록(소재 내역을 남기기 위해 무조건 소재는 신규로 등록)
'			If subexeseq="" Then
'			Else
'				sql="update wb_subseq_exe set thmno=? where seq =?"
'				cmd.commandText = sql
'				cmd.parameters.append cmd.createparameter("thmno", adchar, adparaminput, 12, thmno)
'				cmd.parameters.append cmd.createparameter("subexeseq", adInteger, adparaminput, , subexeseq)
'			End If
			sql= "insert into wb_subseq_exe (mdidx, side, cyear, cmonth, thmno, no) values (?, ?, ?, ?, ?, ?)"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput, ,mdidx)
			cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1, side)
			cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4, cyear)
			cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2, cmonth)
			cmd.parameters.append cmd.createparameter("thmno", adchar, adparaminput, 12, thmno)
			cmd.parameters.append cmd.createparameter("no", adUnsignedTinyInt, adparaminput,,no)
			cmd.Execute ,, adExecuteNoRecords
			clearParameter(cmd)
			' 매체 면별 광고료 변경
			sql = "update wb_contact_exe set qty=?, monthly =?, expense=? where mdidx=? and side=? and cyear+cmonth >= ?"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("qty", adUnsignedTinyInt, adparaminput,,qty)
			cmd.parameters.append cmd.createparameter("monthly", adCurrency, adparaminput,,monthly)
			cmd.parameters.append cmd.createparameter("expense", adCurrency, adparaminput,, expense)
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput,,mdidx)
			cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1,side)
			cmd.parameters.append cmd.createparameter("yearmon", adChar, adparaminput, 6, cyear+cmonth)
			cmd.Execute ,, adExecuteNoRecords
			clearParameter(cmd)
		Case "D"
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput)
			cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1)
			' 등록 관리 사진 삭제
			sql = "select desc_01, desc_02, desc_03, desc_04 from wb_contact_photo where mdidx=? and side=?"
			cmd.commandText = sql
			cmd.parameters("mdidx").value = mdidx
			cmd.parameters("side").value = side
			Set rs = cmd.execute

			Dim fso : Set fso = server.CreateObject("scripting.filesystemobject")
			Do Until rs.eof
				For intLoop = 0 To 3
					attachFile = "\\11.0.12.201\adportal\media\"&rs(intLoop)
					if fso.fileexists(attachFile) then
						fso.deletefile(attachFile)
					end If
				Next
				rs.movenext
			Loop
			'소재 사진 삭제
			sql = "delete from wb_contact_photo where mdidx=? and side=?"
			cmd.commandText = sql
			cmd.parameters("mdidx").value = mdidx
			cmd.parameters("side").value = side
			cmd.Execute ,, adExecuteNoRecords
			' 소재 집행 내역 삭제
			sql="delete from wb_subseq_exe where mdidx=? and side=?"
			cmd.commandText = sql
			cmd.parameters("mdidx").value = mdidx
			cmd.parameters("side").value = side
			cmd.Execute ,, adExecuteNoRecords
			' 집행 내역 삭제
			sql = "delete from wb_contact_exe where mdidx=? and side=?"
			cmd.commandText = sql
			cmd.parameters("mdidx").value = mdidx
			cmd.parameters("side").value = side
			cmd.Execute ,, adExecuteNoRecords
			' 매체 면 정보 삭제
			sql = "delete from wb_contact_md_dtl where mdidx=? and side=?"
			cmd.commandText = sql
			cmd.parameters("mdidx").value = mdidx
			cmd.parameters("side").value = side
			cmd.Execute ,, adExecuteNoRecords
			' 매체 기초 정보 삭제
			sql = "delete from wb_contact_account where mdidx=? and side=?"
			cmd.commandText = sql
			cmd.parameters("mdidx").value = mdidx
			cmd.parameters("side").value = side
			cmd.Execute ,, adExecuteNoRecords
			clearParameter(cmd)


	End select
	clearParameter(cmd)

	cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput)
	sql = "select m.contidx from wb_contact_mst  m inner join wb_contact_md m2 on m.contidx = m2.contidx where m2.mdidx =?"
	cmd.commandText = sql
	cmd.parameters("mdidx").value = mdidx
	Set rs = cmd.Execute
	Dim contidx : contidx = rs(0)
	clearParameter(cmd)

	cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput)
	cmd.parameters.append cmd.createparameter("contidx", adinteger, adparaminput)
	sql= "update wb_contact_mst set totalprice = (select isnull(sum(monthly),0) from wb_contact_exe where mdidx = ?) where contidx  = ?"
	cmd.commandText = sql
	cmd.parameters("mdidx").value = mdidx
	cmd.parameters("contidx").value = contidx
	cmd.Execute ,, adExecuteNoRecords
	clearParameter(cmd)

	Set cmd = Nothing

%>
<script type="text/javascript">
<!--
	window.opener.getcontact();
	window.close();
//-->
</script>