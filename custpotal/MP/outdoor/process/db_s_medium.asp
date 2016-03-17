<!--#include virtual="/mp/outdoor/inc/Function.asp" -->
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

	Dim sql : sql = "select startdate, enddate from wb_contact_mst where contidx = ?"
	cmd.commandText = sql
	cmd.parameters.append cmd.createparameter("contidx", adinteger, adparaminput)
	cmd.parameters("contidx").value = contidx
	Dim rs : Set rs = cmd.execute
	clearparameter(cmd)
	' 광고 시작 기점 (시작년도, 시작월)
	Dim startdate : startdate = rs(0)
	Dim enddate : enddate = rs(1)
	rs.close

	Select Case UCase(crud)
		Case "C"
			Dim eyear : eyear = CStr(Year(enddate))
			Dim emonth
			If Len(Month(enddate)) =  1 Then emonth = "0"&Month(enddate) Else emonth = CStr(Month(enddate))

			Dim term

			' 광고 기간(월) 수
			pstartdate = startdate
			term = DateDiff("m", startdate, enddate)+1
			dim dterm : dterm = DateDiff("d", startdate, enddate) + 1


			pcyear = CStr(Year(startdate))
			If Len(Month(startdate)) =  1 Then pcmonth = "0"&Month(startdate) Else pcmonth = CStr(Month(startdate))

			'0 매체 정보 추가
			sql ="select mdidx, categoryidx, region, locate, unit, map, medcode, trust, medclass, validclass, contidx, empid, cuser, cdate, uuser, udate from wb_contact_md where contidx = " & contidx
			rs.cursorlocation = aduseclient
			rs.cursorType = adOenStatic
			rs.lockType = adLockOptimistic
			rs.source = sql
			rs.open

			rs.addnew
				rs("categoryidx").value = categoryidx
				rs("region").value = region
				rs("locate").value = locate
				rs("unit").value = unit
				rs("map").value = null
				rs("medcode").value = medcode
				rs("trust").value = trust
				rs("medclass").value = null
				rs("validclass").value = Null
				rs("contidx").value = contidx
				rs("empid").value = empid
				rs("cuser").value = request.cookies("userid")
				rs("cdate").value = date
				rs("uuser").value = null
				rs("udate").value = Null
			rs.update

			mdidx = rs("mdidx")
			rs.close


			' 1. 광고 매체 면 정보 추가
			sql = "insert into wb_contact_md_dtl (mdidx, side, cyear, cmonth, standard, quality) values (?, ?, ?, ?, ?, ?)"
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
			cmd.parameters("cmonth").value = pcmonth
			cmd.parameters("standard").value = standard
			cmd.parameters("quality").value = quality
			cmd.Execute ,, adExecuteNoRecords
			clearparameter(cmd)

			' 2. 매체 광고비 정보 임시 추가
			' 이 데이터를 기준으로

			sql = "insert into wb_contact_account (mdidx, side, monthly, expense, qty) values (?, ?, ?, ?, ?)"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput)
			cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1)
			cmd.parameters.append cmd.createparameter("monthly", adCurrency, adparaminput)
			cmd.parameters.append cmd.createparameter("expense", adCurrency, adparaminput)
			cmd.parameters.append cmd.createparameter("qty", adinteger, adparaminput)
			cmd.parameters("mdidx").value = mdidx
			cmd.parameters("side").value = "F"
			cmd.parameters("monthly").value = monthly
			cmd.parameters("expense").value = expense
			cmd.parameters("qty").value = qty
			cmd.Execute ,, adExecuteNoRecords
			clearparameter(cmd)

			sql = "insert into wb_contact_exe (mdidx, side, cyear, cmonth , monthly, expense, qty, isHold, uuser, udate) values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput )
			cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1)
			cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4)
			cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2)
			cmd.parameters.append cmd.createparameter("monthly", adCurrency, adparaminput)
			cmd.parameters.append cmd.createparameter("expense", adCurrency, adparaminput)
			cmd.parameters.append cmd.createparameter("qty", adinteger, adparaminput)
			cmd.parameters.append cmd.createparameter("isHold", adchar, adparaminput, 1)
			cmd.parameters.append cmd.createparameter("uuser", advarchar, adparaminput, 12)
			cmd.parameters.append cmd.createparameter("udate", adDBTimeStamp, adparaminput)

			For intLoop = 1 To term
				cyear = CStr(Year(startdate))
				If Len(Month(startdate)) =  1 Then cmonth = "0"&Month(startdate) Else cmonth = CStr(Month(startdate))

				cmd.parameters("mdidx").value = mdidx
				cmd.parameters("side").value = "F"
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
				cmd.parameters("isHold").value = Null
				cmd.parameters("uuser").value = session("userid")
				cmd.parameters("udate").value = Date

				cmd.Execute ,, adExecuteNoRecords
				startdate = DateAdd("m", 1, startdate)
			Next
			clearparameter(cmd)
			' 3.소재 정보 추가

			sql= "insert into wb_subseq_exe (mdidx, side, cyear, cmonth, thmno, no) values (?, ?, ?, ?, ?, ?)"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput )
			cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1)
			cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4)
			cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2)
			cmd.parameters.append cmd.createparameter("thmno", adchar, adparaminput, 12)
			cmd.parameters.append cmd.createparameter("no", adUnsignedTinyInt, adparaminput)
			cmd.parameters("mdidx").value = mdidx
			cmd.parameters("side").value = "F"
			cmd.parameters("cyear").value = pcyear
			cmd.parameters("cmonth").value = pcmonth
			cmd.parameters("thmno").value = thmno
			cmd.parameters("no").value = 1
			cmd.Execute ,, adExecuteNoRecords
			clearParameter(cmd)

		Case "U"
		'정산 중 또는 정산 완료된 경우 수정은 어디까지?

			' 매체 정보 변경
			sql = "update wb_contact_md set categoryidx=?, region=?, locate=?, unit=?, medcode=?, trust=?, empid=? where mdidx=?"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("categoryidx", adinteger, adparaminput, , categoryidx)
			cmd.parameters.append cmd.createparameter("region", advarchar, adparaminput, 10, region)
			cmd.parameters.append cmd.createparameter("locate", advarchar, adparaminput, 200, locate)
			cmd.parameters.append cmd.createparameter("unit", advarchar, adparaminput, 10, unit)
			cmd.parameters.append cmd.createparameter("medcode", adchar, adparaminput, 6, medcode)
			cmd.parameters.append cmd.createparameter("trust", advarchar, adparaminput, 10, trust)
			cmd.parameters.append cmd.createparameter("empid", adchar, adparaminput, 9, empid)
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput, , mdidx)
			cmd.Execute ,, adExecuteNoRecords
			clearParameter(cmd)


			' 매체 면 상세 정보 변경

			sql = "insert into wb_contact_md_dtl (mdidx, cyear, cmonth, side, standard, quality) values (?, ?, ?, ?,?,?)"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput, , mdidx)
			cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4, cyear)
			cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2, cmonth)
			cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1, "F")
			cmd.parameters.append cmd.createparameter("standard", advarchar, adparaminput, 200, standard)
			cmd.parameters.append cmd.createparameter("quality", advarchar, adparaminput, 200, quality)
			cmd.Execute ,, adExecuteNoRecords
			clearparameter(cmd)

'			sql  = "update wb_contact_md_dtl set standard = ?, quality = ? where mdidx = ? "
'			cmd.commandText = sql
'			cmd.parameters.append cmd.createparameter("standard", advarchar, adparaminput, 200, standard)
'			cmd.parameters.append cmd.createparameter("quality", advarchar, adparaminput, 200, quality)
'			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput, ,mdidx)
'			cmd.Execute ,, adExecuteNoRecords
'			clearParameter(cmd)

			'메체 면별 소재 정보 등록
'			If subseq="" Then
				sql= "insert into wb_subseq_exe (mdidx, side, cyear, cmonth, thmno, no) values (?, ?, ?, ?, ?,?)"
				cmd.commandText = sql
				cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput, ,mdidx)
				cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1, "F")
				cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4, cyear)
				cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2, cmonth)
				cmd.parameters.append cmd.createparameter("thmno", adchar, adparaminput, 12, thmno)
			cmd.parameters.append cmd.createparameter("no", adUnsignedTinyInt, adparaminput,,1)
'			Else
'				sql="update wb_subseq_exe set thmno=? where mdidx =? and cyear=? and cmonth=?"
'				cmd.commandText = sql
'				cmd.parameters.append cmd.createparameter("thmno", adchar, adparaminput, 12, thmno)
'				cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput, ,mdidx)
'				cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4, cyear)
'				cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2, cmonth)
'			End If
			cmd.Execute ,, adExecuteNoRecords
			clearParameter(cmd)
			' 매체별 광고료 변경

'			sql = "insert into wb_contact_account (mdidx, side, monthly, expense, qty) values (?, ?, ?, ?, ?)"
'			cmd.commandText = sql
'			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput, , mdidx)
'			cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1, "F")
'			cmd.parameters.append cmd.createparameter("monthly", adCurrency, adparaminput,,monthly)
'			cmd.parameters.append cmd.createparameter("expense", adCurrency, adparaminput,, expense)
'			cmd.parameters.append cmd.createparameter("qty", adUnsignedTinyInt, adparaminput,,qty)
'			cmd.Execute ,, adExecuteNoRecords
'			clearparameter(cmd)

			sql = "update wb_contact_exe set qty=?, monthly =?, expense=? where mdidx=? and  cyear+cmonth >= ?"
'			sql = "update wb_contact_exe set qty=?, monthly =?, expense=? where mdidx=? and cyear=? and cmonth=?"
			cmd.commandText = sql
			cmd.parameters.append cmd.createparameter("qty", adinteger, adparaminput,,qty)
			cmd.parameters.append cmd.createparameter("monthly", adCurrency, adparaminput,,monthly)
			cmd.parameters.append cmd.createparameter("expense", adCurrency, adparaminput,, expense)
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput,,mdidx)
			cmd.parameters.append cmd.createparameter("yearmon", adChar, adparaminput, 6, cyear+cmonth)
			cmd.Execute ,, adExecuteNoRecords
			clearParameter(cmd)
		Case "D"
			cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput)
			sql = "select count(mdidx) from wb_contact_exe where mdidx = ? and isHold is not null"
			cmd.parameters("mdidx").value = mdidx
			cmd.commandText = sql
			Set rs = cmd.execute
			If rs(0) > 0 Then
				response.write "<script> alert('정산중 또는 정산완료 내역이 있는 매체는 삭제할 수 없습니다.'); </script>"
			else

				' 등록 관리 사진 삭제
				sql = "select desc_01, desc_02, desc_03, desc_04 from wb_contact_photo where mdidx=?"
				cmd.commandText = sql
				cmd.parameters("mdidx").value = mdidx
				Set rs = cmd.execute

				Dim fso : Set fso = server.CreateObject("scripting.filesystemobject")
				Do Until rs.eof
					For intLoop = 0 To 3
						attachFile = "C:\pds\media\"&rs(intLoop)
						if fso.fileexists(attachFile) then
							fso.deletefile(attachFile)
						end If
					Next
					rs.movenext
				Loop
				'소재 사진 삭제
				sql = "delete from wb_contact_photo where mdidx=? "
				cmd.commandText = sql
				cmd.parameters("mdidx").value = mdidx
				cmd.Execute ,, adExecuteNoRecords
				' 소재 집행 내역 삭제
				sql="delete from wb_subseq_exe where mdidx=? "
				cmd.commandText = sql
				cmd.parameters("mdidx").value = mdidx
				cmd.Execute ,, adExecuteNoRecords
				' 집행 내역 삭제
				sql = "delete from wb_contact_exe where mdidx=? "
				cmd.commandText = sql
				cmd.parameters("mdidx").value = mdidx
				cmd.Execute ,, adExecuteNoRecords
				' 매체 면 정보 삭제
				sql = "delete from wb_contact_md_dtl where mdidx=? "
				cmd.commandText = sql
				cmd.parameters("mdidx").value = mdidx
				cmd.Execute ,, adExecuteNoRecords
				' 매체 정보 삭제
				sql = "delete from wb_contact_md where mdidx=? "
				cmd.commandText = sql
				cmd.parameters("mdidx").value = mdidx
				cmd.Execute ,, adExecuteNoRecords
				' 매체 기초 정보 삭제
				sql = "delete from wb_contact_account where mdidx=? "
				cmd.commandText = sql
				cmd.parameters("mdidx").value = mdidx
				cmd.Execute ,, adExecuteNoRecords
				clearParameter(cmd)

			End If

	End select
	clearParameter(cmd)

	cmd.parameters.append cmd.createparameter("contidx", adinteger, adparaminput)
	cmd.parameters.append cmd.createparameter("contidx2", adinteger, adparaminput)
	sql= "update wb_contact_mst set totalprice = (select sum(monthly) from wb_contact_exe a inner join wb_contact_md b on a.mdidx=b.mdidx where b.contidx=?) where contidx=?"

	cmd.commandText = sql
	cmd.parameters("contidx").value = contidx
	cmd.parameters("contidx2").value = contidx
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