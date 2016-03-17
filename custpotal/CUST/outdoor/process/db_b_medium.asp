<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->	<% '���Ȱ��ö��̺귯�� %>
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

			' ���� ���� ���� (���۳⵵, ���ۿ�)
			Dim startdate : startdate = rs(0)
			Dim enddate : enddate = rs(1)
			Dim eyear : eyear = CStr(Year(enddate))
			Dim emonth
			If Len(Month(enddate)) =  1 Then emonth = "0"&Month(enddate) Else emonth = CStr(Month(enddate))

			Dim term

			' ���� �Ⱓ(��) ��
			pstartdate = startdate
			term = DateDiff("m", startdate, enddate)+1

			pcyear = CStr(Year(startdate))
			If Len(Month(startdate)) =  1 Then pcmonth = "0"&Month(startdate) Else pcmonth = CStr(Month(startdate))


			' 1. ���� ��ü �� ���� �߰�
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
			' 2. ��ü ����� ���� �ӽ� �߰�
			' �� �����͸� ��������
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
			' 3.���� ���� �߰�
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

			' ���� ���� ���� (���۳⵵, ���ۿ�)
			startdate = rs(0)
			enddate = rs(1)

			' ��ü �� �� ���� ���� ���� �߰�
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
			'��ü �麰 ���� ���� ���(���� ������ ����� ���� ������ ����� �űԷ� ���)
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
			' ��ü �麰 ����� ����
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
			' ��� ���� ���� ����
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
			'���� ���� ����
			sql = "delete from wb_contact_photo where mdidx=? and side=?"
			cmd.commandText = sql
			cmd.parameters("mdidx").value = mdidx
			cmd.parameters("side").value = side
			cmd.Execute ,, adExecuteNoRecords
			' ���� ���� ���� ����
			sql="delete from wb_subseq_exe where mdidx=? and side=?"
			cmd.commandText = sql
			cmd.parameters("mdidx").value = mdidx
			cmd.parameters("side").value = side
			cmd.Execute ,, adExecuteNoRecords
			' ���� ���� ����
			sql = "delete from wb_contact_exe where mdidx=? and side=?"
			cmd.commandText = sql
			cmd.parameters("mdidx").value = mdidx
			cmd.parameters("side").value = side
			cmd.Execute ,, adExecuteNoRecords
			' ��ü �� ���� ����
			sql = "delete from wb_contact_md_dtl where mdidx=? and side=?"
			cmd.commandText = sql
			cmd.parameters("mdidx").value = mdidx
			cmd.parameters("side").value = side
			cmd.Execute ,, adExecuteNoRecords
			' ��ü ���� ���� ����
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