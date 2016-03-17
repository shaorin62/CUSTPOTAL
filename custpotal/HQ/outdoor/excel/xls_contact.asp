<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<!--<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">-->
<%
	Dim pcustcode : pcustcode = request("custcode")
	Dim pteamcode : pteamcode = request("teamcode")
	Dim cyear : cyear = request("cyear")
	Dim cmonth : cmonth = request("cmonth")
	If cyear = "" Then cyear = Year(date)
	If cmonth = "" Then cmonth = Month(date)
	If Len(cmonth) = 1 Then cmonth = "0"&cmonth
	dim sdate : sdate = dateserial(cyear, cmonth, "01")
	dim edate : edate = dateadd("d", -1, dateadd("m", 1,  sdate))

	Dim sql : sql = "select  contidx, title, firstdate, startdate, enddate, isnull(totalprice,0) totalprice, a.custcode, a.comment from wb_contact_mst  a inner join sc_cust_dtl b on a.custcode=b.custcode where a.custcode like '"&pteamcode&"%' and b.highcustcode like '"&pcustcode&"%' and a.startdate <= '"&edate&"' and a.enddate >= '"&sdate&"' and a.contidx <> 845 order by  a.flag, a.contidx"

	'Dim sql : sql = "select c.contidx, c.title, c.firstdate, c.startdate, c.enddate, isnull(c.totalprice,0) as totalprice, isnull(m.monthly,0) as monthly,"
	'sql = sql  & " isnull(m.expense,0) as expense, c.custcode as teamcode, d.custname as teamname, d2.custname as custname, c.flag "
	'sql = sql  & " from wb_contact_mst c "
	'sql = sql  & " left outer join sc_cust_dtl d on c.custcode = d.custcode "
	'sql = sql  & " left outer join sc_cust_hdr d2 on d.highcustcode = d2.highcustcode "
	'sql = sql  & " left outer join vw_contact_exe_monthly m on m.contidx = c.contidx and m.cyear='"&cyear&"' and m.cmonth='"&cmonth&"' "
	'sql = sql  & " where c.startdate <= '"&edate&"' and c.enddate >= '"&sdate&"' and d.highcustcode like '"&pcustcode&"%' and c.custcode like '"&pteamcode&"%' "
	'sql = sql  & " order by contidx desc "

	'response.write sql
	'response.end

	Dim rs : Set rs = server.CreateObject("adodb.recordset")
	rs.activeconnection = application("connectionstring")
	rs.cursorLocation = adUseClient
	rs.cursorType = adOpenStatic
	rs.lockType = adLockOptimistic
	rs.source = sql
	rs.open

	If Not rs.eof Then
		Dim contidx : Set contidx = rs(0)
		Dim title : Set title = rs(1)
		Dim firstdate : Set firstdate = rs(2)
		Dim startdate : Set startdate = rs(3)
		Dim enddate : Set enddate = rs(4)
		Dim totalprice : Set totalprice = rs(5)
		Dim custcode : Set custcode = rs(6)
		Dim comment : Set comment = rs(7)
	End If

	sql = "select  a.contidx,a.mdidx, isnull(a.locate,'') locate from wb_contact_md a inner join  wb_contact_mst d on a.contidx=d.contidx inner join sc_cust_dtl e on d.custcode=e.custcode where d.custcode like '"&pteamcode&"%' and e.highcustcode like '"&pcustcode&"%' and d.startdate <= '"&edate&"' and d.enddate >= '"&sdate&"' order by d.flag, a.contidx "

'	response.write sql
'	response.end

	Dim rs2 : Set rs2 = server.CreateObject("adodb.recordset")
	rs2.activeconnection = application("connectionstring")
	rs2.cursorlocation = aduseclient
	rs2.cursortype = adopenstatic
	rs2.locktype = adLockOptimistic
	rs2.source = sql
	rs2.open

	If Not rs2.eof Then
		Dim mdidx : Set mdidx = rs2(1)
		Dim locate : Set locate = rs2(2)
	End If
'	rs2.movefirst


	sql = "select  a.contidx,a.mdidx, a.locate, b.side, b.standard, b.quality,  isnull(f.qty,0) qty, a.unit, c.thmno, flag from wb_contact_md a left join vw_contact_md_dtl b on a.mdidx=b.mdidx left outer join (select mdidx, side, thmno  from wb_subseq_exe where seq in (select max(seq) from wb_subseq_exe where cyear+cmonth < '"&cyear+cmonth&"' group by mdidx, side)) as c on b.mdidx=c.mdidx and b.side=c.side  inner join wb_contact_mst d on a.contidx=d.contidx inner join sc_cust_dtl e on d.custcode=e.custcode left outer join wb_contact_exe f on b.mdidx=f.mdidx and b.side=f.side and f.cyear='"&cyear&"' and f.cmonth='"&cmonth&"' where d.custcode like '"&pteamcode&"%' and e.highcustcode like '"&pcustcode&"%' and d.startdate <= '"&edate&"' and d.enddate >= '"&sdate&"' order by d.flag, a.contidx, case when b.side <>'L' then ' ' + b.side else b.side end desc "


	Dim rs3 : Set rs3 = server.CreateObject("adodb.recordset")
	rs3.activeconnection = application("connectionstring")
	rs3.cursorlocation = aduseclient
	rs3.cursortype = adopenstatic
	rs3.locktype = adLockOptimistic
	rs3.source = sql
	rs3.open

	If Not rs2.eof Then
		Dim side : Set side = rs3(3)
		Dim standard : Set standard = rs3(4)
		Dim quality : Set quality = rs3(5)
		Dim qty : Set qty = rs3(6)
		Dim unit : Set unit = rs3(7)
		Dim thmno : Set thmno = rs3(8)
		Dim flag : Set flag = rs3(9)
	End If

	sql = "select  a.contidx, b.medcode, sum(isnull(monthly,0)) monthly, sum(isnull(expense,0)) expense from wb_contact_mst a inner join wb_contact_md b on a.contidx=b.contidx inner join wb_contact_exe c on b.mdidx=c.mdidx and cyear='"&cyear&"' and cmonth='"&cmonth&"' inner join sc_cust_dtl d on a.custcode=d.custcode where d.custcode like '"&pteamcode&"%' and d.highcustcode like '"&pcustcode&"%' and a.startdate <= '"&edate&"' and a.enddate >= '"&sdate&"' group by a.contidx, medcode, flag order by a.flag, a.contidx "

	Dim rs4 : Set rs4 = server.CreateObject("adodb.recordset")
	rs4.activeconnection = application("connectionstring")
	rs4.cursorlocation = aduseclient
	rs4.cursortype = adopenstatic
	rs4.locktype = adLockOptimistic
	rs4.source = sql
	rs4.open

	If Not rs3.eof Then
		Dim medcode : Set medcode = rs4(1)
		Dim monthly : Set monthly = rs4(2)
		Dim expense : Set expense = rs4(3)
	End If

	Response.CacheControl  = "public"
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader  "Content-Disposition" , "attachment; filename="&cyear&"년"&cmonth&"월 옥외광고 집행현황.xls"

%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  >
<h2> <u>옥외광고현황 ('<%=cyear%>.<%=CInt(cmonth)%>)</u> </h2>
	  <table border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" width='2000'>
	  <thead bgcolor='#cccccc'>
		  <tr height='20'>
			<th rowspan="2">No</th>
			<th rowspan="2">매체명</th>
			<th rowspan="2" style="width:500px;">장소</th>
			<th rowspan="2" style="width:80px;">수량(면)</th>
			<th rowspan="2" style="width:50px;">면</th>
			<th rowspan="2" style="width:300px;">규 격(M) / 재 질</th>
			<th rowspan="2">최초<br />
			  계약일자</th>
			<th colspan="2">계약기간</th>
			<th rowspan="2">총광고료(원)</th>
			<th rowspan="2">월광고료(원)</th>
			<th rowspan="2">월외주비</th>
			<th rowspan="2">내수액</th>
			<th rowspan="2">내수율</th>
			<th rowspan="2">광고내용</th>
			<th rowspan="2">운영팀</th>
			<th rowspan="2">매체사</th>
			<th rowspan="2">비고</th>
		  </tr>
		  <tr height='22'>
			<th>시작일</th>
			<th>종료일</th>
		  </tr>
		</thead>
		<tbody id='tbody'>
		<%
				Dim no : no = 1
				Dim income, rate
				Do Until rs.eof
		%>
			<tr>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;'><%=no%></td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;' > <%=title%></td>
				<td style='text-align:left;vertical-align:top;background-color:#FFFFFF;width:930px;' colspan='4'><table border ="1" cellpadding="0" cellspacing="0">
				<!-- 매체정보 -->
				<%
					rs2.filter="contidx=" & contidx
					Do Until rs2.eof
				%>
				</td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;500px' > <%=locate%></td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;' colspan='3' ><table border ="1" cellpadding="0" cellspacing="0">
				<%
					rs3.filter="mdidx="&mdidx
					Do Until rs3.eof
				%>
				</td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;width:80px;' > <%=qty%><%=unit%></td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;width:50px;' > <%If flag = "B" Then response.write side%></td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;width:300px;' > <%=standard%> (<%=getStringQuality(quality)%>)</td>
				</tr>
				<%
					rs3.movenext
					Loop
					rs3.movefirst
				%>
				</table>
				</td>
				</tr>
				<%
					rs2.movenext
					Loop
'					rs2.movefirst
				%>
				</table>
				<!-- 매체정보 -->
				</td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;' > <%=firstdate%></td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;' > <%=startdate%></td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;' > <%=enddate%></td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;' > <%=FormatNumber(totalprice,0)%></td>
				<td style='text-align:right;vertical-align:top;background-color:#FFFFFF;' >
				<!-- 매체사별 월광고료 금액정보 -->
				<%
					rs4.Filter = "contidx="&contidx
					Do Until rs4.eof
						response.write FormatNumber(monthly,0) & "<br>"
					rs4.movenext
					Loop
'					rs4.movefirst
				%>
				<!-- 매체사별 금액정보 -->
				</td>
				<td style='text-align:right;vertical-align:top;background-color:#FFFFFF;' >
				<!-- 매체사별 월지급액 금액정보 -->
				<%
					rs4.Filter = "contidx="&contidx
					Do Until rs4.eof
						response.write FormatNumber(expense,0) & "<br>"
					rs4.movenext
					Loop
'					rs4.movefirst
				%>
				<!-- 매체사별 금액정보 -->
				</td>
				<td style='text-align:right;vertical-align:top;background-color:#FFFFFF;' >
				<!-- 매체사별 내수율 금액정보 -->
				<%
					rs4.Filter = "contidx="&contidx
					Do Until rs4.eof
					income = monthly-expense
						Response.write FormatNumber(income,0) & "<br>"
					rs4.movenext
					Loop
'					rs4.movefirst
				%>
				<!-- 매체사별 금액정보 -->
				</td>
				<td style='text-align:right;vertical-align:top;background-color:#FFFFFF;' >
				<!-- 매체사별 내수율 정보 -->
				<%
					rs4.Filter = "contidx="&contidx
					Do Until rs4.eof
					income = monthly-expense
					If monthly = 0 Then rate="0.00" Else rate = income/monthly
						response.write formatPercent(rate,2) & "<br>"
					rs4.movenext
					Loop
'					rs4.movefirst
				%>
				<!-- 매체사별 금액정보 -->
				</td>
				<td style='text-align:left;vertical-align:top;background-color:#FFFFFF;' >
				<!-- 광고내용 -->
				<%
					rs2.filter="contidx=" & contidx
					Do Until rs2.eof
						response.write getthmname(thmno) & "<br>"
					rs2.movenext
					Loop
'					rs2.movefirst
				%>
				<!-- 광고내용 -->
				</td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;' > <%=getteamname(custcode)%></td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;' >
				<!-- 매체사정보 -->
				<%
					rs4.Filter = "contidx="&contidx
					Do Until rs4.eof
						Response.write getmedname(medcode) & "<br>"
					rs4.movenext
					Loop
'					rs4.movefirst
				%>
				<!-- 매체사정보 -->
				</td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;' ><%=comment%>&nbsp;</td>
			<%
						no = no + 1
						rs.movenext
					Loop
			%>
		</tbody>
        </table>
</body>
</html>