<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<%
'	For Each item In request.querystring
'		response.write item & " :" & request.querystring(item) & "<br>"
'	Next
	
	Dim pcustcode : pcustcode = request("custcode")
	Dim pteamcode : pteamcode = request("teamcode")
	Dim cyear : cyear = request("cyear")
	Dim cmonth : cmonth = request("cmonth")
	If cyear = "" Then cyear = Year(date)
	If cmonth = "" Then cmonth = Month(date)
	If Len(cmonth) = 1 Then cmonth = "0"&cmonth
	dim sdate : sdate = dateserial(cyear, cmonth, "01")
	dim edate : edate = dateadd("d", -1, dateadd("m", 1,  sdate))
	
	Dim sql : sql = "select a.contidx, a.title, a.firstdate, a.startdate, a.enddate, isnull(totalprice, 0) totalprice , c.monthly, c.expense, d.highcustcode, a.custcode, a.comment from wb_contact_mst a  inner join (select a.contidx, sum(monthly) monthly, sum(expense) expense from wb_contact_md a inner join wb_contact_exe  b on a.mdidx=b.mdidx where cyear = '"&cyear&"' and cmonth='"&cmonth&"' group by contidx) as c on a.contidx=c.contidx inner join sc_cust_dtl d on a.custcode=d.custcode where a.custcode like '" & pteamcode &"%' and d.highcustcode like '" & pcustcode &"%' and a.startdate <= '"&edate&"' and a.enddate >= '"&sdate&"'  order by flag, contidx"

	response.write sql

	Dim rs : Set rs = server.CreateObject("adodb.recordset")
	rs.activeconnection = application("connectionstring")
	rs.cursorLocation = aduseclient
	rs.cursortype = adopenstatic
	rs.locktype = adlockoptimistic
	rs.source = sql
	rs.open 
	
	If Not rs.eof Then 
		Dim contidx : Set contidx = rs(0)
		Dim title : Set title = rs(1)
		Dim firstdate : Set firstdate = rs(2)
		Dim startdate : Set startdate = rs(3)
		Dim enddate : Set enddate = rs(4)
		Dim totalprice : Set totalprice = rs(5)
		Dim monthly : Set monthly = rs(6)
		Dim expense : Set expense = rs(7)
		Dim highcustcode : Set highcustcode = rs(8)
		Dim custcode : Set custcode = rs(9)
		Dim comment : Set comment = rs(10)
	
	End If 

	'response.write sql


	sql = "select a.contidx, mdidx, locate from wb_contact_mst a inner join wb_contact_md b on a.contidx=b.contidx inner join sc_cust_dtl c on a.custcode=c.custcode where a.startdate <= '"&edate&"' and a.enddate >= '"&sdate&"' and a.custcode like '" & pteamcode &"%' and c.highcustcode like '" & pcustcode &"%'  order by flag, contidx "

'	response.write sql

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

	sql = "select c.mdidx, c.side, standard, quality, qty, unit, flag from wb_contact_mst a  inner join wb_contact_md b on a.contidx=b.contidx inner join wb_contact_md_dtl c on b.mdidx=c.mdidx inner join (select mdidx, side, sum(qty) qty from wb_contact_exe where cyear='"&cyear&"' and cmonth='"&cmonth&"' group by mdidx, side) as d on d.mdidx=c.mdidx and d.side=c.side inner join sc_cust_dtl e on a.custcode=e.custcode where e.custcode like '"&pteamcode&"%' and e.highcustcode like '"&pcustcode&"%' and a.startdate <= '"&edate&"' and a.enddate >= '"&sdate&"' order by flag, a.contidx, c.mdidx, side desc"

'	response.write sql

	Dim rs3 : Set rs3 = server.CreateObject("adodb.recordset")
	rs3.activeconnection = application("connectionstring")
	rs3.cursorlocation = aduseclient
	rs3.cursortype = adopenstatic
	rs3.locktype = adLockOptimistic
	rs3.source = sql
	rs3.open

	If Not rs3.eof Then 
		Dim side : Set side = rs3(1)
		Dim standard : Set standard = rs3(2)
		Dim quality : Set quality = rs3(3)
		Dim qty : Set qty = rs3(4)
		Dim unit : Set unit = rs3(5)
		Dim flag : Set flag = rs3(6)
	End If 

	
	sql = "select c.mdidx, c.side, thmno from wb_contact_mst a  inner join wb_contact_md b on a.contidx=b.contidx inner join wb_contact_md_dtl c on b.mdidx=c.mdidx inner join (select mdidx, side, thmno from wb_subseq_exe where seq in (select max(seq) from wb_subseq_exe where cyear+cmonth <= '"&cyear+cmonth&"' and no=1 group by mdidx, side)) as d on d.mdidx=c.mdidx and d.side=c.side inner join sc_cust_dtl e on a.custcode=e.custcode where e.custcode like '"&pteamcode&"%' and e.highcustcode like '"&pcustcode&"%' and a.startdate <= '"&edate&"' and a.enddate >= '"&sdate&"' order by flag, a.contidx, c.mdidx, side desc"

'	response.write sql

	Dim rs4 : Set rs4 = server.CreateObject("adodb.recordset")
	rs4.activeconnection = application("connectionstring")
	rs4.cursorlocation = aduseclient
	rs4.cursortype = adopenstatic
	rs4.locktype = adLockOptimistic
	rs4.source = sql
	rs4.open

	If Not rs4.eof Then 
		Dim thmno : Set thmno = rs4(2)
	End If 

	sql = "select distinct a.contidx, medcode, flag from wb_contact_mst a inner join wb_contact_md b on a.contidx=b.contidx inner join sc_cust_dtl c on a.custcode=c.custcode where a.startdate <= '"&edate&"' and a.enddate >= '"&sdate&"' and a.custcode like '" & pteamcode &"%' and c.highcustcode like '" & pcustcode &"%'  order by flag, contidx "

'	response.write sql

	Dim rs5 : Set rs5 = server.CreateObject("adodb.recordset")
	rs5.activeconnection = application("connectionstring")
	rs5.cursorlocation = aduseclient
	rs5.cursortype = adopenstatic
	rs5.locktype = adLockOptimistic
	rs5.source = sql
	rs5.open
	If Not rs5.eof Then 
		Dim medcode : Set medcode = rs5(1)
	End If 
	

'	Response.CacheControl  = "public"
'	Response.ContentType = "application/vnd.ms-excel"
'	Response.AddHeader  "Content-Disposition" , "attachment; filename="&cyear&"년"&cmonth&"월 옥외광고 집행현황.xls"
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
				income = monthly - expense
				If monthly = 0 Then rate = 0 Else rate = income/monthly

		%>
			<tr>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;'><%=no%></td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;' > <%=title%></td>
				<td style='text-align:left;vertical-align:top;background-color:#FFFFFF;' colspan='4'>
					<!-- 매체 정보 -->
					<table border ="1" cellpadding="0" cellspacing="0">
					<% 
						rs2.Filter = "contidx="& contidx 
							Do Until rs2.eof  
					%>
						<tr>
							<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;width:498px;' ><%=locate%></td>
							<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;' colspan='3'>
								<table border ="1" cellpadding="0" cellspacing="0">
								<%
									rs3.Filter = "mdidx="&mdidx
									Do Until rs3.eof 
								%>
								<tr>
									<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;width:78px;border-top:0px; border-bottom:0px;' > <%=qty%><%=unit%></td>
									<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;width:48px;border-top:0px; border-bottom:0px;' > <%If flag = "B" Then response.write side%></td>
									<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;width:300px;border-top:0px; border-bottom:0px;' ><%=standard%> (<%=quality%>)</td>
								</tr>
								<%
									rs3.movenext
									Loop
								%>
								</table>
							</td>
						</tr>
					<%
							rs2.movenext
							Loop
							rs2.movefirst
					%>
					</table>
					<!--  -->
				</td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;' > <%=firstdate%></td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;' > <%=startdate%></td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;' > <%=enddate%></td>
				<td style='text-align:right;vertical-align:top;background-color:#FFFFFF;' > <%=FormatNumber(totalprice,0)%></td>
				<td style='text-align:right;vertical-align:top;background-color:#FFFFFF;' > <%=FormatNumber(monthly,0)%></td>
				<td style='text-align:right;vertical-align:top;background-color:#FFFFFF;' > <%=FormatNumber(expense,0)%></td>
				<td style='text-align:right;vertical-align:top;background-color:#FFFFFF;' > <%=FormatNumber(income,0)%></td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;' > <%=FormatPercent(rate,0)%></td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;' >
				<%
					Do Until rs2.eof 
					rs4.Filter = "mdidx="&mdidx
						Do Until rs4.eof 
							response.write getthmname(thmno) & "<br>"
						rs4.movenext
						Loop
					rs2.movenext
					Loop
					rs2.movefirst
				%>
				</td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;' > <%=getteamname(custcode)%></td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;' >
					<!--  -->
					<%
						rs5.Filter = "contidx="&contidx
						Do Until rs5.eof 
							response.write getmedname(medcode) & "<br>"
							rs5.movenext
						Loop
					%>
					<!--  -->
				</td>
				<td style='text-align:center;vertical-align:top;background-color:#FFFFFF;' > <%'=comment%></td>
			<%
						no = no + 1
						rs.movenext
					Loop
			%>
		</tbody>
        </table>
</body>
</html>