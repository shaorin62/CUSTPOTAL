<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<!--<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">-->
<%
	Dim pcustcode : pcustcode = request("cmbcustcode")
	Dim pteamcode : pteamcode = request("cmbteamcode")
	Dim cmbseqno : cmbseqno = request("cmbseqno")
	Dim cmbsubno : cmbsubno = request("cmbsubno")
	Dim cmbthmno : cmbthmno = request("cmbthmno")
	If pcustcode = "" Then pcustcode = "%"
	If pteamcode = "" Then pteamcode = "%"
	Dim cyear : cyear = request("cyear")
	Dim cmonth : cmonth = request("cmonth")
	If cyear = "" Then cyear = Year(date)
	If cmonth = "" Then cmonth = Month(date)
	If Len(cmonth) = 1 Then cmonth = "0"&cmonth
	dim sdate : sdate = dateserial(cyear, cmonth, "01")
	dim edate : edate = dateadd("d", -1, dateadd("m", 1,  sdate))

	Dim sql

	if cmbseqno = "" then
		sql = "select a.contidx, a.custcode, a.title, a.startdate, a.enddate, a.flag from wb_contact_mst a inner join sc_cust_dtl b on a.custcode=b.custcode inner join wb_contact_md c on a.contidx=c.contidx inner join vw_contact_md_dtl d on c.mdidx=d.mdidx left outer join vw_subseq_exe e on e.mdidx=d.mdidx and d.side=e.side and e.cyear = '" & cyear &"' and e.cmonth = '"&cmonth &"' left outer join tmp_subseq_mtx f on e.thmno=f.thmno and seqno like '"&cmbseqno&"%' and subno like '"&cmbsubno&"%' and e.thmno like '"&cmbthmno&"%' where a.startdate <= '"&edate&"' and a.enddate >= '"&sdate&"' and a.custcode like '"&pteamcode&"%' and b.highcustcode like '"&pcustcode&"%' group by a.contidx, a.custcode, a.title, a.startdate, a.enddate, a.flag order by a.contidx desc "
	else
		sql = "select a.contidx, a.custcode, a.title, a.startdate, a.enddate, a.flag from wb_contact_mst a inner join sc_cust_dtl b on a.custcode=b.custcode inner join wb_contact_md c on a.contidx=c.contidx inner join vw_contact_md_dtl d on c.mdidx=d.mdidx inner join vw_subseq_exe e on e.mdidx=d.mdidx and d.side=e.side and e.cyear = '" & cyear &"' and e.cmonth = '"&cmonth &"'  inner join tmp_subseq_mtx f on e.thmno=f.thmno and seqno like '"&cmbseqno&"%' and subno like '"&cmbsubno&"%' and e.thmno like '"&cmbthmno&"%' where a.startdate <= '"&edate&"' and a.enddate >= '"&sdate&"' and a.custcode like '"&pteamcode&"%' and b.highcustcode like '"&pcustcode&"%' group by a.contidx, a.custcode, a.title, a.startdate, a.enddate, a.flag order by a.contidx desc  "
	end if

	Dim rs : Set rs = server.CreateObject("adodb.recordset")
	rs.activeconnection = application("connectionstring")
	rs.cursorLocation = aduseclient
	rs.cursortype = adopenstatic
	rs.locktype = adlockoptimistic
	rs.source = sql
	rs.open

	Dim totalrecord : totalrecord = rs.recordcount

	Dim contidx : Set contidx = rs(0)
	Dim custcode : Set custcode = rs(1)
	Dim title : Set title = rs(2)
	Dim startdate : Set startdate = rs(3)
	Dim enddate : Set enddate = rs(4)
	Dim flag : Set flag = rs(5)

'	response.write sql


	sql = "select a.contidx,  b.side, c.monthly, d.thmno ,  a.mdidx from wb_contact_md a inner join vw_contact_md_dtl b on a.mdidx=b.mdidx inner join wb_contact_exe c on b.mdidx=c.mdidx and b.side=c.side and c.cyear='"&cyear&"' and c.cmonth='"&cmonth&"' left outer join  vw_subseq_exe d on c.mdidx=d.mdidx and c.side=d.side and d.cyear='"&cyear&"' and d.cmonth='"&cmonth&"' left outer join tmp_subseq_mtx e on d.thmno=e.thmno where seqno like '"&cmbseqno&"%' and subno like '"&cmbsubno&"%' and d.thmno like '"&cmbthmno&"%' order by a.contidx desc, case when  b.side <> 'L' then ' ' +b.side else b.side end  desc"

'	response.write sql

	Dim rs2 : Set rs2 = server.CreateObject("adodb.recordset")
	rs2.activeconnection = application("connectionstring")
	rs2.cursorlocation = aduseclient
	rs2.cursortype = adopenstatic
	rs2.locktype = adLockOptimistic
	rs2.source = sql
	rs2.open

	If Not rs2.eof Then
		Dim side : Set side = rs2(1)
		Dim monthly : Set monthly = rs2(2)
		Dim thmno : Set thmno = rs2(3)
	End If

	Response.CacheControl  = "public"
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader  "Content-Disposition" , "attachment; filename="&cyear&"년"&cmonth&"월 브랜드별 집행현황.xls"
%>

<h2> <u>브랜드별 광고현황 ('<%=cyear%>.<%=CInt(cmonth)%>)</u> </h2>
	  <table border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" width='2000'>
	  <thead bgcolor='#cccccc'>
		  <tr height='20'>
			<th rowspan="2">No</th>
			<th rowspan="2">매체명</th>
			<th colspan="2">계약기간</th>
			<th rowspan="2">면</th>
			<th rowspan="2">브랜드</th>
			<th rowspan="2">서브 브랜드</th>
			<th rowspan="2">소재명</th>
			<th rowspan="2">월광고료(원)</th>
			<th rowspan="2">광고주</th>
			<th rowspan="2">운영팀</th>
		  </tr>
		  <tr height='22'>
			<th>시작일</th>
			<th>종료일</th>
		  </tr>
		</thead>
		<tbody id='tbody'>
		<%
				Do Until rs.eof
		%>
			<tr height='32'>
				<td  class="hd none" style='text-align:center;padding-top:9px;padding-left:11px;vertical-align:top;'  width="30"><%=totalrecord%> </td>
				<td  class="hd none" style='text-align:left;padding-top:9px;vertical-align:top;' width="210" ><%=title%></td>
				<td  class="hd none"style='text-align:center;padding-top:9px;vertical-align:top;' width="80"><%=startdate%></td>
				<td  class="hd none" style='text-align:center;padding-top:9px;vertical-align:top;' width="80"><%=enddate%></td>
				<td  class="hd none" colspan='5'><table  width='450' border=1 style="table-layout:fixed;" >
				<%
					rs2.Filter = "contidx="&contidx
					Do Until rs2.eof
				%>
					<tr height='32'>
						<td  width="45" style='text-align:center;'><%=side%></td>
						<td  width="110" style='padding-left:5px;'><%=getbrand(thmno)%></td>
						<td  width="110" style='padding-left:5px;'><%=getsubbrand(thmno)%></td>
						<td  width="110" style='padding-left:5px;'><%=getthmname(thmno)%></td>
						<td  width="80" style='text-align:right;padding-right:10px;'><%=FormatNumber(monthly,0)%></td>
					</tr>
				<%
					rs2.movenext
					Loop
				%>
			</table></td>
			<td  class="hd none" width="90" style='text-align:left;padding-top:9px;padding-left:5px;vertical-align:top;' ><%=getcustname(custcode)%></td>
			<td  class="hd none" width="90" style='text-align:left;padding-top:9px;padding-left:5px;vertical-align:top;'  ><%=getteamname(custcode)%></td>
			</tr>
			<%
						totalrecord = totalrecord-1
						rs.movenext
					Loop
			%>
		</tbody>
        </table>