<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<!--<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">-->

<%
	Dim pcustcode : pcustcode = request("custcode")
	Dim pteamcode : pteamcode = request("teamcode")
	Dim cyear : cyear = request("cyear")
	If cyear = "" Then cyear = Year(date)
	Dim sdate : sdate = DateSerial(cyear, "01", "01")
	Dim edate : edate = DateSerial(cyear, "12", "31")

	Dim sql : sql = "select d.highcustcode, a.custcode"
	sql = sql & ",sum(case when c.cmonth = '01' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '02' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '03' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '04' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '05' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '06' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '07' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '08' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '09' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '10' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '11' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '12' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(isnull(c.monthly,0)) "
	sql = sql & "from wb_contact_mst a inner join wb_contact_md b on a.contidx=b.contidx "
	sql = sql & "inner join wb_contact_exe c on c.mdidx=b.mdidx and cyear='"&cyear&"' "
	sql = sql & "inner join sc_cust_dtl d on a.custcode=d.custcode "
	sql = sql & "inner join wb_contact_trans e  "
	sql = sql & " on a.contidx=e.contidx and b.medcode=e.medcode and c.cyear = e.cyear and c.cmonth = e.cmonth and e.ishold in('Y','N') "
	sql = sql & "where d.highcustcode like '" & pcustcode & "%' and a.custcode like '" & pteamcode & "%' "
	sql = sql & "group by d.highcustcode, a.custcode with rollup"

'
	'response.write sql

	Dim rs : Set rs = server.CreateObject("adodb.recordset")
	rs.activeconnection = application("connectionstring")
	rs.cursorLocation = adUseClient
	rs.cursorType = adOpenStatic
	rs.lockType = adLockOptimistic
	rs.source = sql
	rs.open

	If Not rs.eof Then
		Dim custcode : Set custcode = rs(1)
		Dim jan : Set jan = rs(2)
		Dim feb : Set feb = rs(3)
		Dim mar : Set mar = rs(4)
		Dim apr : Set apr = rs(5)
		Dim may : Set may = rs(6)
		Dim jun : Set jun = rs(7)
		Dim jul : Set jul = rs(8)
		Dim aug : Set aug = rs(9)
		Dim sep : Set sep = rs(10)
		Dim oct_ : Set oct_ = rs(11)
		Dim nov : Set nov = rs(12)
		Dim dec : Set dec = rs(13)
		Dim sum : Set sum = rs(14)
	End If


	sql = "select distinct b.highcustcode, c.custname from wb_contact_mst a inner join sc_cust_dtl b on a.custcode=b.custcode inner join sc_cust_hdr c on b.highcustcode=c.highcustcode where a.custcode like '"&pteamcode&"%' and b.highcustcode like '"&pcustcode&"%' and a.startdate <= '"&edate&"' and a.enddate >= '"&sdate&"' "
'	response.write sql
	Dim cmd : set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adcmdText
	Dim rs2 : Set rs2 = cmd.execute
	If Not rs2.eof Then
		Dim highcustcode : Set highcustcode = rs2(0)
		Dim custname : Set custname = rs2(1)
	End If
	Set cmd = Nothing

	Dim s01 : s01 = 0
	Dim s02 : s02 = 0
	Dim s03 : s03 = 0
	Dim s04 : s04 = 0
	Dim s05 : s05 = 0
	Dim s06 : s06 = 0
	Dim s07 : s07 = 0
	Dim s08 : s08 = 0
	Dim s09 : s09 = 0
	Dim s10 : s10 = 0
	Dim s11 : s11 = 0
	Dim s12 : s12 = 0
	Dim total : total = 0


	Response.CacheControl  = "public"
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader  "Content-Disposition" , "attachment; filename="&cyear&"년 광고주별 집행현황.xls"
%>
<h2> <u>광고주별 집행현황 ('<%=cyear%>)</u> </h2>
	  <table border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" width='2000'>
	  <thead bgcolor='#cccccc'>
		  <tr height='20'>
			<th >광고주</th>
			<th >운영팀</th>
			<th >01월</th>
			<th >02월</th>
			<th >03월</th>
			<th >04월</th>
			<th >05월</th>
			<th >06월</th>
			<th >07월</th>
			<th >08월</th>
			<th >09월</th>
			<th >10월</th>
			<th >11월</th>
			<th >12월</th>
			<th >합계</th>
		  </tr>
		</thead>
		<tbody id='tbody'>
					<%
							Do Until rs2.eof
							rs.Filter = "highcustcode='"& highcustcode &"' "
							If rs.recordcount <> 0 Then
					%>
						<tr height='32'>
							<td  width='100' class="hd none" style='padding-left:3px;padding-top:9px;' valign='top'><%=custname%></td>
							<td  class="hd none" colspan='14' ><table  width='930' border=1 style="table-layout:fixed;">
							<%
								Do Until rs.eof
							%>
								<tr height='30' <% If IsNull(custcode) Then response.write "bgcolor=#ececec" End If %>>
									<td  width='100'  style='padding-left:3px;'><%If IsNull(custcode) Then response.write "소계" Else response.write getteamname(custcode) %></td>
									<td  width='62' style='padding-right:5px;text-align:right;'><%=FormatNumber(jan,0)%></td>
									<td  width='62' style='padding-right:5px;text-align:right;'><%=FormatNumber(feb,0)%></td>
									<td  width='62' style='padding-right:5px;text-align:right;'><%=FormatNumber(mar,0)%></td>
									<td  width='62' style='padding-right:5px;text-align:right;'><%=FormatNumber(apr,0)%></td>
									<td  width='62' style='padding-right:5px;text-align:right;'><%=FormatNumber(may,0)%></td>
									<td  width='62' style='padding-right:5px;text-align:right;'><%=FormatNumber(jun,0)%></td>
									<td  width='62' style='padding-right:5px;text-align:right;'><%=FormatNumber(jul,0)%></td>
									<td  width='62' style='padding-right:5px;text-align:right;'><%=FormatNumber(aug,0)%></td>
									<td  width='62' style='padding-right:5px;text-align:right;'><%=FormatNumber(sep,0)%></td>
									<td  width='62' style='padding-right:5px;text-align:right;'><%=FormatNumber(oct_,0)%></td>
									<td  width='62' style='padding-right:5px;text-align:right;'><%=FormatNumber(nov,0)%></td>
									<td  width='62' style='padding-right:5px;text-align:right;'><%=FormatNumber(dec,0)%></td>
									<td  width='80' style='padding-right:10px;text-align:right;'><%=FormatNumber(sum,0)%></td>
								</tr>
							<%
								rs.movenext
								Loop

							%>
						</table></td>
						</tr>
						<%
							rs.movelast
							End If
								rs2.movenext
							Loop
						%>
		</tbody>
              </table>