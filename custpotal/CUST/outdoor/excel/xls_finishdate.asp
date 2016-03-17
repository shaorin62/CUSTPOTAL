<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<%
	Dim userid : userid = request.cookies("userid")
	Dim pcustcode : pcustcode = request("custcode")
	Dim pteamcode : pteamcode = request("teamcode")
	Dim cyear : cyear = request("cyear")
	Dim cmonth : cmonth = request("cmonth")
	Dim cyear2 : cyear2 = request("cyear2")
	Dim cmonth2 : cmonth2 = request("cmonth2")

	If cyear = "" Then cyear = Year(date)
	If cmonth = "" Then cmonth = Month(date)
	If cyear2 = "" Then cyear2 = Year(date)
	If cmonth2 = "" Then cmonth2 = Month(date)
	If Len(cmonth) = 1 Then cmonth = "0"&cmonth
	If Len(cmonth2) = 1 Then cmonth2 = "0"&cmonth2
	dim sdate : sdate = dateserial(cyear, cmonth, "01")
	dim edate : edate = dateadd("d", -1, dateadd("m", 1,  DateSerial(cyear2, cmonth2, "01")))




	Dim sql
	Dim chkcount
	chkcount= 1
	Dim Custcodesql
	Dim Custcoderecord
	Dim Timcodesql
	Dim Timcoderecord
	Dim objrs_1
	Dim objrs
	Dim rs
	dim Custcoderecord_cnt
	Dim objrs_cnt

	if pcustcode = "" or pcustcode = null then
		'=========================================================================================
		Custcodesql = "select clientcode from wb_account_cust where userid ='" & userid & "' "

		Set objrs_cnt = server.CreateObject("adodb.recordset")
		objrs_cnt.activeconnection = application("connectionstring")
		objrs_cnt.cursorLocation = aduseclient
		objrs_cnt.cursortype = adopenstatic
		objrs_cnt.locktype = adlockoptimistic
		objrs_cnt.source = Custcodesql
		objrs_cnt.open

		Custcoderecord_cnt = objrs_cnt.recordcount
		'=========================================================================================

		if Custcoderecord_cnt = 0 then
			Custcodesql = "select distinct h.highcustcode as clientcode from wb_contact_mst m inner join sc_cust_dtl d on m.custcode = d.custcode inner join sc_cust_hdr h on d.highcustcode = h.highcustcode "

			Set objrs_1 = server.CreateObject("adodb.recordset")
			objrs_1.activeconnection = application("connectionstring")
			objrs_1.cursorLocation = aduseclient
			objrs_1.cursortype = adopenstatic
			objrs_1.locktype = adlockoptimistic
			objrs_1.source = Custcodesql
			objrs_1.open

			Custcoderecord = objrs_1.recordcount
		else
			Custcodesql = "select clientcode from wb_account_cust where userid ='" & userid & "' "

			Set objrs_1 = server.CreateObject("adodb.recordset")
			objrs_1.activeconnection = application("connectionstring")
			objrs_1.cursorLocation = aduseclient
			objrs_1.cursortype = adopenstatic
			objrs_1.locktype = adlockoptimistic
			objrs_1.source = Custcodesql
			objrs_1.open

			Custcoderecord = objrs_1.recordcount
		end if
		'=========================================================================================



		if not objrs_1.eof then
			do until objrs_1.eof
				'=========================================================================================
				Timcodesql = "select timcode from wb_account_tim where userid ='" & userid & "' and clientcode = '" & objrs_1("clientcode") &"'"

				Set objrs = server.CreateObject("adodb.recordset")
				objrs.activeconnection = application("connectionstring")
				objrs.cursorLocation = aduseclient
				objrs.cursortype = adopenstatic
				objrs.locktype = adlockoptimistic
				objrs.source = Timcodesql
				objrs.open

				Timcoderecord = objrs.recordcount
				'=========================================================================================

				if chkcount > 1 then
					sql = sql  & " Union all "
				end if
				sql = sql  & " select c.contidx, c.title, c.firstdate, c.startdate, c.enddate, isnull(c.totalprice,0) as totalprice, isnull(sum(m.monthly),0) as monthly,"
				sql = sql  & " isnull(sum(m.expense),0) as expense, c.custcode , c.flag "
				sql = sql  & " from wb_contact_mst c "
				sql = sql  & " left outer join sc_cust_dtl d on c.custcode = d.custcode "
				sql = sql  & " inner  join wb_account_cust n on d.highcustcode  = n.clientcode and n.userid='"&userid&"'  and n.clientcode =  '" & objrs_1("clientcode") &"' "
				If Timcoderecord > 0 then
					sql = sql  & " inner  join wb_account_tim t on c.custcode = t.timcode and t.userid='"&userid&"' and t.clientcode = '" & objrs_1("clientcode") &"' "
				End If
				sql = sql  & " left outer join vw_contact_exe_monthly2 m on m.contidx = c.contidx "
				sql = sql  & " where c.enddate <= '"&edate&"' and c.enddate >= '"&sdate&"' "
				sql = sql & " group by c.contidx, c.title, c.firstdate, c.startdate, c.enddate, isnull(c.totalprice,0), c.custcode ,c.flag "


				chkcount = chkcount +1
				objrs_1.movenext
			Loop
			'sql = sql  & " order by c.enddate,  c.title,  contidx desc "
		end if

else

	'=========================================================================================
	Timcodesql = "select timcode from wb_account_tim where userid ='" & userid & "' and clientcode ='" & pcustcode & "'"

	Set objrs = server.CreateObject("adodb.recordset")
	objrs.activeconnection = application("connectionstring")
	objrs.cursorLocation = aduseclient
	objrs.cursortype = adopenstatic
	objrs.locktype = adlockoptimistic
	objrs.source = Timcodesql
	objrs.open

	Timcoderecord = objrs.recordcount
	'=========================================================================================

	sql = sql  & " select c.contidx, c.title, c.firstdate, c.startdate, c.enddate, isnull(c.totalprice,0) as totalprice, isnull(sum(m.monthly),0) as monthly,"
	sql = sql  & " isnull(sum(m.expense),0) as expense, c.custcode , c.flag "
	sql = sql  & " from wb_contact_mst c "
	sql = sql  & " left outer join sc_cust_dtl d on c.custcode = d.custcode "
	sql = sql  & " inner  join wb_account_cust n on d.highcustcode  = n.clientcode and n.userid='"&userid&"'   "
	If Timcoderecord > 0 then
		sql = sql  & " inner  join wb_account_tim t on c.custcode = t.timcode and t.userid='"&userid&"' "
	End If
	sql = sql  & " left outer join vw_contact_exe_monthly2 m on m.contidx = c.contidx "
	sql = sql  & " where c.enddate <= '"&edate&"' and c.enddate >= '"&sdate&"' and d.highcustcode like '"&pcustcode&"%' and c.custcode like '"&pteamcode&"%' "
	sql = sql & " group by c.contidx, c.title, c.firstdate, c.startdate, c.enddate, isnull(c.totalprice,0), c.custcode ,c.flag "
	'sql = sql  & " order by c.enddate,  c.title,  contidx desc "

end if




	Set rs = server.CreateObject("adodb.recordset")
	rs.activeconnection = application("connectionstring")
	rs.cursorLocation = adUseClient
	rs.cursorType = adOpenStatic
	rs.lockType = adLockOptimistic
	rs.source = sql
	rs.open

	Dim totalrecord : totalrecord = rs.recordcount

	Dim contidx : Set contidx = rs(0)
	Dim title : Set title = rs(1)
	Dim firstdate : Set firstdate = rs(2)
	Dim startdate : Set startdate = rs(3)
	Dim enddate : Set enddate = rs(4)
	Dim totalprice : Set totalprice = rs(5)
	Dim monthly : Set monthly = rs(6)
	Dim expense : Set expense = rs(7)
	Dim teamcode : Set teamcode = rs(8)
	Dim flag : Set flag = rs(9)
	Dim income : income = 0
	Dim incomerate : incomerate = "0.00"

	Dim grandtotalprice : grandtotalprice =  0
	Dim grandmonthly : grandmonthly = 0
	Dim grandexpense : grandexpense = 0
	Dim grandincome : grandincome = 0
	Dim grandincomerate : grandincomerate = 0
	Dim grandprice : grandprice = 0

	Response.CacheControl  = "public"
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader  "Content-Disposition" , "attachment; filename="&cyear&"년"&cmonth&"월 종료일별 집행현황.xls"
%>
<h2> <u>종료일별 광고현황 ('<%=cyear%>.<%=CInt(cmonth)%>)</u> </h2>
	  <table border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" width='2000'>
	  <thead bgcolor='#cccccc'>
		  <tr height='20'>
			<th rowspan="2">No</th>
			<th rowspan="2">매체명</th>
			<th rowspan="2">최초<br />
			  계약일자</th>
			<th colspan="2">계약기간</th>
			<th rowspan="2">총광고료(원)</th>
			<th rowspan="2">월광고료(원)</th>
			<!-- <th rowspan="2">월지급액</th>
			<th rowspan="2">내수액</th>
			<th rowspan="2">내수율</th> -->
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
							income = monthly-expense
							If monthly = 0 Then incomerate = "0.00" Else 	incomerate = income/monthly*100
					%>
					<tr height='32'>
						<td  class="hd none" style='padding-left:10px; text-align:left;'><span name="totalrecord"><%=totalrecord%></span></td>
						<td  class="hd none" style="padding-left:5px;"><%=title%></a></td>
						<td  class="hd none" style=' text-align:center;'><%=firstdate%></td>
						<td  class="hd none" style=' text-align:center;'><%=startdate%></td>
						<td  class="hd none" style=' text-align:center;'><%=enddate%></td>
						<td  class="hd none" style='padding-right:3px; text-align:right;'><%=FormatNumber(totalprice, 0)%></td>
						<td  class="hd none" style='padding-right:3px; text-align:right;'><%=FormatNumber(monthly, 0)%></td>
						<!-- <td  class="hd none" style='padding-right:3px; text-align:right;'><%=formatnumber(expense,0)%></td>
						<td  class="hd none" style='padding-right:3px; text-align:right;'><%=formatnumber(income,0)%></td>
						<td  class="hd none" style='padding-right:10px; text-align:right;'><%=formatnumber(incomerate,2)%></td> -->
						<td  class="hd none" style='padding-left:3px;'><%=getcustname(teamcode)%></td>
						<td  class="hd none" style='padding-left:3px;'><%=getteamname(teamcode)%></td>
					</tr>
				  <%
							totalrecord = totalrecord - 1
							rs.movenext
						Loop
				  %>
				  </tbody>
              </table>