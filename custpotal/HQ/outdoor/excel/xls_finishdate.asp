<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<!--<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">-->
<%
	Dim pcustcode : pcustcode = request("custcode")
	Dim pteamcode : pteamcode = request("teamcode")
	Dim cyear : cyear = request("cyear")
	Dim cmonth : cmonth = request("cmonth")
	Dim cyear2 : cyear2 = request("cyear2")
	Dim cmonth2 : cmonth2 = request("cmonth2")

	'response.write cyear & cmonth & "===="
	'response.write cyear2 & cmonth2 & "===="
	'response.write pcustcode & " === "  & pteamcode
	'response.End

	If cyear = "" Then cyear = Year(date)
	If cmonth = "" Then cmonth = Month(date)
	If cyear2 = "" Then cyear2 = Year(date)
	If cmonth2 = "" Then cmonth2 = Month(date)
	If Len(cmonth) = 1 Then cmonth = "0"&cmonth
	If Len(cmonth2) = 1 Then cmonth2 = "0"&cmonth2

	dim sdate : sdate = dateserial(cyear, cmonth, "01")
	dim edate : edate = dateadd("d", -1, dateadd("m", 1,  DateSerial(cyear2, cmonth2, "01")))


	Dim sql : sql = "select c.contidx, c.title, c.firstdate, c.startdate, c.enddate, isnull(c.totalprice,0) as totalprice, isnull(sum(m.monthly),0) as monthly,"
	sql = sql  & " isnull(sum(m.expense),0) as expense, c.custcode , c.flag "
	sql = sql  & " from wb_contact_mst c "
	sql = sql  & " left outer join sc_cust_dtl d on c.custcode = d.custcode "
	sql = sql  & " left outer join VW_CONTACT_EXE_MONTHLY2 m on m.contidx = c.contidx  "
	sql = sql  & " where c.enddate <= '"&edate&"' and c.enddate >= '"&sdate&"' and d.highcustcode like '"&pcustcode&"%' and c.custcode like '"&pteamcode&"%' "
	sql = sql & " group by c.contidx, c.title, c.firstdate, c.startdate, c.enddate, isnull(c.totalprice,0), c.custcode ,c.flag "
	sql = sql  & " order by c.enddate,  c.title,  contidx desc "


	Dim rs : Set rs = server.CreateObject("adodb.recordset")
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
	Response.AddHeader  "Content-Disposition" , "attachment; filename="&cyear&"��"&cmonth&"�� �����Ϻ� ������Ȳ.xls"
%>
<h2> <u>�����Ϻ� ������Ȳ ('<%=cyear%>.<%=CInt(cmonth)%>)</u> </h2>
	  <table border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" width='2000'>
	  <thead bgcolor='#cccccc'>
		  <tr height='20'>
			<th rowspan="2">No</th>
			<th rowspan="2">��ü��</th>
			<th rowspan="2">����<br />
			  �������</th>
			<th colspan="2">���Ⱓ</th>
			<th rowspan="2">�ѱ����(��)</th>
			<th rowspan="2">�������(��)</th>
			<th rowspan="2">�����޾�</th>
			<th rowspan="2">������</th>
			<th rowspan="2">������</th>
			<th rowspan="2">������</th>
			<th rowspan="2">���</th>
		  </tr>
		  <tr height='22'>
			<th>������</th>
			<th>������</th>
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
						<td  class="hd none" style='padding-right:3px; text-align:right;'><%=formatnumber(expense,0)%></td>
						<td  class="hd none" style='padding-right:3px; text-align:right;'><%=formatnumber(income,0)%></td>
						<td  class="hd none" style='padding-right:10px; text-align:right;'><%=formatnumber(incomerate,2)%></td>
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