
<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
	body {
		background-color:transparent;
		font-size:12px;
	}
	.trhd {
		font-size:12px;
		height: 30px;
		color: #F9F1EA;
		background-color:#9A9A9A;
		font-weight: bolder;
	}

	.trbd {
		font-size:12px;
		height: 30px;
		color: #000000;
	}
</style>
<%
	dim yearmon, yearmon2

	yearmon = cstr(request("yearmon"))
	yearmon2 = cstr(request("yearmon2"))

	Dim custcode : custcode = request("custcode")
	Dim custcode2 : custcode2 = request("custcode2")
	dim total_monthprice, total_expense, total_income, total_incomeratio, prev

	if custcode2 = "" or custcode = custcode2 then custcode2 = Null
	
	if request.cookies("class") = "D" or request.cookies("class") = "H"  then
		custcode2 = request.cookies("custcode2")
	end if

	Response.CacheControl  = "public"
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader  "Content-Disposition" , "attachment; filename=옥외광고 현황.xls"

	dim objrs, sql

	sql = "select custname from dbo.sc_cust_temp where custcode = '" & custcode &"' "
	Call get_recordset(objrs, sql)

	Dim clientname : clientname = objrs(0).value
	objrs.close

	' 이하로 붙여 넣기

	if isnull(custcode2) then

	sql = "select c.custcode, c.custname as custname2, m.contidx, m.title, m.firstdate, m.startdate, m.enddate, m.canceldate, isnull(t.monthprice,0) as totalprice, isnull(sum(d.monthprice),0) as monthprice, isnull(sum(d.expense),0) as expense from dbo.wb_contact_mst m inner join dbo.vw_contact_totalprice t on m.contidx = t.contidx left outer join dbo.wb_contact_md_dtl d on m.contidx = d.contidx inner join dbo.sc_cust_temp c on m.custcode = c.custcode inner join dbo.sc_cust_temp c2 on c.highcustcode = c2.custcode where d.cyear = '"&left(yearmon,4)&"' and d.cmonth = '"&right(yearmon,2)&"' and c2.custcode = '"&custcode&"' group by c.custcode,m.contidx, m.title, m.firstdate, m.startdate, m.enddate, m.canceldate, t.monthprice, c.custname with rollup having (c.custcode is not null and c.custname is not null and m.contidx is not null) or (c.custcode is not null and c.custname is null and m.contidx is null and m.title is null and  m.firstdate is null and m.startdate is null and m.enddate is null)"
'	response.write sql
	call get_recordset(objrs, sql)

	dim cnt, contidx, title, firstdate, startdate, enddate, period, monthprice, expense, income, incomeratio, custname2, totalprice,canceldate, prev_custname2 ,grand_total

	cnt = objrs.recordcount

	if not objrs.eof Then
		set contidx = objrs("contidx")
		set title = objrs("title")
		set firstdate = objrs("firstdate")
		set startdate = objrs("startdate")
		set enddate = objrs("enddate")
		set totalprice = objrs("totalprice")
		set monthprice = objrs("monthprice")
		set expense = objrs("expense")
		set canceldate = objrs("canceldate")
		set custname2 = objrs("custname2")
	end if

%>
				  <table width="1015" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" >
                      <tr class="trhd">
                        <td width="220" align="center" >매체명</td>
                        <td width="75" align="center">최초<br>계약일자</td>
                        <td width="80" align="center">시작일</td>
                        <td width="80" align="center">종료일</td>
                        <td width="80" align="center">총광고료</td>
                        <td width="80" align="center">월광고료</td>
                        <td width="80" align="center">월지급액</td>
                        <td width="80" align="center">내수액</td>
                        <td width="60" align="center">내수율</td>
                        <td width="100" align="center">사업부서</td>
                      </tr>
	     <%
			do until objrs.eof
			if day(startdate) = "1" then
				period = datediff("m", startdate, enddate)+1
			else
				period = datediff("m", startdate, enddate)
			end if


		%>
		<% if  isnull(custname2) then %>
                  <tr class="trbd" bgcolor="#FFFFC1">
                    <td align="left"  style="padding-left:5px;"><%=prev_custname2%> 소계</td>
		<% else %>
                  <tr class="trbd" bgcolor="#FFFFFF">
                    <td  align="left"  style="padding-left:5px;"><%=title%> </td>
		<% end if %>
                    <td  align="center"><%=firstdate%></td>
                    <td align="center"><%=startdate%></td>
                    <td align="center"><%=enddate%></td>
                    <td align="right"><%If Not IsNull(monthprice + expense) Then response.write formatnumber(monthprice + expense,0) Else response.write "0"%>&nbsp;</td>
                    <td align="right"><%If Not IsNull(monthprice) or monthprice <> 0 Then response.write formatnumber(monthprice,0) Else response.write "0"%>&nbsp;</td>
                    <td  align="right"><%If Not IsNull(expense) Then response.write formatnumber(expense,0) Else response.write "0"%>&nbsp;</td>
                    <td align="right"><%If expense <> 0  Then response.write formatnumber(monthprice-expense,0) Else response.write "0"%>&nbsp;</td>
                    <td align="right"><%If monthprice <> 0 Then response.write formatnumber((monthprice-expense)/monthprice*100, 2) Else response.write "0.00"%>&nbsp;</td>
                    <td  align="center"><%=custname2%>&nbsp;</td>
                  </tr>
				<%
						if  not isnull(custname2) then
							grand_total = grand_total + monthprice + expense
							total_monthprice = total_monthprice + monthprice
							total_expense = total_expense + expense
						end if
						prev_custname2 = custname2
						objrs.movenext

					loop
					objrs.close
					set objrs = nothing

					total_income = total_monthprice - total_expense
					if total_income = 0 then
						total_incomeratio = "0.00"
					else
						total_incomeratio = total_monthprice - total_expense / total_monthprice * 100
					end if

				%>
                  <tr height="40" class="trbd"  bgcolor="#FFC1C1">
                    <td  align="center"  >총합계 </td>
                    <td  align="center"  > </td>
                    <td  align="center"  > </td>
                    <td  align="center"  > </td>
                    <td  align="right"  > <%If Not IsNull(grand_total) Then response.write formatnumber(grand_total,0) Else response.write "0"%>&nbsp;</td>
                    <td  align="right" ><%If Not IsNull(total_monthprice) Then response.write formatnumber(total_monthprice,0) Else response.write "0"%>&nbsp;</td>
                    <td  align="right" ><%If Not IsNull(total_expense) Then response.write formatnumber(total_expense,0) Else response.write "0"%>&nbsp;</td>
                    <td align="right" ><%If total_monthprice <> 0  Then response.write formatnumber(total_monthprice-total_expense,0) Else response.write "0"%>&nbsp;</td>
                    <td  align="right" ><%If total_monthprice <> 0 Then response.write formatnumber((total_monthprice-total_expense)/total_monthprice*100, 2) Else response.write "0.00"%>&nbsp;</td>
                    <td  align="center">&nbsp;</td>
                  </tr>
              </table>
<% else


	sql = "select m.contidx, m.title, m.firstdate, m.startdate, m.enddate, m.canceldate, isnull(t.monthprice,0) as totalprice, isnull(sum(d.monthprice),0) as monthprice, isnull(sum(d.expense),0) as expense, c.custname as custname2 from dbo.wb_contact_mst m inner join dbo.vw_contact_totalprice t on m.contidx = t.contidx inner join dbo.wb_contact_md_dtl d on m.contidx = d.contidx  inner join dbo.sc_cust_temp c on m.custcode = c.custcode left outer join dbo.sc_cust_temp c2 on c.highcustcode = c2.custcode where m.custcode = '"&custcode2&"' and d.cyear =  '"&left(yearmon,4)&"' and d.cmonth = '"&right(yearmon,2)&"' group by m.contidx, m.title, m.firstdate, m.startdate, m.enddate, m.canceldate, t.monthprice, c.custname order by m.title"
	call get_recordset(objrs, sql)

	cnt = objrs.recordcount

	if not objrs.eof Then
		set contidx = objrs("contidx")
		set title = objrs("title")
		set firstdate = objrs("firstdate")
		set startdate = objrs("startdate")
		set enddate = objrs("enddate")
		set totalprice = objrs("totalprice")
		set monthprice = objrs("monthprice")
		set expense = objrs("expense")
		set canceldate = objrs("canceldate")
		set custname2 = objrs("custname2")
	end if

%>
				  <table width="1015" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" >
                      <tr class="trhd" >
                        <td width="220" align="center" >매체명</td>
                        <td width="75" align="center">최초<br>계약일자</td>
                        <td width="75" align="center">시작일</td>
                        <td width="75" align="center">종료일</td>
                        <td width="80" align="center">총광고료</td>
                        <td width="80" align="center">월광고료</td>
                        <td width="80" align="center">월지급액</td>
                        <td width="80" align="center">내수액</td>
                        <td width="50" align="center">내수율</td>
                      </tr>
	     <%
			do until objrs.eof
			if day(startdate) = "1" then
				period = datediff("m", startdate, enddate)+1
			else
				period = datediff("m", startdate, enddate)
			end if

		%>
                  <tr class="trbd" bgcolor="#FFFFFF">
                    <td width="220" align="left"  style="padding-left:5px;"><%=title%></td>
                    <td width="75" align="center"><%=firstdate%></td>
                    <td width="75" align="center"><%=startdate%></td>
                    <td width="75" align="center"><%=enddate%></td>
                    <td width="80" align="right"><%If Not IsNull(monthprice + expense) Then response.write formatnumber(monthprice + expense,0) Else response.write "0"%>&nbsp;</td>
                    <td width="80" align="right"><%If Not IsNull(monthprice) or monthprice <> 0 Then response.write formatnumber(monthprice,0) Else response.write "0"%>&nbsp;</td>
                    <td width="80" align="right"><%If Not IsNull(expense) Then response.write formatnumber(expense,0) Else response.write "0"%>&nbsp;</td>
                    <td width="80" align="right"><%If expense <> 0  Then response.write formatnumber(monthprice-expense,0) Else response.write "0"%>&nbsp;</td>
                    <td width="50" align="right"><%If monthprice <> 0 Then response.write formatnumber((monthprice-expense)/monthprice*100, 2) Else response.write "0.00"%>&nbsp;</td>
                  </tr>
				<%
						grand_total = grand_total + monthprice + expense
						total_monthprice = total_monthprice + monthprice
						total_expense = total_expense + expense
						objrs.movenext
					loop
					objrs.close
					set objrs = nothing

					total_income = total_monthprice - total_expense
					if total_income = 0 then
						total_incomeratio = "0.00"
					else
						total_incomeratio = total_monthprice - total_expense / total_monthprice * 100
					end if
				%>
                  <tr height="40" class="trbd"  bgcolor="#FFC1C1">
                    <td  align="center"  >총합계 </td>
                    <td  align="center"  > </td>
                    <td  align="center"  > </td>
                    <td  align="center"  > </td>
                    <td  align="right"  > <%If Not IsNull(grand_total) Then response.write formatnumber(grand_total,0) Else response.write "0"%>&nbsp;</td>
                    <td   align="right" ><%If Not IsNull(total_monthprice) Then response.write formatnumber(total_monthprice,0) Else response.write "0"%>&nbsp;</td>
                    <td  align="right" ><%If Not IsNull(total_expense) Then response.write formatnumber(total_expense,0) Else response.write "0"%>&nbsp;</td>
                    <td   align="right" ><%If total_monthprice <> 0  Then response.write formatnumber(total_monthprice-total_expense,0) Else response.write "0"%>&nbsp;</td>
                    <td  align="right" ><%If total_monthprice <> 0 Then response.write formatnumber((total_monthprice-total_expense)/total_monthprice*100, 2) Else response.write "0.00"%>&nbsp;</td>
                  </tr>
              </table>
<% end if%>