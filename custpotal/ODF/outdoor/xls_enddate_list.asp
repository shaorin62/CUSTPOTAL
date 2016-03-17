<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim gotopage : gotopage = request.querystring("gotopage")
	if gotopage = "" then gotopage = 1
	dim searchstring : searchstring = request("txtsearchstring")
	dim custcode : custcode = request("selcustcode")
	dim custcode2 : custcode2 = request("selcustcode2")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	dim cyear2 : cyear2 = request("cyear2")
	dim cmonth2 : cmonth2 = request("cmonth2")
	if cyear = "" then cyear = year(date)
	if cmonth = "" then cmonth = month(date)
	if cyear2 = "" then cyear2 = year(date)
	if cmonth2 = "" then cmonth2 = month(date)
	dim c_date : c_date = Dateserial(cyear, cmonth, "01")
	dim c_date2 : c_date2 = dateadd("d", -1, dateAdd("m", 1, Dateserial(cyear2, cmonth2, "01")))

	dim objrs, sql
	sql = "select m.contidx, m.title, m.firstdate, m.startdate, m.enddate, m.canceldate, isnull(t.monthprice,0) as totalprice, isnull(sum(d.monthprice),0) as monthprice, isnull(sum(d.expense),0) as expense, c.custname as custname2 from dbo.wb_contact_mst m left outer join dbo.vw_contact_totalprice t on m.contidx = t.contidx left outer join dbo.wb_contact_md_dtl d on m.contidx = d.contidx and d.cyear =  '"&cyear&"' and d.cmonth = '"&cmonth&"' left outer join dbo.sc_cust_temp c on m.custcode = c.custcode left outer join dbo.sc_cust_temp c2 on c.highcustcode = c2.custcode where m.title like '%" & searchstring & "%' and m.custcode like '"&custcode2&"%' and c2.custcode like '"&custcode&"%' and (m.enddate between '" & c_date & "' and '" & c_date2 & "' )  group by m.contidx, m.title, m.firstdate, m.startdate, m.enddate, m.canceldate, t.monthprice, c.custname order by m.title"

	call get_recordset(objrs, sql)

	dim cnt, contidx, title, firstdate, startdate, enddate, period, monthprice, expense, income, incomeratio, custname2, totalprice,canceldate

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


	Response.CacheControl  = "public"
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader  "Content-Disposition" , "attachment; filename="&cyear&"."&cmonth&"종료일별집행현황.xls"
%>
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
	.trbd2 {
		font-size:13px;
		height: 30px;
		color: #000000;
		font-weight: bolder;
		background-color:#9A9A9A;
	}
</style>

<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
				  <table width="1024" border="0" cellspacing="0" cellpadding="0" class="header">
                      <tr>
                        <td width="220" align="center" >매체명</td>
                        <td width="75" align="center">최초<br>계약일자</td>
                        <td width="75" align="center">시작일</td>
                        <td width="75" align="center">종료일</td>
                        <td width="80" align="center">총광고료</td>
                        <td width="80" align="center">월광고료</td>
                        <td width="80" align="center">월지급액</td>
                        <td width="80" align="center">내수액</td>
                        <td width="50" align="center">내수율</td>
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
                  <tr class="trbd">
                    <td width="220" align="left"><%=title%></td>
                    <td width="75" align="center"><%=firstdate%></td>
                    <td width="75" align="center"><%=startdate%></td>
                    <td width="75" align="center"><%=enddate%></td>
                    <td width="80" align="right"><%If Not IsNull(monthprice) Then response.write formatnumber(monthprice,0) Else response.write "0"%></td>
                    <td width="80" align="right"><%If Not IsNull(monthprice) Then response.write formatnumber(monthprice/period,0) Else response.write "0"%></td>
                    <td width="80" align="right"><%If Not IsNull(expense) Then response.write formatnumber(expense/period,0) Else response.write "0"%></td>
                    <td width="80" align="right"><%If expense <> 0  Then response.write formatnumber(monthprice/period-expense/period,0) Else response.write "0"%></td>
                    <td width="50" align="right"><%If monthprice/period <> 0 Then response.write formatnumber((monthprice/period-expense/period)/(monthprice/period)*100, 2) Else response.write "0.00"%></td>
                    <td width="100" align="center"><%=custname2%>&nbsp;</td>
                  </tr>
				<%
						objrs.movenext
					loop
					objrs.close
					set objrs = nothing
				%>
              </table>