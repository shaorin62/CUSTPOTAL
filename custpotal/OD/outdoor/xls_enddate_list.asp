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
	if Len(cmonth) = 1 then cmonth = "0"&cmonth
	if Len(cmonth2) = 1 then cmonth2= "0"&cmonth2
	dim sdate : sdate = Dateserial(cyear, cmonth, "01")
	dim edate : edate = dateadd("d", -1, dateAdd("m", 1, Dateserial(cyear2, cmonth2, "01")))

	dim objrs, sql
	sql = "select m.contidx, title, firstdate, startdate, enddate, isnull(totalprice,0) as totalprice, monthprice, expense, custname, medname,canceldate from ( select m.contidx, m.title, m.firstdate, m.startdate, m.enddate, isnull(sum(a.monthprice),0) as monthprice, isnull(sum(a.expense),0) as expense, c.custname, c2.custname as medname ,m.canceldate from dbo.wb_contact_mst m inner join dbo.sc_cust_temp c on c.custcode = m.custcode left outer join dbo.wb_contact_md m2 on m.contidx = m2.contidx  left outer join dbo.sc_cust_temp c2 on c2.custcode = m2.medcode left outer join dbo.wb_contact_md_dtl d on m2.sidx = d.sidx left outer join dbo.wb_contact_md_dtl_account a on d.idx = a.idx   where m.custcode like '"&custcode2&"%' and c.highcustcode like '"&custcode&"%'   and m.title like '%"&searchstring&"%' and m.canceldate <= m.enddate group by m.contidx, m.title, m.firstdate, m.startdate, m.enddate, c.custname, m.canceldate, c2.custname ) as m left outer join (select m.contidx, sum(a.monthprice) as totalprice from dbo.wb_contact_md m inner join dbo.wb_contact_md_dtl d on m.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on d.idx = a.idx group by m.contidx ) as d on m.contidx = d.contidx where (enddate between '" & sdate & "' and '" & edate & "') and m.canceldate >= '" & sdate & "'  order by enddate"

	'response.write sql

	call get_recordset(objrs, sql)

	dim cnt, contidx, title, firstdate, startdate, enddate, period, monthprice, expense, income, incomeratio, custname, medname, totalprice,canceldate

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
		set custname = objrs("custname")
		set medname = objrs("medname")
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
<tr class="trhd">
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
                        <td width="100" align="center">매체사</td>
                      </tr>
	     <%
			do until objrs.eof
			if day(startdate) = "1" then
				period = datediff("m", startdate, enddate)+1
			else
				period = datediff("m", startdate, enddate)
			end if

			if period = 0 then period = 1

			dim monPrice, expPrice, incPrice, incRate

			if period <> 0 then
				monPrice = monthprice/period
				expPrice = expense/period
				incPrice = monPrice-expPrice
				if monPrice <> 0 then 	incRate = incPrice/monPrice*100 	else incRate = "0.00" end if
			else
				monPrice = 0
				expPrice =0
				incPrice = 0
				incRate = 0
			end if
		%>
<tr  class="trbd">
                    <td width="220" align="left"><%=title%></td>
                    <td width="75" align="center"><%=firstdate%></td>
                    <td width="75" align="center"><%=startdate%></td>
                    <td width="75" align="center"><%=enddate%></td>
                    <td width="80" align="right"><%If Not IsNull(totalprice) Then response.write formatnumber(totalprice,0) Else response.write "0"%></td>
                    <td width="80" align="right"><%=formatNumber(monPrice,0)%></td>
                    <td width="80" align="right"><%=formatNumber(expPrice,0)%></td>
                    <td width="80" align="right"><%=formatNumber(incPrice,0)%></td>
                    <td width="50" align="right"><%=formatNumber(incRate,2)%></td>
                    <td width="100" align="center"><%=custname%>&nbsp;</td>
                    <td width="100" align="center"><%=medname%>&nbsp;</td>
                  </tr>
				<%
						objrs.movenext
					loop
					objrs.close
					set objrs = nothing
				%>
              </table>