<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim custcode : custcode = request("custcode")
	dim custcode2 : custcode2 = request("custcode2")
	dim searchstring : searchstring = request("searchstring")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	if Len(cmonth) = 1 then cmonth = "0" & cmonth
	dim custname : custname = request("custname")
	dim sdate : sdate = dateserial(cyear, cmonth, "01")
	dim edate : edate = dateadd("d", -1, dateadd("m", 1,  sdate))

	dim temp_filename : temp_filename = cyear&"."&cmonth&"옥외광고집행현황"
'
	Response.CacheControl  = "public"
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader  "Content-Disposition" , "attachment; filename="&temp_filename&".xls"
%>

<%
	dim objrs, sql
	sql = "select title, locate, side, qty, unit, m.idx, standard, quality, firstdate, startdate, enddate, isnull(totalprice,0) as totalprice, isnull(monthprice,0) as monthprice, isnull(expense,0) as expense, seqname, thema, custname, medname from ( select m.contidx, m.title, m2.locate, d.idx, d.side, a.qty, m2.unit, d.standard, d.quality, m.firstdate, m.startdate, m.enddate, a.monthprice, a.expense,j2.seqname, j.thema, c.custname, c2.custname as medname from dbo.wb_contact_mst m left outer join dbo.wb_contact_md m2 on m.contidx = m2.contidx left outer join dbo.wb_contact_md_dtl d on m2.sidx = d.sidx left outer join dbo.wb_contact_md_dtl_account a on d.idx = a.idx and a.cyear = '"&cyear&"' and a.cmonth = '"&cmonth&"' left outer join dbo.sc_cust_temp c on c.custcode = m.custcode left outer join dbo.sc_cust_temp c2 on c2.custcode = m2.medcode left outer join dbo.wb_jobcust j on a.jobidx = j.jobidx left outer join dbo.sc_jobcust j2 on j.seqno = j2.seqno where m.canceldate >= '" & sdate & "' and m.custcode like '" & custcode2 &"%' and c.highcustcode like '" & custcode &"%' ) as m left outer join ( select d_.idx, sum(monthprice) as totalprice from dbo.wb_contact_mst m_ left outer join dbo.wb_contact_md m2_ on m_.contidx = m2_.contidx left outer join dbo.wb_contact_md_dtl d_ on m2_.sidx = d_.sidx left outer join dbo.wb_contact_md_dtl_account a_ on d_.idx = a_.idx group by d_.idx) as d on m.idx = d.idx  order by custname "

	call get_recordset(objrs, sql)

	dim total_price, total_monthprice, total_expense, total_income, total_incomeratio
%>
<style type="text/css">
	body {
		background-color:transparent;
		font-size:12px;
	}
	.trhd {
		font-size:12px;
		height: 30px;
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
		font-weight: bolder;
	}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr class="trhd">
    <td rowspan="2" align="center">매체명</td>
    <td rowspan="2" align="center">장소</td>
    <td rowspan="2" align="center">수량(면)</td>
    <td rowspan="2" align="center">면</td>
    <td rowspan="2" align="center">규격/재질</td>
    <td rowspan="2" align="center">최초<br>
      계약일자</td>
    <td colspan="2" align="center">계약기간</td>
    <td rowspan="2" align="center">총광고료</td>
    <td rowspan="2" align="center">월광고료</td>
    <td rowspan="2" align="center">월청구금액</td>
    <td rowspan="2" align="center">내수액</td>
    <td rowspan="2" align="center">내수율</td>
    <td rowspan="2" align="center">광고내용</td>
    <td rowspan="2" align="center">관련부서</td>
    <td rowspan="2" align="center">매채사</td>
  </tr>
  <tr class="trhd">
    <td align="center">시작일</td>
    <td align="center">종료일</td>
  </tr>
  <%
	do until objrs.eof
	%>
  <tr class="trbd">
    <td><%=objrs("title")%></td>
    <td><%=objrs("locate")%></td>
    <td><%=objrs("qty")%> <%if not isnull(objrs("unit")) then response.write "(" & objrs("unit") &")"%></td>
    <td><%=objrs("side")%>&nbsp;</td>
    <td><%=objrs("standard")%> <%if not isnull(objrs("quality")) then response.write "(" & objrs("quality") &")"%> </td>
    <td><%=objrs("firstdate")%></td>
    <td><%=objrs("startdate")%></td>
    <td><%=objrs("enddate")%></td>
    <td><%=formatnumber(objrs("totalprice"),0)%></td>
    <td><%=formatnumber(objrs("monthprice"),0)%></td>
    <td><%=formatnumber(objrs("expense"),0)%></td>
    <td><%=formatnumber(objrs("monthprice") - objrs("expense"),0)%></td>
    <td><%if objrs("monthprice") <> 0 then response.write formatnumber((objrs("monthprice")- objrs("expense"))/objrs("monthprice")*100,2)  else response.write formatnumber(0,2)%></td>
    <td><%=objrs("seqname")%></td>
    <td><%=objrs("custname") %></td>
    <td><%=objrs("medname")%></td>
  </tr>
  <%
	total_price = total_price + objrs("totalprice")
	total_monthprice = total_monthprice + objrs("monthprice")
	total_expense = total_expense + objrs("expense")


	total_income = total_monthprice - total_expense
	if total_income = 0 then
		total_incomeratio = 0
	else
		total_incomeratio = total_income / total_monthprice * 100
	end if
	objrs.movenext
	Loop


	objrs.close
	set objrs = nothing
  %>
  <tr class="trbd2">
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td><%=formatnumber(total_price,0)%></td>
    <td><%=formatnumber(total_monthprice,0)%></td>
    <td><%=formatnumber(total_expense,0)%></td>
    <td><%=formatnumber(total_income,0)%></td>
    <td><%=formatnumber(total_incomeratio,2)%></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
</table>
</body>