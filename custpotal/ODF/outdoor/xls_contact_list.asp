<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim custcode : custcode = request("custcode")
	dim custcode2 : custcode2 = request("custcode2")
	dim searchstring : searchstring = request("searchstring")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	dim custname : custname = request("custname")

	dim temp_filename : temp_filename = cyear&"."&cmonth&"옥외광고집행현황"

	Response.CacheControl  = "public"
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader  "Content-Disposition" , "attachment; filename="&temp_filename&".xls"
%>

<%
	dim objrs, sql
	sql = "select m.title, m2.locate ,m2.qty , m2.unit, m2.side, m2.standard, m2.quality, m.firstdate, m.startdate, m.enddate , isnull(t.monthprice,0) as totalprice, isnull(sum(d.monthprice),0) as monthprice, isnull(sum(d.expense),0) as expense, j2.seqname, c.custname as custname2, c3.custname as custname3 from dbo.wb_contact_mst m inner join dbo.wb_contact_md m2 on m.contidx = m2.contidx inner join dbo.wb_contact_md_dtl d on m2.contidx = d.contidx and m2.sidx = d.sidx inner join dbo.sc_cust_temp c on m.custcode = c.custcode  inner join dbo.sc_cust_temp c2 on c.highcustcode = c2.custcode inner join dbo.wb_medium_mst m3 on m2.mdidx = m3.mdidx inner join dbo.sc_cust_temp c3 on m3.custcode = c3.custcode inner join dbo.wb_jobcust j on d.jobidx = j.jobidx inner join dbo.sc_jobcust j2 on j.seqno = j2.seqno left outer join dbo.vw_contact_totalprice t on m.contidx = t.contidx where c.highcustcode like '"&custcode&"%' and d.cyear = '"&cyear&"' and d.cmonth = '"&cmonth&"'  and m.title like '%"& searchstring &"%' group by m.title, m.firstdate, m.startdate, m.enddate, m.canceldate, t.monthprice , c3.custname, m2.side, m2.standard, m2.quality, j2.seqname , m2.locate , m2.qty, m2.unit, c.custname order by m.title "

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
<body  oncontextmenu="return false">
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
    <td rowspan="2" align="center">총광고료(원)</td>
    <td rowspan="2" align="center">월광고료(원)</td>
    <td rowspan="2" align="center">월청구금액(원)</td>
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
	Dim prev_title
	Dim prev_locate
	Dim prev_custname2
	do until objrs.eof
	%>
  <tr class="trbd">
    <td><%If prev_title <> objrs("title") Then response.write objrs("title")%></td>
    <td><%If prev_locate <> objrs("locate")  Then response.write objrs("locate")%></td>
    <td><%=objrs("qty")%>(<%=objrs("unit")%>)</td>
    <td><%=objrs("side")%>&nbsp;</td>
    <td><%=objrs("standard")%> (<%=objrs("quality")%>)</td>
    <td><%=objrs("firstdate")%></td>
    <td><%=objrs("startdate")%></td>
    <td><%=objrs("enddate")%></td>
    <td><%=formatnumber(objrs("totalprice"),0)%></td>
    <td><%=formatnumber(objrs("monthprice"),0)%></td>
    <td><%=formatnumber(objrs("expense"),0)%></td>
    <td><%=formatnumber(objrs("monthprice") - objrs("expense"),0)%></td>
    <td><%if objrs("monthprice") <> 0 then response.write formatnumber((objrs("monthprice")- objrs("expense"))/objrs("monthprice")*100,2)  else response.write formatnumber(0,2)%></td>
    <td><%=objrs("seqname")%></td>
    <td><%=objrs("custname2") %></td>
    <td><%=objrs("custname3")%></td>
  </tr>
  <%
	prev_title = objrs("title")
	prev_locate = objrs("locate")
	prev_custname2 = objrs("custname2")
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