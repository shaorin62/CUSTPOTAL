<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim gotopage : gotopage = request.querystring("gotopage")
	if gotopage = "" then gotopage = 1
	dim searchstring : searchstring = request("txtsearchstring")
	dim custcode : custcode = request("selcustcode")
	dim custcode2 : custcode2 = request("selcustcode2")
	dim seqno : seqno = request("seljobcust")
	dim thema : thema = request("selthema")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	if cyear = "" then cyear = year(date)
	if cmonth = "" then cmonth = month(date)
	dim sdate : sdate = Dateserial(cyear, cmonth, "01")

	dim objrs, sql
	sql = "select contidx, title, locate, side, qty, unit, m.idx, standard, quality, firstdate, startdate, enddate, isnull(totalprice,0) as totalprice, isnull(monthprice,0) as monthprice, isnull(expense,0) as expense, seqname, thema, custname, medname from ( select m.contidx, m.title, m2.locate, d.idx, d.side, a.qty, m2.unit, d.standard, d.quality, m.firstdate, m.startdate, m.enddate, a.monthprice, a.expense,j2.seqname, j.thema, c.custname, c2.custname as medname from dbo.wb_contact_mst m left outer join dbo.wb_contact_md m2 on m.contidx = m2.contidx left outer join dbo.wb_contact_md_dtl d on m2.sidx = d.sidx left outer join dbo.wb_contact_md_dtl_account a on d.idx = a.idx and a.cyear = '"&cyear&"' and a.cmonth = '"&cmonth&"' left outer join dbo.sc_cust_temp c on c.custcode = m.custcode left outer join dbo.sc_cust_temp c2 on c2.custcode = m2.medcode left outer join dbo.wb_jobcust j on a.jobidx = j.jobidx left outer join dbo.sc_jobcust j2 on j.seqno = j2.seqno where m.canceldate >= '" & sdate & "' and m.custcode like '" & custcode2 & "%' and c.highcustcode like '" & custcode & "%' and j.seqno like '" & seqno & "%') as m left outer join ( select d_.idx, sum(monthprice) as totalprice from dbo.wb_contact_mst m_ left outer join dbo.wb_contact_md m2_ on m_.contidx = m2_.contidx left outer join dbo.wb_contact_md_dtl d_ on m2_.sidx = d_.sidx left outer join dbo.wb_contact_md_dtl_account a_ on d_.idx = a_.idx group by d_.idx) as d on m.idx = d.idx  "
	if thema <> "" then sql = sql & " where thema = '" & old_thema & "' "
	sql = sql & " order by enddate "

	call get_recordset(objrs, sql)

	dim contidx, title, startdate, enddate, side, seqname, standard, quality, monthprice, custname, canceldate, contactcancel

	dim cnt : cnt = objrs.recordcount

	if not objrs.eof Then
		set contidx = objrs("contidx")
		set title = objrs("title")
		set startdate = objrs("startdate")
		set enddate = objrs("enddate")
		set side = objrs("side")
		set seqname = objrs("seqname")
		set standard = objrs("standard")
		set quality = objrs("quality")
		set monthprice = objrs("monthprice")
		set custname = objrs("custname")
		set thema = objrs("thema")
	end if
	if seqno = "" then seqno = "0"

	Response.CacheControl  = "public"
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader  "Content-Disposition" , "attachment; filename="&cyear&"."&cmonth&"브랜드별 집행현황.xls"
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
<body>
<table width="1024" border="0" cellspacing="0" cellpadding="0" class="header">
<tr class="trhd">
  <td align="center" >매체명</td>
  <td align="center">시작일</td>
  <td align="center">종료일</td>
  <td align="center">면</td>
  <td align="center">소재명</td>
  <td align="center">브랜드</td>
  <td align="center">규격 / 재질</td>
  <td align="center">월광고료</td>
  <td align="center">운영팀</td>
</tr>
<% do until objrs.eof	%>
<tr  class="trbd">
  <td align="left"><%=title %></td>
  <td align="center"><%=startdate%></td>
  <td align="center"><%=enddate%></td>
  <td align="center"><%=side%></td>
  <td align="left"><%=thema%></td>
  <td align="center"><%=seqname%></td>
  <td align="center"><%if not isnull(standard) then response.write standard %> <%if not isnull(quality) then response.write " / " & quality %></td>
  <td align="right"><%if not isnull(monthprice) then response.write formatnumber(monthprice,0) else response.write "0"%></td>
  <td align="center"><%=custname%></td>
</tr>
<%
		objrs.movenext
	loop
	objrs.close
	set objrs = nothing
%>
</table>
</body>