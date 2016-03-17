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
	dim c_date : c_date = Dateserial(cyear, cmonth, "01")

	dim objrs, sql
	sql = "select p.contidx, p.title, p.startdate, p.enddate, p.canceldate, m.side, j2.seqname, m.standard, m.quality, isnull(d.monthprice,0) as monthprice ,p.custcode as custcode2, c.custname as custname2, d.contactcancel, j.thema  from dbo.wb_contact_mst p left outer join dbo.wb_contact_md m on p.contidx = m.contidx inner join dbo.wb_contact_md_dtl d on m.contidx = d.contidx and m.sidx = d.sidx and d.cyear = '"&cyear&"' and d.cmonth='"&cmonth&"'  left outer join dbo.sc_cust_temp c on p.custcode = c.custcode    inner join dbo.sc_cust_temp c2 on c.highcustcode = c2.custcode  inner join dbo.wb_jobcust j on d.jobidx = j.jobidx  left outer join dbo.sc_jobcust j2 on j.seqno = j2.seqno where c2.custcode like '" & custcode &"%' and p.custcode like '"&custcode2&"%' and p.cancel = 0  and j.seqno like '"&seqno&"%' and thema like '%" & thema & "%' "
	call get_recordset(objrs, sql)

	dim contidx, title, startdate, enddate, side, seqname, standard, quality, monthprice, custname2,canceldate, contactcancel

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
		set custname2 = objrs("custname2")
		set canceldate = objrs("canceldate")
		set contactcancel = objrs("contactcancel")
		set thema = objrs("thema")
	end if

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
<body  oncontextmenu="return false">
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
  <td align="left"><% if not isnull(canceldate) then response.write "<del>"&title&"</del>" else response.write title %></td>
  <td align="center"><%=startdate%></td>
  <td align="center"><%=enddate%></td>
  <td align="center"><%=side%></td>
  <td align="left"><%=thema%></td>
  <td align="center"><%=seqname%></td>
  <td align="center"><%if not isnull(standard) then response.write standard %> <%if not isnull(quality) then response.write " / " & quality %></td>
  <td align="right"><%if not isnull(monthprice) then response.write formatnumber(monthprice,0) else response.write "0"%></td>
  <td align="center"><%=custname2%></td>
</tr>
<%
		objrs.movenext
	loop
	objrs.close
	set objrs = nothing
%>
</table>
</body>