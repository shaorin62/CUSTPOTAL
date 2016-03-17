
<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim hearder' 테이블 헤어로 이용할 레코드셋 배열
	dim intTR, intTD' for 을 위한 일련번호
	Dim intLoop

	dim cyear : cyear = request.querystring("cyear")
	dim cmonth : cmonth = request.querystring("cmonth")
	dim cyear2 : cyear2 = request.querystring("cyear2")
	dim cmonth2 : cmonth2 = request.querystring("cmonth2")
	dim custcode2 : custcode2 = request.querystring("custcode2")
	dim initpage : initpage  = request.querystring("initpage")

	if cyear =  "" then cyear = Cstr(Year(date))
	if cmonth = "" then cmonth = Cstr(Month(Date))
	if cyear2 =  "" then cyear2 = Cstr(Year(date))
	if cmonth2 = "" then cmonth2 = Cstr(Month(Date))

	if len(cmonth) = 1 then cmonth = "0"&cmonth
	if len(cmonth2) = 1 then cmonth2 = "0"&cmonth2


	dim yearmon : yearmon = cyear&cmonth
	dim yearmon2 : yearmon2 = cyear2&cmonth2

	Response.CacheControl  = "public"
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader  "Content-Disposition" , "attachment; filename=월별큐시트.xls"

	dim objrs, sql


	sql = "select subseq, dbo.pd_jobcust_name_fun(subseq) as subname,  count(campaign_code) as subcount from (select subseq, campaign_code from dbo.md_internet_medium m where yearmon ='" & cyear&cmonth & "'  and m.clientcode LIKE '%" & custcode2 & "%' group by subseq, campaign_code) as a group by subseq"


	call get_recordset(objrs, sql)

	if not objrs.eof then
	dim hdrow : hdrow = objrs.getrows
	dim hdrowcount : hdrowcount = ubound(hdrow, 1)
	dim hdcolcount : hdcolcount = ubound(hdrow, 2)
	objrs.close


	sql = "select b.seqname, c.campaign_name,  substring(m.tbrdstdate,5,2) + '/' +  substring(m.tbrdstdate,7,2) + '~' +  substring(m.tbrdeddate,5,2) + '/' +  substring(m.tbrdeddate,7,2), m.subseq,  m.campaign_code  from dbo.md_internet_medium m  inner join dbo.sc_subseq_dtl b  on b.seqno = m.subseq  inner join dbo.md_internet_campaign c on m.campaign_code = c.campaign_code  where m.clientcode LIKE '%" & custcode2 & "%' and m.yearmon = '" & cyear&cmonth & "'  group by b.seqname, c.campaign_name, m.tbrdstdate, m.tbrdeddate,m.subseq, m.campaign_code  order by m.subseq, m.campaign_code "

	call get_recordset(objrs, sql)
	if not objrs.eof then
		hearder = objrs.getrows()
	dim tmpRow
	dim trRow : trRow = ubound(hearder, 1)
	dim tdRow : tdRow = ubound(hearder, 2)
	dim intTD1
	dim prev_brand, colcount

	'hearder, 1 = tr
	'hreader, 2 = td
%>

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
 <TABLE  width="<%=((tdRow+4) * 100)+150%>" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC">
<% for intTR = 0 to trRow-2 %>
<TR class="trhd2">


	<% if intTR = 0 then %>
		<TD rowspan="3" width="150" align="center">사이트명</TD>
		<% for intLoop = 0 to hdcolcount %>
			<TD    align="center" colspan="<%=hdrow(hdrowcount, intLoop)%>"><%=hdrow(1, intLoop)%> (<%=hdrow(hdrowcount, intLoop)%>건)</TD>
		<% next %>
		<TD rowspan="3" width="100"  align="center"> 소 계 </TD>
		<TD rowspan="3" width="200"  align="center"> 청구지 </TD>
	<%else%>
	<% for intTD = 0 to tdRow %>
				<TD  width="100" align="center"><%=hearder(intTR,intTD)%></TD>
	<%next	%>
	<% end if %>
</TR>
<% next%>
<!-- 데이터 릿스트 -->
<%
	sql = "select m.subseq, m.campaign_code  from dbo.md_internet_medium m  inner join dbo.sc_subseq_dtl b on b.seqno = m.subseq  inner join dbo.md_internet_campaign c on m.campaign_code = c.campaign_code  where m.clientcode LIKE '%" & custcode2 & "%' and m.yearmon = '" & cyear&cmonth & "'  group by m.subseq, m.campaign_code order by m.subseq, m.campaign_code "

	call get_recordset(objrs, sql)

	sql = "select c.custname"
	do until objrs.eof
		sql = sql & ", sum(case when m.campaign_code = '"&objrs("campaign_code")&"' and m.subseq = '"&objrs("subseq")&"' then isnull(amt,0) else 0 end)"
		objrs.movenext
		loop
	sql = sql & ", sum(isnull(amt,0)), max(c2.companyname)"
	sql = sql & " from dbo.md_internet_medium m inner join dbo.sc_cust_dtl c on m.medcode = c.custcode inner join dbo.sc_cust_hdr c2 on m.real_med_code = c2.highcustcode where m.clientcode LIKE '%"&custcode2&"%'  and  m.yearmon = '"&cyear&cmonth&"' group by c.custname with rollup "

	objrs.close

	call get_recordset(objrs, sql)

	dim tblData : tblData = objrs.getrows
	trRow = ubound(tblData, 2)
	tdRow = ubound(tblData, 1)
%>
<% for intTR = 0 to trRow %>
<% if intTR = trRow then %>
	<TR  class="trbd" bgcolor="#FFFFC1" >
<% else %>
	<TR  class="trbd" bgcolor="#FFFFFF" >
<% end if %>
		<% for intTD = 0 to tdRow %>

<TD   style="<%if intTD =0 then response.write "padding-left:10px;" else response.write "padding-right:10px;text-align:right"%>">
<%
	if intTD > 0 and intTD < tdRow then
		response.write formatnumber(tblData(intTD,intTR),0)
	else
		if isnull(tblData(intTD,intTR)) then
			response.write "소계"
		elseif intTR = trRow and intTD = tdRow then
			response.write ""
		else
			response.write tblData(intTD,intTR)
		end if
	end if
%> </TD>
		<% next %>
	</TR>
<% next %>
<!-- page end  -->
<% end if %>
<% end if %>