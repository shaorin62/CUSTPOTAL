<%@ language="vbscript" codepage="65001" %>
<!--#include virtual="/inc/func.asp" -->
<%


	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1000

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
%>


<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<div style='margin-top:10px;'>
<TABLE  width="100%">
	<TR>
		<TD ><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle" id="subtitle"> 광고주/CIC별 매체비</span></TD>
		<TD  align="right" valign="top">  <span class="navigator"  id="navigate">광고비집행 &gt;  광고주/CIC별 매체비</span></TD>
	</TR>
</TABLE>
</div>

<!-- 검색 영역 -->
	<Div id='searchtag' style='margin-top:10px;'>
	<table width="1030" height="35" border="0" cellpadding="0" cellspacing="0" >
	   <tr>
		 <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
		 <td width="80%" align="left" background="/images/bg_search.gif"><%=search_cyear(cyear,  custcode2)%> <A HREF="#" ><img src="/images/btn_search.gif" width="39" height="20" align="absmiddle" border="0" onclick="getdata('trd_11.asp'); return false;"></A></td>
		 <td width="20%" align="right" background="/images/bg_search.gif"><img src="/images/btn_excel.gif" width="85" height="22" align="absmiddle" vspace="5" class="stylelink" onclick="get_excel_sheet('trans_11.asp');"></td>
		 <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
	   </tr>
	</table>
	</div>

<!-- 컨텐츠 영역 -->
<%

	if initpage = 1 then

	sql = " select m.timcode, c.custname  from dbo.md_report_mst_v m inner join dbo.sc_cust_dtl c on m.timcode = c.custcode where m.med_flag='O' and substring(m.yearmon, 1, 4)='"&cyear&"' and m.clientcode LIKE '%"&custcode2&"%' group by m.timcode, c.custname order by m.timcode  "



	call get_recordset(objrs, sql)

	if not objrs.eof then
	dim clientsubcode, clientsubname, tRow
	tRow = objrs.recordcount+1
	hearder = objrs.getRows()

	dim trRow : trRow = ubound(hearder, 1)    ' tr count
	dim tdRow : tdRow = ubound(hearder, 2)	'td count
%>
<div id='#contents' style='margin-top:10px;width:1030px;overflow-x:scroll;' ALIGN="LEFT">

<link href="/style.css" rel="stylesheet" type="text/css">
			<TABLE  width="<%=(tRow* 13 * 100)+350%>" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC">
<% for intTR = 0 to trRow %>
	<tr class="trhd2">
	<% if intTR = 0 then %>
		<TD rowspan="2" width="150" align="center">사이트명</TD>
	<% for intLoop = 1 to 12 %>
		<TD colspan="<%=tRow%>"  align="center"><%=intLoop%>월</TD>
	<% next %>
		<TD colspan="<%=tRow%>"  align="center">TOTAL</TD>
		<TD rowspan="2" width="200"  align="center"> 청구지 </TD>
	<% else%>
	<% for intLoop = 1 to 13 %>
	<% for intTD = 0 to tdRow %>
		<TD width="100" align="center"><%=hearder(intTR,intTD)%></TD>
	<% next %>
		<TD width="100" align="center">소계</TD>
	<% next %>
	<% end if %>
	</tr>
<% next %>

<!-- 데이터 릿스트 -->
<%
	sql = "select c2.custname "

	for intLoop = 1 to 12

	if len(intLoop) = 1 then intLoop = "0" & intLoop
		for intTD = 0 to tdRow
			sql = sql & ", sum(case when substring(m.yearmon, 5, 2) = '" & intLoop & "' and m.timcode = '" & hearder(0,intTD) & "' then isnull(amt, 0) else 0 end) as '" & intLoop &"' "
		next
			sql = sql & ", sum(case when substring(m.yearmon, 5, 2) = '" & intLoop & "' then isnull(amt, 0) else 0 end) as  '" & intLoop & " total'"
	next
		for intTD = 0 to tdRow
			sql = sql & ", sum(case when m.timcode = '" & hearder(0,intTD) & "' then isnull(amt, 0) else 0 end) as 'TOTAL'"
		next
			sql = sql & ", sum(isnull(amt, 0)) as  '" & intLoop & " total', max(c.companyname) "
	sql = sql & " from dbo.md_report_mst_v m inner join dbo.sc_cust_hdr c on m.real_med_code = c.highcustcode inner join dbo.sc_cust_dtl c2 on m.medcode = c2.custcode where m.med_flag='O' and substring(m.yearmon, 1, 4)='"&cyear&"' and m.clientcode LIKE '%"&custcode2&"%'  group by c2.custname  with rollup"




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
			response.write "합계"
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
<%end if%>
</div>
<%end if%>