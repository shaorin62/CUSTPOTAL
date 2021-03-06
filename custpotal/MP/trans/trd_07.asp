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
		<TD ><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle" id="subtitle"> 매체별 집행내역</span></TD>
		<TD  align="right" valign="top">  <span class="navigator"  id="navigate">광고비집행 &gt;  매체별 집행내역</span></TD>
	</TR>
</TABLE>
</div>

<!-- 검색 영역 -->
	<Div id='searchtag' style='margin-top:10px;'>
	<table width="1030" height="35" border="0" cellpadding="0" cellspacing="0" >
	   <tr>
		 <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
		 <td width="80%" align="left" background="/images/bg_search.gif"><%=search_cyear2cmonth2(cyear, cmonth, cyear2, cmonth2, custcode2)%> <A HREF="#" ><img src="/images/btn_search.gif" width="39" height="20" align="absmiddle" border="0" onclick="getdata('trd_07.asp'); return false;"></A></td>
		 <td width="20%" align="right" background="/images/bg_search.gif"><img src="/images/btn_excel.gif" width="85" height="22" align="absmiddle" vspace="5" class="stylelink" onclick="get_excel_sheet('trans_07.asp');"></td>
		 <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
	   </tr>
	</table>
	</div>

<!-- 컨텐츠 영역 -->
<%

	 if initpage = 1 then

	sql = " select case when m.med_flag = 'B' then '신문'  when m.med_flag = 'C' then '잡지' end as medflag, c2.custname ,  isnull(sum(case when substring(m.yearmon, 5,2) = '01' then isnull(amt,0) else 0 end),0) as 'A01', isnull(sum(case when substring(m.yearmon, 5,2) = '02' then isnull(amt,0) else 0 end),0) as 'A02',  isnull(sum(case when substring(m.yearmon, 5,2) = '03' then isnull(amt,0) else 0 end),0) as 'A03',  isnull(sum(case when substring(m.yearmon, 5,2) = '04' then isnull(amt,0) else 0 end),0) as 'A04',  isnull(sum(case when substring(m.yearmon, 5,2) = '05' then isnull(amt,0) else 0 end),0) as 'A05',  isnull(sum(case when substring(m.yearmon, 5,2) = '06' then isnull(amt,0) else 0 end),0) as 'A06',  isnull(sum(case when substring(m.yearmon, 5,2) = '07' then isnull(amt,0) else 0 end),0) as 'A07',  isnull(sum(case when substring(m.yearmon, 5,2) = '08' then isnull(amt,0) else 0 end),0) as 'A08',  isnull(sum(case when substring(m.yearmon, 5,2) = '09' then isnull(amt,0) else 0 end),0) as 'A09',  isnull(sum(case when substring(m.yearmon, 5,2) = '10' then isnull(amt,0) else 0 end),0) as 'A10',  isnull(sum(case when substring(m.yearmon, 5,2) = '11' then isnull(amt,0) else 0 end),0) as 'A11',  isnull(sum(case when substring(m.yearmon, 5,2) = '12' then isnull(amt,0) else 0 end),0) as 'A12',  sum(isnull(amt,0)) as 'TOTAL' from dbo.md_report_mst_v m  inner join dbo.sc_cust_hdr c on m.clientcode = c.highcustcode  inner join dbo.sc_cust_dtl c2 on m.medcode = c2.custcode where m.med_flag in ('b', 'C')  and m.yearmon between '"&yearmon&"' and '"&yearmon2&"' and m.clientcode LIKE '%"&custcode2&"%'  group by case when m.med_flag = 'B' then '신문'  when m.med_flag = 'C' then '잡지' end,c2.custname  with rollup "


	call get_recordset(objrs, sql)

	Dim custname, medflag, A01, A02, A03, A04, A05, A06, A07, A08, A09, A10, A11, A12, total, prev_medflag, prev_custname, prev_seqname
	If Not objrs.eof Then
		Set medflag = objrs("medflag")
		Set custname = objrs("custname")
		Set A01 = objrs("A01")
		Set A02 = objrs("A02")
		Set A03 = objrs("A03")
		Set A04 = objrs("A04")
		Set A05 = objrs("A05")
		Set A06 = objrs("A06")
		Set A07 = objrs("A07")
		Set A08 = objrs("A08")
		Set A08 = objrs("A08")
		Set A09 = objrs("A09")
		Set A10 = objrs("A10")
		Set A11 = objrs("A11")
		Set A12 = objrs("A12")
		Set total = objrs("total")
	End if

	if custcode2 = "" then custcode2 = null

%>
<div id='#contents' style='margin-top:10px;width:1030px;overflow-x:scroll;'>

<link href="/style.css" rel="stylesheet" type="text/css">
			   <table width="1420" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" >
                      <tr class="trhd">
                        <td width="90" align="center">구분</td>
                        <td width="150" align="center">매체</td>
                        <td width="90" align="center" >1월</td>
                        <td width="90" align="center">2월</td>
                        <td width="90" align="center">3월</td>
                        <td width="90" align="center">4월</td>
                        <td width="90" align="center">5월</td>
                        <td width="90" align="center">6월</td>
                        <td width="90" align="center">7월</td>
                        <td width="90" align="center">8월</td>
                        <td width="90" align="center">9월</td>
                        <td width="90" align="center">10월</td>
                        <td width="90" align="center">11월</td>
                        <td width="90" align="center">12월</td>
                        <td width="90" align="center">계</td>
                      </tr>
				<!--  -->
				<% do until objrs.eof 	%>
				<% If IsNull(medflag) And IsNull(custname) Then %>
                  <tr  class="trbd" bgcolor="#FFFFC1" >
                        <td width="220" align="center" colspan="2">&nbsp;&nbsp;합계</td>
				  <% ElseIf Not IsNull(medflag) And IsNull(custname) then%>
                  <tr  class="trbd" bgcolor="#CCFFFF" >
                        <td width="220" align="center" colspan="2">&nbsp;&nbsp;<%=medflag%> 소계</td>
				<%Else %>
                  <tr  class="trbd" bgcolor="#FFFFFF" >
                        <td width="90" align="left" >&nbsp;&nbsp;<%If prev_medflag <> medflag Then response.write medflag Else response.write "&nbsp;"%></td>
                        <td width="150" align="left">&nbsp;&nbsp;<%=custname%></td>
				  <% End if%>
                        <td align="right" ><%If A01.value <> "0" Then response.write FormatNumber(A01,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td align="right" ><%If A02.value <> "0" Then response.write FormatNumber(A02,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td align="right" ><%If A03.value <> "0" Then response.write FormatNumber(A03,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td align="right" ><%If A04.value <> "0" Then response.write FormatNumber(A04,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td align="right" ><%If A05.value <> "0" Then response.write FormatNumber(A05,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td align="right" ><%If A06.value <> "0" Then response.write FormatNumber(A06,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td align="right" ><%If A07.value <> "0" Then response.write FormatNumber(A07,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td align="right" ><%If A08.value <> "0" Then response.write FormatNumber(A08,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td align="right" ><%If A09.value <> "0" Then response.write FormatNumber(A09,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td align="right" ><%If A10.value <> "0" Then response.write FormatNumber(A10,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td align="right" ><%If A11.value <> "0" Then response.write FormatNumber(A11,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td align="right" ><%If A12.value <> "0" Then response.write FormatNumber(A12,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td align="right" ><%If total.value <> "0" Then response.write FormatNumber(total,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                      </tr>
				<%
						'End if
						prev_medflag = medflag
						objrs.movenext
					loop
					objrs.close
					set objrs = nothing
				%>
              </table>
</div>
<%end if %>