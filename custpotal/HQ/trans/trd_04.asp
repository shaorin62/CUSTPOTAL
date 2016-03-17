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
		<TD ><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle" id="subtitle"> 실집행 광고비</span></TD>
		<TD  align="right" valign="top">  <span class="navigator"  id="navigate">광고비집행 &gt;  실집행 광고비</span></TD>
	</TR>
</TABLE>
</div>

<!-- 검색 영역 -->
	<Div id='searchtag' style='margin-top:10px;'>
	<table width="1030" height="35" border="0" cellpadding="0" cellspacing="0" >
	   <tr>
		 <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
		 <td width="80%" align="left" background="/images/bg_search.gif"><%=search_cyear(cyear, custcode2)%> <A HREF="#" ><img src="/images/btn_search.gif" width="39" height="20" align="absmiddle" border="0" onclick="getdata('trd_04.asp'); return false;"></A></td>
		 <td width="20%" align="right" background="/images/bg_search.gif"><img src="/images/btn_excel.gif" width="85" height="22" align="absmiddle" vspace="5" class="stylelink" onclick="get_excel_sheet('trans_04.asp');"></td>
		 <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
	   </tr>
	</table>
	</div>

<!-- 컨텐츠 영역 -->
<%

	if initpage = 1 then
	sql = " select isnull(custname, 'Z') as custname, seqname, x.medflag, custpart,  sum(P01) as 'P01', sum(P02) as 'P02', sum(P03) as 'P03',  sum(P04) as 'P04', sum(P05) as 'P05', sum(P06) as 'P06', sum(P07) as 'P07', sum(P08) as 'P08', sum(P09) as 'P09', sum(P10) as 'P10', sum(P11) as 'P11', sum(P12) as 'P12', sum(TOTAl) as 'TOTAL' from dbo.sc_cust_hdr c inner join   (select isnull(j.seqname, 'Z') as seqname, isnull(case when m.med_flag = '01' then 'TV' when m.med_flag in ('02','03') then 'Radio' end,'A') as medflag, case when p.custpart = 'Z' then 'Others' else ' ' + p.custpart end as custpart, m.clientcode , isnull(sum(case when substring(m.yearmon, 5,2) = '01' then isnull(amt,0) else 0 end),0) as 'P01' , isnull(sum(case when substring(m.yearmon, 5,2) = '02' then isnull(amt,0) else 0 end),0) as 'P02' , isnull(sum(case when substring(m.yearmon, 5,2) = '03' then isnull(amt,0) else 0 end),0) as 'P03' , isnull(sum(case when substring(m.yearmon, 5,2) = '04' then isnull(amt,0) else 0 end),0) as 'P04' , isnull(sum(case when substring(m.yearmon, 5,2) = '05' then isnull(amt,0) else 0 end),0) as 'P05' , isnull(sum(case when substring(m.yearmon, 5,2) = '06' then isnull(amt,0) else 0 end),0) as 'P06' , isnull(sum(case when substring(m.yearmon, 5,2) = '07' then isnull(amt,0) else 0 end),0) as 'P07' , isnull(sum(case when substring(m.yearmon, 5,2) = '08' then isnull(amt,0) else 0 end),0) as 'P08' , isnull(sum(case when substring(m.yearmon, 5,2) = '09' then isnull(amt,0) else 0 end),0) as 'P09' , isnull(sum(case when substring(m.yearmon, 5,2) = '10' then isnull(amt,0) else 0 end),0) as 'P10' , isnull(sum(case when substring(m.yearmon, 5,2) = '11' then isnull(amt,0) else 0 end),0) as 'P11' , isnull(sum(case when substring(m.yearmon, 5,2) = '12' then isnull(amt,0) else 0 end),0) as 'P12' , sum(isnull(amt,0)) as 'TOTAL'  from dbo.md_report_mst_v m inner join dbo.vw_cust_part p on m.medcode = p.custcode left outer join dbo.sc_subseq_dtl j on j.seqno = m.subseq where m.med_flag in ('01', '02', '03') and substring(m.yearmon , 1, 4) = '"&cyear&"' and m.clientcode LIKE '%"&custcode2&"%' and amt <> 0 group by j.seqname, m.clientcode , case when m.med_flag = '01' then 'TV' when m.med_flag in ('02','03') then 'Radio' end , custpart , m.clientcode ) as x on c.highcustcode = x.clientcode group by c.custname, seqname, x.medflag, custpart with rollup "

	call get_recordset(objrs, sql)

	Dim seqname, medflag, custpart, custname, P01, P02, P03, P04, P05, P06, P07, P08, P09, P10, P11, P12, total, prev_seqname, prev_medflag
	If Not objrs.eof Then
		Set seqname = objrs("seqname")
		Set medflag = objrs("medflag")
		Set custpart = objrs("custpart")
		Set custname = objrs("custname")
		Set P01 = objrs("P01")
		Set P02 = objrs("P02")
		Set P03 = objrs("P03")
		Set P04 = objrs("P04")
		Set P05 = objrs("P05")
		Set P06 = objrs("P06")
		Set P07 = objrs("P07")
		Set P08 = objrs("P08")
		Set P09 = objrs("P09")
		Set P10 = objrs("P10")
		Set P11 = objrs("P11")
		Set P12 = objrs("P12")
		Set total = objrs("total")
	End if

	if custcode2 = "" then custcode2 = null

%>
<div id='#contents' style='margin-top:10px;width:1030px;overflow-x:scroll;'>

<link href="/style.css" rel="stylesheet" type="text/css">
			 <table width="1420" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" >
  <tr class="trhd">
    <td width="180" align="center">조정브랜드</td>
    <td width="90" align="center">구분</td>
    <td width="90" align="center">조정매체</td>
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
<!--  seqname, medflag, custpart, P01, P02, P03, P04, P05, P06, P07, P08, P09, P10, P11, P12, total, prev_seqname, prev_medflag -->
<%
	do until objrs.eof

	If Not (custname <> "Z" And Not IsNull(seqname) And IsNull(medflag) And IsNull(trim(custpart)) ) then
		If custname = "Z"  And  IsNull(seqname) And  IsNull(medflag) And IsNull(trim(custpart)) Then %> <!-- 총합계 -->
		  <tr class="trbd" bgcolor="#FFFFC1" >
			<td width="360" align="center" colspan="3">&nbsp;&nbsp;총합계</td>
		<% Elseif custname <> "Z" And Not IsNull(seqname) And Not IsNull(medflag) And IsNull(trim(custpart)) Then%>
		  <tr class="trbd" bgcolor="#FFDFDF" >
			<td width="180" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;<%If prev_seqname <> seqname Then if seqname ="Z" then response.write "&nbsp;" else response.write seqname end if  Else response.write "&nbsp;" end if%></td>
			<td width="180" align="center" colspan="2">&nbsp;&nbsp;<%=medflag%> 요약 </td>
		<% ElseIf custname <> "Z" And IsNull(seqname) And IsNull(medflag) And IsNull(trim(custpart))Then%>
		  <tr class="trbd" bgcolor="#CCFFFF" >
			<td width="360" align="center" colspan="3">&nbsp;&nbsp;<%=custname%> 소계</td>
		<% Else %>
		  <tr class="trbd" bgcolor="#FFFFFF" >
			<td width="180" align="left" >&nbsp;&nbsp;<%If prev_seqname <> seqname Then if seqname ="Z" then response.write "&nbsp;" else response.write seqname end if  Else response.write "&nbsp;" end if%> </td>
			<td width="90" align="left">&nbsp;&nbsp;<%if prev_medflag <> medflag then response.write medflag else response.write "&nbsp;"%> </td>
			<td width="90" align="left">&nbsp;&nbsp;<%=trim(custpart)%></td>
		<% End If %>
			<td align="right" ><%If P01.value <> "0" Then response.write FormatNumber(P01,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td align="right" ><%If P02.value <> "0" Then response.write FormatNumber(P02,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td align="right" ><%If P03.value <> "0" Then response.write FormatNumber(P03,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td align="right" ><%If P04.value <> "0" Then response.write FormatNumber(P04,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td align="right" ><%If P05.value <> "0" Then response.write FormatNumber(P05,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td align="right" ><%If P06.value <> "0" Then response.write FormatNumber(P06,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td align="right" ><%If P07.value <> "0" Then response.write FormatNumber(P07,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td align="right" ><%If P08.value <> "0" Then response.write FormatNumber(P08,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td align="right" ><%If P09.value <> "0" Then response.write FormatNumber(P09,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td align="right" ><%If P10.value <> "0" Then response.write FormatNumber(P10,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td align="right" ><%If P11.value <> "0" Then response.write FormatNumber(P11,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td align="right" ><%If P12.value <> "0" Then response.write FormatNumber(P12,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td align="right" ><%If total.value <> "0" Then response.write FormatNumber(total,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
		  </tr>
<%
		'End if
			If custname <> "Z" And IsNull(seqname) And IsNull(medflag) And IsNull(trim(custpart)) Then
				prev_seqname = ""
				prev_medflag = ""
			else
				prev_seqname = seqname
				prev_medflag = medflag
			End if
		End if

		objrs.movenext
	loop
	objrs.close
	set objrs = nothing
%>
              </table>
</div>

<%end if%>