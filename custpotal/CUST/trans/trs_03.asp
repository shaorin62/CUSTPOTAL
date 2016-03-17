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
		<TD ><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle" id="subtitle"> CATV/NEW MEDIA 내역</span></TD>
		<TD  align="right" valign="top">  <span class="navigator"  id="navigate">광고비집행 &gt;  CATV/NEW MEDIA 내역</span></TD>
	</TR>
</TABLE>
</div>

<!-- 검색 영역 -->
	<Div id='searchtag' style='margin-top:10px;'>
	<table width="1030" height="35" border="0" cellpadding="0" cellspacing="0" >
	   <tr>
		 <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
		 <td width="80%" align="left" background="/images/bg_search.gif"><%=search_cyear2cmonth2(cyear, cmonth, cyear2, cmonth2, custcode2)%> <A HREF="#" ><img src="/images/btn_search.gif" width="39" height="20" align="absmiddle" border="0" onclick="getdata('trs_03.asp'); return false;"></A></td>
		 <td width="20%" align="right" background="/images/bg_search.gif"><img src="/images/btn_excel.gif" width="85" height="22" align="absmiddle" vspace="5" class="stylelink" onclick="get_excel_sheet('public_03.asp');"></td>
		 <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
	   </tr>
	</table>
	</div>

<!-- 컨텐츠 영역 -->
<%

	if initpage = 1 then

	sql = "select isnull(c.yearmon, '총합계') as yearmon, c2.custname, isnull(sum(case when c.mpp = 'p00005' then isnull(amt,0) end),0) as 'P01'	, isnull(sum(case when c.mpp = 'p00007' then isnull(amt,0) end),0) as 'P02', isnull(sum(case when c.mpp = 'p00004' then isnull(amt,0) end),0) as 'P03', isnull(sum(case when isnull(c.mpp , '') not in('p00005','p00007','p00004') and c.medcode in('B00046', 'B00460','B00868')  then isnull(amt,0) end),0) as 'P04', isnull(sum(case when isnull(c.mpp , '') not in('p00005','p00007','p00004') and c.medcode in('B00047', 'B00461','B00517','B00869')  then isnull(amt,0) end),0) as 'P05' , isnull(sum(case when isnull(c.mpp , '') not in('p00005','p00007','p00004') and c.medcode not in ('B00046', 'B00460','B00047', 'B00461','B00868','B00517','B00869') then isnull(amt,0) end ),0) as 'OTH', sum(isnull(amt,0)) as 'TOTAL' from dbo.md_report_mst_v c inner join dbo.sc_cust_hdr c2 on c.clientcode = c2.highcustcode inner join dbo.sc_cust_dtl c3 on c.medcode = c3.custcode where  c.clientcode like '" & custcode2 & "%' and  yearmon between '"&yearmon&"' and '"&yearmon2&"' and med_flag ='A2'  group by c.yearmon, c2.custname with rollup "

	call get_recordset(objrs, sql)

	Dim cyearmon, custname, P01, P02, P03, P04, P05, OTH, total, prev
	If Not objrs.eof Then
		Set cyearmon = objrs("yearmon")
		Set custname = objrs("custname")
		Set P01 = objrs("P01")
		Set P02 = objrs("P02")
		Set P03 = objrs("P03")
		Set P04 = objrs("P04")
		Set P05 = objrs("P05")
		Set OTH = objrs("OTH")
		Set total = objrs("total")
	End if

	if custcode2 = "" then custcode2 = null

%>

<div id='#contents' style='margin-top:10px;'>

<link href="/style.css" rel="stylesheet" type="text/css">
				  <table width="1020" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" >
		  <tr  class="trhd">
			<td  rowspan="2" align="center" >구분</td>
			<td  align="center" >Category</td>
			<td colspan="3" align="center">케이블 TV</td>
			<td   align="center" >IPTV</td>
			<td  align="center"  >위성DMB</td>
			<td  rowspan="2" align="center" >Others</td>
			<td   rowspan="2" align="center" >총 집행 금액</td>
		  </tr>
		  <tr  class="trhd">
			<td  align="center">MPP</td>
			<td  align="center">CU 미디어</td>
			<td align="center">CJ 미디어</td>
			<td align="center">온미디어</td>
				<td  align="center">브로드앤TV</td>
				<td  align="center">TU</td>
			  </tr>
		<!--  -->
		<% do until objrs.eof 	%>
		<% If cyearmon = "총합계" Then %>
		  <tr  class="trbd" bgcolor="#FFFFC1" >
			<td width="240" align="center" colspan="2"> 총합계 </td>
			<td width="100" align="right" ><%If P01.value <> "0" Then response.write FormatNumber(P01,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If P02.value <> "0" Then response.write FormatNumber(P02,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If P03.value <> "0" Then response.write FormatNumber(P03,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If P04.value <> "0" Then response.write FormatNumber(P04,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If P05.value <> "0" Then response.write FormatNumber(P05,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If OTH.value <> "0" Then response.write FormatNumber(OTH,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If total.value <> "0" Then response.write FormatNumber(total,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
		  </tr>
		  <% ElseIf cyearmon <> "총합계" And IsNull(custname) Then %>
		  <tr  class="trbd" bgcolor="#CCFFFF" >
			<td width="100" align="center"><%If prev <> cyearmon Then response.write left(cyearmon,4) & chr(45) & right(cyearmon,2) Else response.write "&nbsp;"%></td>
			<td width="140" align="left" style="padding-left:10px;">TOTAL</td>
			<td width="100" align="right" ><%If P01.value <> "0" Then response.write FormatNumber(P01,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If P02.value <> "0" Then response.write FormatNumber(P02,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If P03.value <> "0" Then response.write FormatNumber(P03,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If P04.value <> "0" Then response.write FormatNumber(P04,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If P05.value <> "0" Then response.write FormatNumber(P05,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If OTH.value <> "0" Then response.write FormatNumber(OTH,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If total.value <> "0" Then response.write FormatNumber(total,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
		  </tr>
		  <tr  class="trbd" bgcolor="#FFFFFF" >
			<td width="100" align="center"><%If prev <> cyearmon Then response.write left(cyearmon,4) & chr(45) & right(cyearmon,2) Else response.write "&nbsp;"%></td>
			<td width="140" align="left" style="padding-left:10px;">%</td>
			<td width="100" align="right" ><%If P01.value <> "0" and total.value <> "0" then  response.write replace(FormatPercent(CDBL(P01)/Cdbl(total),0),"%","") else response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If P02.value <> "0" and total.value<> "0" Then response.write replace(FormatPercent(CDBL(P02)/Cdbl(total),0),"%","") Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If P03.value <> "0" and total.value<> "0" Then response.write replace(FormatPercent(CDBL(P03)/Cdbl(total),0),"%","") Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If P04.value <> "0" and total.value<> "0" Then response.write replace(FormatPercent(CDBL(P04)/Cdbl(total),0),"%","") Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If P05.value <> "0" and total.value<> "0" Then response.write replace(FormatPercent(CDBL(P05)/Cdbl(total),0),"%","") Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If OTH.value <> "0" and total.value<> "0" Then response.write replace(FormatPercent(CDBL(OTH)/Cdbl(total),0),"%","") Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If total.value <> "0" and total.value<> "0" Then response.write replace(FormatPercent(Cdbl(total)/Cdbl(total),0),"%","") Else  response.write "-"%>&nbsp;&nbsp;</td>
		  </tr>
		  <tr  class="trbd" bgcolor="#FFFFC1" >
			<td width="100" align="center"><%If prev <> cyearmon Then response.write left(cyearmon,4) & chr(45) & right(cyearmon,2) Else response.write "&nbsp;"%></td>
			<td width="140" align="left" style="padding-left:10px;">On Media vs. CJ Media </td>
			<td width="100" align="right" >-</td>
			<td width="100" align="right" ><%If P02.value <> "0" and  P03.value <> "0"  Then response.write replace(FormatPercent(CDBL(P02)/(CDBL(P02)+CDBL(P03)),0),"%","") Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If P02.value <> "0" and  P03.value <> "0"  Then response.write replace(FormatPercent(CDBL(P03)/(CDBL(P02)+CDBL(P03)),0),"%","") Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" >-</td>
			<td width="100" align="right" >-</td>
			<td width="100" align="right" >-</td>
			<td width="100" align="right" >-</td>
		  </tr>
		  <%Else %>
		  <tr  class="trbd" bgcolor="#FFFFFF" >
			<td width="100" align="center"><%If prev <> cyearmon Then response.write left(cyearmon,4) & chr(45) & right(cyearmon,2) Else response.write "&nbsp;"%></td>
			<td width="140" align="left" style="padding-left:10px;"><%=custname%></td>
			<td width="100" align="right" ><%If P01.value <> "0" Then response.write FormatNumber(P01,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If P02.value <> "0" Then response.write FormatNumber(P02,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If P03.value <> "0" Then response.write FormatNumber(P03,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If P04.value <> "0" Then response.write FormatNumber(P04,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If P05.value <> "0" Then response.write FormatNumber(P05,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If OTH.value <> "0" Then response.write FormatNumber(OTH,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td width="100" align="right" ><%If total.value <> "0" Then response.write FormatNumber(total,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
		  </tr>
		 <% End If %>
		<%
				'End if
				prev = cyearmon
				objrs.movenext
			loop
			objrs.close
			set objrs = nothing
		%>
           </table>
</div>

<%end if %>

