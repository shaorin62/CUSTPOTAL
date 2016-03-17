
<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
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
<%

	dim cyear : cyear = cstr(request.querystring("cyear"))
	dim cmonth : cmonth = cstr(request.querystring("cmonth"))
	dim cyear2 : cyear2 = cstr(request.querystring("cyear2"))
	dim cmonth2 : cmonth2 = cstr(request.querystring("cmonth2"))
	dim custcode2 : custcode2 = cstr(request.querystring("custcode2"))
	dim initpage : initpage  = cstr(request.querystring("initpage"))

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
	Response.AddHeader  "Content-Disposition" , "attachment; filename=CATV_NewMedia내역.xls"

	dim objrs, sql


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
	<% end if %>