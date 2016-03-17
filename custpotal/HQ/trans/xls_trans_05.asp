
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
	Response.AddHeader  "Content-Disposition" , "attachment; filename=케이블월별매체별실집행광고비.xls"

	dim objrs, sql


	sql = "select isnull(dbo.sc_get_custname_fun(m.timcode),'z') as custname , isnull(c2.custname,'z') as custname2 , isnull(sum(case when substring(m.yearmon, 5,2) = '01' then isnull(amt,0) else 0 end),0) as 'A01' , isnull(sum(case when substring(m.yearmon, 5,2) = '02' then isnull(amt,0) else 0 end),0) as 'A02', isnull(sum(case when substring(m.yearmon, 5,2) = '03' then isnull(amt,0) else 0 end),0) as 'A03', isnull(sum(case when substring(m.yearmon, 5,2) = '04' then isnull(amt,0) else 0 end),0) as 'A04', isnull(sum(case when substring(m.yearmon, 5,2) = '05' then isnull(amt,0) else 0 end),0) as 'A05', isnull(sum(case when substring(m.yearmon, 5,2) = '06' then isnull(amt,0) else 0 end),0) as 'A06', isnull(sum(case when substring(m.yearmon, 5,2) = '07' then isnull(amt,0) else 0 end),0) as 'A07', isnull(sum(case when substring(m.yearmon, 5,2) = '08' then isnull(amt,0) else 0 end),0) as 'A08', isnull(sum(case when substring(m.yearmon, 5,2) = '09' then isnull(amt,0) else 0 end),0) as 'A09', isnull(sum(case when substring(m.yearmon, 5,2) = '10' then isnull(amt,0) else 0 end),0) as 'A10', isnull(sum(case when substring(m.yearmon, 5,2) = '11' then isnull(amt,0) else 0 end),0) as 'A11', isnull(sum(case when substring(m.yearmon, 5,2) = '12' then isnull(amt,0) else 0 end),0) as 'A12', sum(isnull(amt,0)) as 'TOTAL' from dbo.md_report_mst_v m inner join dbo.sc_cust_hdr c on m.timcode = c.highcustcode inner join dbo.sc_cust_dtl c2 on c2.custcode = m.medcode where substring(m.yearmon, 1, 4) = '"&cyear&"' and m.med_flag = 'A2' and m.clientcode LIKE '%"& custcode2 &"%' group by dbo.sc_get_custname_fun(m.timcode) , c2.custname with cube order by custname , custname2 "

	call get_recordset(objrs, sql)

	Dim custname, custname2, A01, A02, A03, A04, A05, A06, A07, A08, A09, A10, A11, A12, total, prev
	If Not objrs.eof Then
		Set custname = objrs("custname")
		Set custname2 = objrs("custname2")
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

%>

			<table width="1630" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" >
                      <tr class="trhd">
                        <td width="90" align="center">구분</td>
                        <td width="136" align="center">조정매체명</td>
                        <td width="88" align="center" >1월</td>
                        <td width="88" align="center">2월</td>
                        <td width="88" align="center">3월</td>
                        <td width="88" align="center">4월</td>
                        <td width="88" align="center">5월</td>
                        <td width="88" align="center">6월</td>
                        <td width="88" align="center">7월</td>
                        <td width="88" align="center">8월</td>
                        <td width="88" align="center">9월</td>
                        <td width="88" align="center">10월</td>
                        <td width="88" align="center">11월</td>
                        <td width="88" align="center">12월</td>
                        <td width="88" align="center">계</td>
                      </tr>
				<!--  -->
				<% do until objrs.eof
							If custname = "z" Then custname = "TOTAL"%>
				  <% If custname2 = "z" Then 	%>
                  <tr  class="trbd" bgcolor="#FFFFC1" >
                        <td width="90" align="left" bgcolor="#FFFFFF" >&nbsp;&nbsp;<%If prev <> custname Then response.write custname Else response.write "&nbsp;"%></td>
                        <td width="136" align="left">&nbsp;&nbsp;<%=custname%> 요약 </td>
				  <% Else %>
                  <tr  class="trbd" bgcolor="#FFFFFF" >
                        <td width="90" align="left" >&nbsp;&nbsp;<%If prev <> custname Then response.write custname Else response.write "&nbsp;"%></td>
                        <td width="136" align="left">&nbsp;&nbsp;<%=custname2%></td>
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
						prev = custname
						objrs.movenext
					loop
					objrs.close
					set objrs = nothing
				%>
              </table>