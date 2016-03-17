
<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%

	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1000

	dim cyear, cyear2, cmonth, cmonth2
	cyear = request("cyear")
	if cyear = "" then cyear = year(date)
	cmonth = request("cmonth")
	if cmonth = "" then cmonth = month(date)
	If Len(cmonth) = 1 Then cmonth = "0"&cmonth
	cyear2 = request("cyear2")
	if cyear2 = "" then cyear2 = year(date)
	cmonth2 = request("cmonth2")
	if cmonth2 = "" then cmonth2 = month(date)
	If Len(cmonth2) = 1 Then cmonth2 = "0"&cmonth2
	dim yearmon, yearmon2
	yearmon = cyear & cmonth
	yearmon2 = cyear & cmonth2

	Dim custcode : custcode = request("tcustcode")
	Dim custcode2 : custcode2 = request("tcustcode2")

	dim objrs, sql

	sql = "select highcustcode, custname from dbo.sc_cust_hdr where medflag = 'A' order by custname"
	call get_recordset(objrs, sql)

	dim str
	str = "<select name='tcustcode2' id='tcustcode2'>"
	str = str & "<option value='' selected> -- 전체 광고주 -- </option>"
	do until objrs.eof
		str = str & "<option value='" & objrs("highcustcode") & "'"
			if custcode2 <> "" then
				if custcode2 = objrs("highcustcode") then str = str & " selected"
			end if
		str = str & ">" & objrs("custname") & "</option>"
		objrs.movenext
	Loop
	str = str & "</select>"
	objrs.close

	if custcode2 = "" or custcode = custcode2 then custcode2 = null


	sql = "select case when e.med_flag = '01'	then 'TV' when e.med_flag in ('02', '03')	then 'Radio' when e.med_flag = 'A2' then 'CATV'	when e.med_flag in ('10', '20')  then '지상파 DMB' when e.med_flag = 'B' 	then '신문' when e.med_flag = 'C'	then '잡지' when e.med_flag = 'O'	then '인터넷' when e.med_flag = 'D' then '옥외' else '총합계' end as med_flag ,c2.custname as custname2, sum(case when substring(e.yearmon,5,2) = '01' then isnull(amt,0) else 0 end ) as 'A01',sum(case when substring(e.yearmon,5,2) = '02' then isnull(amt,0) else 0 end ) as 'A02',sum(case when substring(e.yearmon,5,2) = '03' then isnull(amt,0) else 0 end ) as 'A03',sum(case when substring(e.yearmon,5,2) = '04' then isnull(amt,0) else 0 end ) as 'A04',sum(case when substring(e.yearmon,5,2) = '05' then isnull(amt,0) else 0 end ) as 'A05',sum(case when substring(e.yearmon,5,2) = '06' then isnull(amt,0) else 0 end ) as 'A06',sum(case when substring(e.yearmon,5,2) = '07' then isnull(amt,0) else 0 end ) as 'A07',sum(case when substring(e.yearmon,5,2) = '08' then isnull(amt,0) else 0 end ) as 'A08',sum(case when substring(e.yearmon,5,2) = '09' then isnull(amt,0) else 0 end ) as 'A09',sum(case when substring(e.yearmon,5,2) = '10' then isnull(amt,0) else 0 end ) as 'A10',sum(case when substring(e.yearmon,5,2) = '11' then isnull(amt,0) else 0 end ) as 'A11',sum(case when substring(e.yearmon,5,2) = '12' then isnull(amt,0) else 0 end ) as 'A12', sum(isnull(amt,0)) as 'TOTAL'from dbo.md_report_mst_v e inner join dbo.sc_cust_hdr c on e.clientcode = c.highcustcode inner join dbo.sc_cust_hdr c2 on e.medcode = c2.highcustcode where yearmon between '"&yearmon&"' and '"&yearmon2&"' and e.clientcode like '"&custcode2&"%' group by  case when e.med_flag = '01' then 'TV' when e.med_flag in ('02', '03') then 'Radio' when e.med_flag = 'A2' then 'CATV' when e.med_flag in ('10', '20') then '지상파 DMB' when e.med_flag = 'B' then '신문' when e.med_flag = 'C' then '잡지' when e.med_flag = 'O' then '인터넷' when e.med_flag = 'D' then '옥외' else '총합계' end , c2.custname with rollup"


	call get_recordset(objrs, sql)

	Dim medflag, custname2, A01, A02, A03, A04, A05, A06, A07, A08, A09, A10, A11, A12, total, prev
	If Not objrs.eof Then
		Set medflag = objrs("med_flag")
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

<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<body oncontextmenu='return false' >
				  <table width="1400" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" >
                      <tr class="trhd">
                        <td width="90" align="center">매체</td>
                        <td width="140" align="center">매체명</td>
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
				<% if not isnull(medflag) and isnull(custname2) then %>
                  <tr  class="trbd" bgcolor="#FFFFC1" >
                        <td width="240" align="center"  colspan="2"><%=medflag%> 소계</td>
				<% elseif isnull(medflag) and isnull(custname2) then %>
                  <tr  class="trbd" bgcolor="#FFC1C1" >
                        <td width="240" align="center"  colspan="2">총합계</td>
				<% else %>
                  <tr  class="trbd" bgcolor="#FFFFFF" >
                        <td width="90" align="center"  ><%If prev <> medflag Then response.write medflag Else response.write "&nbsp;"%></td>
                        <td width="150" align="center"  ><%=custname2%></td>
				<%end if %>
                        <td width="90" align="right" ><%If A01.value <> "0" Then response.write FormatNumber(A01,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="90" align="right" ><%If A02.value <> "0" Then response.write FormatNumber(A02,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="90" align="right" ><%If A03.value <> "0" Then response.write FormatNumber(A03,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="90" align="right" ><%If A04.value <> "0" Then response.write FormatNumber(A04,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="90" align="right" ><%If A05.value <> "0" Then response.write FormatNumber(A05,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="90" align="right" ><%If A06.value <> "0" Then response.write FormatNumber(A06,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="90" align="right" ><%If A07.value <> "0" Then response.write FormatNumber(A07,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="90" align="right" ><%If A08.value <> "0" Then response.write FormatNumber(A08,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="90" align="right" ><%If A09.value <> "0" Then response.write FormatNumber(A09,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="90" align="right" ><%If A10.value <> "0" Then response.write FormatNumber(A10,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="90" align="right" ><%If A11.value <> "0" Then response.write FormatNumber(A11,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="90" align="right" ><%If A12.value <> "0" Then response.write FormatNumber(A12,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="90" align="right" ><%If total.value <> "0" Then response.write FormatNumber(total,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                      </tr>
				<%
						'End if
						prev = medflag
						objrs.movenext
					loop
					objrs.close
					set objrs = nothing
				%>
              </table>
</body>
<SCRIPT LANGUAGE="JavaScript">
<!--
	var custcode2 = parent.document.getElementById("custcode2") ;
	custcode2.innerHTML = "<%=str%>";
//-->
</SCRIPT>