
<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<body oncontextmenu='return false' >
<%
	dim cyear, cyear2, cmonth, cmonth2, yearmon, yearmon2
	cyear = request("cyear")			' 시작년도
	if cyear = "" then cyear = year(date)			' 시작년도가 없으면 현재 년도를 기본 년도로 세팅
	cmonth = request("cmonth")	' 시작월
	if cmonth = "" then cmonth = month(date)' 시작월이 없으면 현재 월을 기본 월로 세팅
	If Len(cmonth) = 1 Then cmonth = "0"&cmonth			' 시작월이 1자리면 0을 붙여서 2자리 월로 변경
	cyear2 = request("cyear2")		' 종료년도
	if cyear2 = "" then cyear2 = year(date)		' 종료년도 기본 세팅
	cmonth2 = request("cmonth2")' 종료월
	if cmonth2 = "" then cmonth2 = month(date)			' 종료월 기본 세팅
	If Len(cmonth2) = 1 Then cmonth2 = "0"&cmonth2	' 종료 자리수 세팅

	yearmon = cyear & cmonth		' 시작년월 세팅
	yearmon2 = cyear2 & cmonth2	' 종료년월 세팅

	Dim custcode : custcode = request("tcustcode")			'광고주 코드
	Dim custcode2 : custcode2 = request("tcustcode2")		'사업부 코드'사업코드를 받지 못하면 null 로 세팅

	dim objrs, sql
	sql = "select highcustcode, custname from dbo.sc_cust_hdr where  MEDFLAG = 'A'  order by custname"
	call get_recordset(objrs, sql)

	dim str
	str = "<select name='tcustcode2' id='tcustcode2'>"
	do until objrs.eof
		str = str & "<option value='" & objrs("highcustcode") & "'"
			if custcode2 = objrs("highcustcode") then str = str & " selected"
		str = str & ">" & objrs("custname") & "</option>"
		objrs.movenext
	Loop
	str = str & "</select>"
	objrs.close

'	if custcode = custcode2 then 	custcode2 = null
'	if custcode2 = "" then custcode2 = Null

	if request.cookies("class") = "D" or request.cookies("class") = "H"  then
		custcode2 = request.cookies("custcode2")
	end if

	if not isnull(custcode2) then


'	sql = "select case when m.medflag = 'B' then '신문'		when m.medflag = 'C' then '잡지' end as medflag,c2.custname , isnull(sum(case when substring(m.yearmon, 5,2) = '01' then isnull(amt,0) else 0 end),0) as 'A01', isnull(sum(case when substring(m.yearmon, 5,2) = '02' then isnull(amt,0) else 0 end),0) as 'A02', isnull(sum(case when substring(m.yearmon, 5,2) = '03' then isnull(amt,0) else 0 end),0) as 'A03', isnull(sum(case when substring(m.yearmon, 5,2) = '04' then isnull(amt,0) else 0 end),0) as 'A04', isnull(sum(case when substring(m.yearmon, 5,2) = '05' then isnull(amt,0) else 0 end),0) as 'A05', isnull(sum(case when substring(m.yearmon, 5,2) = '06' then isnull(amt,0) else 0 end),0) as 'A06', isnull(sum(case when substring(m.yearmon, 5,2) = '07' then isnull(amt,0) else 0 end),0) as 'A07', isnull(sum(case when substring(m.yearmon, 5,2) = '08' then isnull(amt,0) else 0 end),0) as 'A08', isnull(sum(case when substring(m.yearmon, 5,2) = '09' then isnull(amt,0) else 0 end),0) as 'A09', isnull(sum(case when substring(m.yearmon, 5,2) = '10' then isnull(amt,0) else 0 end),0) as 'A10', isnull(sum(case when substring(m.yearmon, 5,2) = '11' then isnull(amt,0) else 0 end),0) as 'A11', isnull(sum(case when substring(m.yearmon, 5,2) = '12' then isnull(amt,0) else 0 end),0) as 'A12', sum(isnull(amt,0)) as 'TOTAL' from dbo.md_report_mst m  inner join dbo.sc_cust_temp c on m.clientcode = c.custcode inner join dbo.sc_cust_temp c2 on m.medcode = c2.custcode where m.medflag in ('b', 'C') and m.yearmon between '"&yearmon&"' and '"&yearmon2&"' and m.clientcode = '"&custcode2&"' group by case when m.medflag = 'B' then '신문' when m.medflag = 'C' then '잡지' end,c2.custname with rollup"


	sql = " select case when m.med_flag = 'B' then '신문'  when m.med_flag = 'C' then '잡지' end as medflag, c2.custname ,  isnull(sum(case when substring(m.yearmon, 5,2) = '01' then isnull(amt,0) else 0 end),0) as 'A01', isnull(sum(case when substring(m.yearmon, 5,2) = '02' then isnull(amt,0) else 0 end),0) as 'A02',  isnull(sum(case when substring(m.yearmon, 5,2) = '03' then isnull(amt,0) else 0 end),0) as 'A03',  isnull(sum(case when substring(m.yearmon, 5,2) = '04' then isnull(amt,0) else 0 end),0) as 'A04',  isnull(sum(case when substring(m.yearmon, 5,2) = '05' then isnull(amt,0) else 0 end),0) as 'A05',  isnull(sum(case when substring(m.yearmon, 5,2) = '06' then isnull(amt,0) else 0 end),0) as 'A06',  isnull(sum(case when substring(m.yearmon, 5,2) = '07' then isnull(amt,0) else 0 end),0) as 'A07',  isnull(sum(case when substring(m.yearmon, 5,2) = '08' then isnull(amt,0) else 0 end),0) as 'A08',  isnull(sum(case when substring(m.yearmon, 5,2) = '09' then isnull(amt,0) else 0 end),0) as 'A09',  isnull(sum(case when substring(m.yearmon, 5,2) = '10' then isnull(amt,0) else 0 end),0) as 'A10',  isnull(sum(case when substring(m.yearmon, 5,2) = '11' then isnull(amt,0) else 0 end),0) as 'A11',  isnull(sum(case when substring(m.yearmon, 5,2) = '12' then isnull(amt,0) else 0 end),0) as 'A12',  sum(isnull(amt,0)) as 'TOTAL' from dbo.md_report_mst_v m  inner join dbo.sc_cust_hdr c on m.clientcode = c.highcustcode  inner join dbo.sc_cust_dtl c2 on m.medcode = c2.custcode where m.med_flag in ('b', 'C')  and m.yearmon between '"&yearmon&"' and '"&yearmon2&"' and m.clientcode = '"&custcode2&"'  group by case when m.med_flag = 'B' then '신문'  when m.med_flag = 'C' then '잡지' end,c2.custname  with rollup "



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

%>
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
<% else

	sql = "select case when m.medflag = 'B' then '신문'		when m.medflag = 'C' then '잡지' end as medflag,c2.custname , isnull(sum(case when substring(m.yearmon, 5,2) = '01' then isnull(amt,0) else 0 end),0) as 'A01', isnull(sum(case when substring(m.yearmon, 5,2) = '02' then isnull(amt,0) else 0 end),0) as 'A02', isnull(sum(case when substring(m.yearmon, 5,2) = '03' then isnull(amt,0) else 0 end),0) as 'A03', isnull(sum(case when substring(m.yearmon, 5,2) = '04' then isnull(amt,0) else 0 end),0) as 'A04', isnull(sum(case when substring(m.yearmon, 5,2) = '05' then isnull(amt,0) else 0 end),0) as 'A05', isnull(sum(case when substring(m.yearmon, 5,2) = '06' then isnull(amt,0) else 0 end),0) as 'A06', isnull(sum(case when substring(m.yearmon, 5,2) = '07' then isnull(amt,0) else 0 end),0) as 'A07', isnull(sum(case when substring(m.yearmon, 5,2) = '08' then isnull(amt,0) else 0 end),0) as 'A08', isnull(sum(case when substring(m.yearmon, 5,2) = '09' then isnull(amt,0) else 0 end),0) as 'A09', isnull(sum(case when substring(m.yearmon, 5,2) = '10' then isnull(amt,0) else 0 end),0) as 'A10', isnull(sum(case when substring(m.yearmon, 5,2) = '11' then isnull(amt,0) else 0 end),0) as 'A11', isnull(sum(case when substring(m.yearmon, 5,2) = '12' then isnull(amt,0) else 0 end),0) as 'A12', sum(isnull(amt,0)) as 'TOTAL' from dbo.md_report_mst m  inner join dbo.sc_cust_temp c on m.clientcode = c.custcode inner join dbo.sc_cust_temp c2 on m.medcode = c2.custcode where m.medflag in ('b', 'C') and m.yearmon between '"&yearmon&"' and '"&yearmon2&"' and m.clientcode = '"&custcode&"' and m.clientsubcode = '" & custcode2 &"' group by case when m.medflag = 'B' then '신문' when m.medflag = 'C' then '잡지' end,c2.custname with rollup"

	call get_recordset(objrs, sql)

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

%>
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
<% end if%>
</body>
<SCRIPT LANGUAGE="JavaScript">
<!--
	var custcode2 = parent.document.getElementById("custcode2") ;
	custcode2.innerHTML = "<%=str%>";
//-->
</SCRIPT>
