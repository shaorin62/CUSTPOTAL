
<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<body oncontextmenu='return false' >
<%

	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1000

	dim cyear, cyear2, cmonth, cmonth2
	cyear = request("cyear")
	cmonth = request("cmonth")
	cyear2 = request("cyear2")
	cmonth2 = request("cmonth2")

	if cyear = "" then cyear = year(date)
	if cmonth = "" then cmonth = month(date)
	if cyear2 = "" then cyear2 = year(date)
	if cmonth2 = "" then cmonth2 = month(date)

	If Len(cmonth) = 1 Then cmonth = "0"&cmonth
	If Len(cmonth2) = 1 Then cmonth2 = "0"&cmonth2

	dim yearmon, yearmon2
	yearmon = cyear & cmonth
	yearmon2 = cyear2 & cmonth2

	Dim custcode : custcode = request("tcustcode")
	Dim custcode2 : custcode2 = request("tcustcode2")

	if custcode = "null" then custcode = null
	if custcode2 = "" then custcode2 = null
	if custcode = "" then custcode =null
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

	if isnull(custcode) then


	sql = "select isnull(c.yearmon, '총합계') as yearmon, c2.custname, isnull(sum(case when c.mpp = 'p00005' then isnull(amt,0) end),0) as 'P01'	, isnull(sum(case when c.mpp = 'p00007' then isnull(amt,0) end),0) as 'P02', isnull(sum(case when c.mpp = 'p00004' then isnull(amt,0) end),0) as 'P03', isnull(sum(case when isnull(c.mpp , '') not in('p00005','p00007','p00004') and c.medcode in('B00046', 'B00497')  then isnull(amt,0) end),0) as 'P04', isnull(sum(case when isnull(c.mpp , '') not in('p00005','p00007','p00004') and c.medcode in('B00047', 'B00498')  then isnull(amt,0) end),0) as 'P05' , isnull(sum(case when isnull(c.mpp , '') not in('p00005','p00007','p00004') and c.medcode not in ('B00046','B00047','B00497', 'B00498') then isnull(amt,0) end ),0) as 'OTH', sum(isnull(amt,0)) as 'TOTAL' from dbo.md_report_mst_v c inner join dbo.sc_cust_hdr c2 on c.clientcode = c2.highcustcode inner join dbo.sc_cust_hdr c3 on c.medcode = c3.highcustcode where  c.clientcode like '" & custcode2 & "%' and  yearmon between '"&yearmon&"' and '"&yearmon2&"' group by c.yearmon, c2.custname with rollup "


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
<% else
	sql = "select isnull(c.yearmon, '총합계') as yearmon, c2.custname, isnull(sum(case when c.mpp = 'p00005' then isnull(amt,0) end),0) as 'P01'	, isnull(sum(case when c.mpp = 'p00007' then isnull(amt,0) end),0) as 'P02', isnull(sum(case when c.mpp = 'p00004' then isnull(amt,0) end),0) as 'P03', isnull(sum(case when isnull(c.mpp , '') not in('p00005','p00007','p00004') and c.medcode in('B00046', 'B00497')  then isnull(amt,0) end),0) as 'P04', isnull(sum(case when isnull(c.mpp , '') not in('p00005','p00007','p00004') and c.medcode in('B00047', 'B00498') then isnull(amt,0) end),0) as 'P05' , isnull(sum(case when isnull(c.mpp , '') not in('p00005','p00007','p00004') and c.medcode not in ('B00046','B00047','B00497', 'B00498') then isnull(amt,0) end ),0) as 'OTH', sum(isnull(amt,0)) as 'TOTAL' from dbo.md_report_mst c inner join dbo.sc_cust_temp c2 on c.clientcode = c2.custcode inner join dbo.sc_cust_temp c3 on c.medcode = c3.custcode where c.clientcode = '" & custcode2 & "' and   yearmon between '"&yearmon&"' and '"&yearmon2&"' group by c.yearmon, c2.custname with rollup "

	call get_recordset(objrs, sql)

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

%>
	<table width="1020" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" >
	  <tr  class="trhd">
		<td  rowspan="2" align="center" >구분</td>
		<td colspan="3" align="center">케이블 TV</td>
		<td   align="center" >IPTV</td>
		<td  align="center"  >위성DMB</td>
		<td  rowspan="2" align="center" >Others</td>
		<td   rowspan="2" align="center" >총 집행 금액</td>
	  </tr>
	  <tr  class="trhd">
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
		<td width="150" align="center" > 총합계 </td>
		<td width="120" align="right" ><%If P01.value <> "0" Then response.write FormatNumber(P01,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
		<td width="120" align="right" ><%If P02.value <> "0" Then response.write FormatNumber(P02,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
		<td width="120" align="right" ><%If P03.value <> "0" Then response.write FormatNumber(P03,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
		<td width="120" align="right" ><%If P04.value <> "0" Then response.write FormatNumber(P04,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
		<td width="120" align="right" ><%If P05.value <> "0" Then response.write FormatNumber(P05,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
		<td width="120" align="right" ><%If OTH.value <> "0" Then response.write FormatNumber(OTH,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
		<td width="120" align="right" ><%If total.value <> "0" Then response.write FormatNumber(total,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
      </tr>
			<% ElseIf cyearmon <> "총합계" And IsNull(custname) Then %>
			<%Else %>
      <tr  class="trbd" bgcolor="#FFFFFF" >
		<td  align="center"><%If prev <> cyearmon Then response.write left(cyearmon,4) & chr(45) & right(cyearmon,2) Else response.write "&nbsp;"%></td>
		<td 			align="right" ><%If P01.value <> "0" Then response.write FormatNumber(P01,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
		<td 	align="right" ><%If P02.value <> "0" Then response.write FormatNumber(P02,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
		<td 	align="right" ><%If P03.value <> "0" Then response.write FormatNumber(P03,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
		<td 	align="right" ><%If P04.value <> "0" Then response.write FormatNumber(P04,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
		<td 	align="right" ><%If P05.value <> "0" Then response.write FormatNumber(P05,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
		<td 	align="right" ><%If OTH.value <> "0" Then response.write FormatNumber(OTH,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
		<td 	align="right" ><%If total.value <> "0" Then response.write FormatNumber(total,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
      <tr  class="trbd" bgcolor="#FFFFFF" >
		<td width="150" align="center">%</td>
		<td 	align="right" ><%If P01.value <> "0" and total.value <> "0"  Then response.write replace(formatpercent(cdbl(P01)/cdbl(total)),"%", "") Else  response.write "-"%>&nbsp;&nbsp;</td>
		<td 	align="right" ><%If P02.value <> "0" Then response.write replace(formatpercent(cdbl(P02)/cdbl(total)),"%", "") Else  response.write "-"%>&nbsp;&nbsp;</td>
		<td 	align="right" ><%If P03.value <> "0" Then response.write replace(formatpercent(cdbl(P03)/cdbl(total)),"%", "") Else  response.write "-"%>&nbsp;&nbsp;</td>
		<td 	align="right" ><%If P04.value <> "0" Then response.write replace(formatpercent(cdbl(P04)/cdbl(total)),"%", "") Else  response.write "-"%>&nbsp;&nbsp;</td>
		<td 	align="right" ><%If P05.value <> "0" Then response.write replace(formatpercent(cdbl(P05)/cdbl(total)),"%", "") Else  response.write "-"%>&nbsp;&nbsp;</td>
		<td 	align="right" ><%If OTH.value <> "0" Then response.write replace(formatpercent(cdbl(OTH)/cdbl(total)),"%", "") Else  response.write "-"%>&nbsp;&nbsp;</td>
		<td 	align="right" ><%If total.value <> "0" Then response.write replace(formatpercent(cdbl(total)/cdbl(total)),"%", "") Else  response.write "-"%>&nbsp;&nbsp;</td>
      <tr  class="trbd" bgcolor="#FFFFC1" >
		<td width="150" align="center">On Media vs. CJ Media</td>
		<td 	align="right" >-</td>
			<td 	align="right" ><%If P02.value <> "0" and  P03.value <> "0"  Then response.write replace(FormatPercent(CDBL(P02)/(CDBL(P02)+CDBL(P03)),0),"%","") Else  response.write "-"%>&nbsp;&nbsp;</td>
			<td 	align="right" ><%If P02.value <> "0" and  P03.value <> "0"  Then response.write replace(FormatPercent(CDBL(P03)/(CDBL(P02)+CDBL(P03)),0),"%","") Else  response.write "-"%>&nbsp;&nbsp;</td>
		<td 	align="right" >-</td>
		<td 	align="right" >-</td>
		<td 	align="right" >-</td>
		<td 	align="right" >-</td>
      </tr>
			<% End If %>


<%
		prev = cyearmon
		objrs.movenext
					loop
					objrs.close
					set objrs = nothing
				%>
	</table>
	<% end if %>
</body>
<SCRIPT LANGUAGE="JavaScript">
<!--
	var custcode2 = parent.document.getElementById("custcode2") ;
	custcode2.innerHTML = "<%=str%>";
//-->
</SCRIPT>