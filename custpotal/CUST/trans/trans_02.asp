
<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim cyear, cyear2, cmonth, cmonth2, yearmon, yearmon2
	cyear = request("cyear")											' 시작년도
	if cyear = "" then cyear = year(date)							' 시작년도가 없으면 현재 년도를 기본 년도로 세팅
	cmonth = request("cmonth")									' 시작월
	if cmonth = "" then cmonth = month(date)				' 시작월이 없으면 현재 월을 기본 월로 세팅
	If Len(cmonth) = 1 Then cmonth = "0"&cmonth			' 시작월이 1자리면 0을 붙여서 2자리 월로 변경
	cyear2 = request("cyear2")										' 종료년도
	if cyear2 = "" then cyear2 = year(date)						' 종료년도 기본 세팅
	cmonth2 = request("cmonth2")								' 종료월
	if cmonth2 = "" then cmonth2 = month(date)			' 종료월 기본 세팅
	If Len(cmonth2) = 1 Then cmonth2 = "0"&cmonth2	' 종료 자리수 세팅

	yearmon = cyear & cmonth										' 시작년월 세팅
	yearmon2 = cyear2 & cmonth2									' 종료년월 세팅

	Dim custcode : custcode = request("tcustcode")			'사업부 코드				'사업코드를 받지 못하면 null 로 세팅
	Dim custcode2 : custcode2 = request("tcustcode2")		'광고주 코드

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

'	if custcode = custcode2 then 	custcode = null
'	if custcode2 = "" then custcode2 = Null

	if request.cookies("class") = "D" or request.cookies("class") = "H"  then
		custcode2 = request.cookies("custcode2")
	end if

	if not isnull(custcode2)  then

	sql = "select grouping(f.category) as g_medflag,grouping(c2.custname) as g_custname2,grouping(seqname) as g_seqname, grouping(e.mattername) as g_progname,f.category as medflag, c2.custname as custname2 ,seqname, e.mattername ,sum(case when substring(yearmon,5,2) = '01' then isnull(amt,0) else 0 end ) as 'A01',sum(case when  substring(yearmon,5,2) = '02' then isnull(amt,0) else 0 end ) as 'A02',sum(case when substring(yearmon,5,2) = '03' then isnull(amt,0) else 0 end ) as 'A03',sum(case when substring(yearmon,5,2) = '04' then isnull(amt,0) else 0 end ) as 'A04',sum(case when substring(yearmon,5,2) = '05' then isnull(amt,0) else 0 end ) as 'A05',sum(case when substring(yearmon,5,2) = '06' then isnull(amt,0) else 0 end ) as 'A06',sum(case when substring(yearmon,5,2) = '07' then isnull(amt,0) else 0 end ) as 'A07',sum(case when substring(yearmon,5,2) = '08' then isnull(amt,0) else 0 end ) as 'A08',sum(case when substring(yearmon,5,2) = '09' then isnull(amt,0) else 0 end ) as 'A09',sum(case when substring(yearmon,5,2) = '10' then isnull(amt,0) else 0 end ) as 'A10',sum(case when substring(yearmon,5,2) = '11' then isnull(amt,0) else 0 end ) as 'A11',sum(case when substring(yearmon,5,2) = '12' then isnull(amt,0) else 0 end ) as 'A12' ,sum(isnull(amt,0)) as 'TOTAL' from dbo.md_report_mst_v e inner join dbo.vw_medflag f on e.med_flag= f.med_flag inner join dbo.sc_cust_hdr c on e.clientcode = c.highcustcode left outer join dbo.sc_subseq_dtl j on j.seqno = e.subseq inner join dbo.sc_cust_dtl c2 on c2.custcode = e.timcode where e.clientcode = '"&custcode2&"' and  (e.yearmon between '"&yearmon&"' and '"&yearmon2&"') group by f.category , c2.custname, j.seqname, e.mattername with rollup "



	call get_recordset(objrs, sql)

	Dim g_medflag, g_custname2, g_seqname, g_progname, medflag, custname2, seqname, progname, A01, A02, A03, A04, A05, A06, A07, A08, A09, A10, A11, A12, total, prev_medflag, prev_custname2, prev_seqname
	If Not objrs.eof Then
		Set g_medflag = objrs("g_medflag")
		Set g_custname2 = objrs("g_custname2")
		Set g_seqname = objrs("g_seqname")
		Set g_progname = objrs("g_progname")
		Set medflag = objrs("medflag")
		Set custname2 = objrs("custname2")
		Set seqname = objrs("seqname")
		Set progname = objrs("mattername")
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
				  <table width="1660" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" >
                      <tr class="trhd">
                        <td width="90" align="center">매체</td>
                        <td width="140" align="center">사업부</td>
                        <td width="110" align="center">브랜드</td>
                        <td width="150" align="center">소재명</td>
                        <td width="100" align="center" >1월</td>
                        <td width="100" align="center">2월</td>
                        <td width="100" align="center">3월</td>
                        <td width="100" align="center">4월</td>
                        <td width="100" align="center">5월</td>
                        <td width="100" align="center">6월</td>
                        <td width="100" align="center">7월</td>
                        <td width="100" align="center">8월</td>
                        <td width="100" align="center">9월</td>
                        <td width="100" align="center">10월</td>
                        <td width="100" align="center">11월</td>
                        <td width="100" align="center">12월</td>
                        <td width="100" align="center">계</td>
                      </tr>
				<!--  -->
				<% do until objrs.eof 	%>
						  <% If   g_custname2 = 0 And g_seqname = 0 And g_progname = 0 Then %>
						  <tr  class="trbd" bgcolor="#FFFFFF" >
								<td width="90" align="left"  bgcolor="#FFFFFF">&nbsp;&nbsp;<%If prev_medflag <> medflag Then response.write medflag Else response.write "&nbsp;"%> </td>
								<td width="140" align="left" bgcolor="#FFFFFF">&nbsp;&nbsp;<%If prev_custname2 <> custname2 Then response.write custname2 Else response.write "&nbsp;"%></td>
								<td width="110" align="left"  bgcolor="#FFFFFF">&nbsp;&nbsp;<%=seqname%> </td>
								<td width="150" align="left"  bgcolor="#FFFFFF">&nbsp;&nbsp;<%=progname%> </td>
						  <% ElseIf  g_custname2 = 0 And g_seqname = 0 And g_progname = 1 Then %>
						  <tr  class="trbd" bgcolor="#FFDFDF" >
								<td width="90" align="left"  bgcolor="#FFFFFF">&nbsp;&nbsp;<%If prev_medflag <> medflag Then response.write medflag Else response.write "&nbsp;"%> </td>
								<td width="140" align="left"  bgcolor="#FFFFFF">&nbsp;&nbsp;<%If prev_custname2 <> custname2 Then response.write custname2 Else response.write "&nbsp;"%></td>
								<td width="260" align="left" colspan="2">&nbsp;&nbsp;<%=seqname%> 요약</td>
							<% ElseIf g_custname2 = 0 And g_seqname = 1 And g_progname =1 Then %>
						  <tr  class="trbd" bgcolor="#CCFFFF" >
								<td width="90" align="left"  bgcolor="#FFFFFF">&nbsp;&nbsp;<%If prev_medflag <> medflag Then response.write medflag Else response.write "&nbsp;"%> </td>
								<td width="400" align="left" colspan="3" bgcolor="">&nbsp;&nbsp;<%=custname2%> 요약</td>
							<% ElseIf g_medflag =0 and g_custname2 = 1 And g_seqname = 1 And g_progname =1 then %>
						  <tr  class="trbd" bgcolor="#FFFFC1" >
								<td width="480" align="left" colspan="4">&nbsp;&nbsp;<%=medflag%> 요약 </td>
							<% ElseIf g_medflag =1 and g_custname2 = 1 And g_seqname = 1 And g_progname =1 then %>
						  <tr  class="trbd" bgcolor="#FFC1C1" >
								<td width="480" align="left" colspan="4">&nbsp;&nbsp;총합계 </td>
							<%End if%>
                        <td width="100" align="right" ><%If A01.value <> "0" Then response.write FormatNumber(A01,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A02.value <> "0" Then response.write FormatNumber(A02,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A03.value <> "0" Then response.write FormatNumber(A03,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A04.value <> "0" Then response.write FormatNumber(A04,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A05.value <> "0" Then response.write FormatNumber(A05,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A06.value <> "0" Then response.write FormatNumber(A06,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A07.value <> "0" Then response.write FormatNumber(A07,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A08.value <> "0" Then response.write FormatNumber(A08,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A09.value <> "0" Then response.write FormatNumber(A09,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A10.value <> "0" Then response.write FormatNumber(A10,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A11.value <> "0" Then response.write FormatNumber(A11,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A12.value <> "0" Then response.write FormatNumber(A12,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If total.value <> "0" Then response.write FormatNumber(total,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                      </tr>
				<%
						'End if
						If g_custname2 =1 And g_seqname =1 And g_progname =1 Then
							prev_custname2 = ""
							prev_medflag = ""
						else
							prev_custname2 = custname2
							prev_medflag = medflag
						End If
						objrs.movenext
					loop
					objrs.close
					set objrs = nothing
				%>
              </table>
<% else

	sql = "select grouping(f.category) as g_medflag, grouping(seqname) as g_seqname, grouping(progname) as g_progname, f.category as medflag, seqname, progname ,sum(case when substring(yearmon,5,2) = '01' then isnull(amt,0) else 0 end ) as 'A01',sum(case when substring(yearmon,5,2) = '02' then isnull(amt,0) else 0 end ) as 'A02',sum(case when substring(yearmon,5,2) = '03' then isnull(amt,0) else 0 end ) as 'A03',sum(case when substring(yearmon,5,2) = '04' then isnull(amt,0) else 0 end ) as 'A04',sum(case when substring(yearmon,5,2) = '05' then isnull(amt,0) else 0 end ) as 'A05',sum(case when substring(yearmon,5,2) = '06' then isnull(amt,0) else 0 end ) as 'A06',sum(case when substring(yearmon,5,2) = '07' then isnull(amt,0) else 0 end ) as 'A07',sum(case when substring(yearmon,5,2) = '08' then isnull(amt,0) else 0 end ) as 'A08',sum(case when substring(yearmon,5,2) = '09' then isnull(amt,0) else 0 end ) as 'A09',sum(case when substring(yearmon,5,2) = '10' then isnull(amt,0) else 0 end ) as 'A10',sum(case when substring(yearmon,5,2) = '11' then isnull(amt,0) else 0 end ) as 'A11',sum(case when substring(yearmon,5,2) = '12' then isnull(amt,0) else 0 end ) as 'A12' , sum(isnull(amt,0)) as 'TOTAL' from dbo.md_report_mst e inner join dbo.vw_medflag f on e.medflag= f.medflag inner join dbo.sc_cust_temp c on e.clientcode = c.custcode left outer join dbo.sc_jobcust j on j.seqno = e.subseq inner join dbo.sc_cust_temp c2 on c2.custcode = e.clientsubcode where e.clientsubcode = '"&custcode2&"' and (e.yearmon between '"&yearmon&"' and '"&yearmon2&"') group by f.category,j.seqname, progname with rollup "


	call get_recordset(objrs, sql)

	If Not objrs.eof Then
		Set g_medflag = objrs("g_medflag")
		Set g_seqname = objrs("g_seqname")
		Set g_progname = objrs("g_progname")
		Set medflag = objrs("med_flag")
'		Set custname2 = objrs("custname2")
		Set seqname = objrs("seqname")
		Set progname = objrs("progname")
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
				  <table width="1520" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" >
                      <tr class="trhd">
                        <td width="90" align="center">매체</td>
                        <td width="110" align="center">브랜드</td>
                        <td width="150" align="center">소재명</td>
                        <td width="100" align="center" >1월</td>
                        <td width="100" align="center">2월</td>
                        <td width="100" align="center">3월</td>
                        <td width="100" align="center">4월</td>
                        <td width="100" align="center">5월</td>
                        <td width="100" align="center">6월</td>
                        <td width="100" align="center">7월</td>
                        <td width="100" align="center">8월</td>
                        <td width="100" align="center">9월</td>
                        <td width="100" align="center">10월</td>
                        <td width="100" align="center">11월</td>
                        <td width="100" align="center">12월</td>
                        <td width="100" align="center">계</td>
                      </tr>
				<!--  -->
				<% do until objrs.eof 	%>
				  <% If   g_medflag = 0 and g_seqname = 0 And g_progname = 0 Then %>
				  <tr  class="trbd" bgcolor="#FFFFFF" >
					<td width="90" align="left"  bgcolor="#FFFFFF">&nbsp;&nbsp;<%If prev_medflag <> medflag Then response.write medflag Else response.write "&nbsp;"%> </td>
					<td width="110" align="left"  bgcolor="#FFFFFF">&nbsp;&nbsp;<%=seqname%> </td>
					<td width="150" align="left"  bgcolor="#FFFFFF">&nbsp;&nbsp;<%=progname%> </td>
				  <% ElseIf  g_medflag = 0 and g_seqname = 0 And g_progname = 1 Then %>
				  <tr  class="trbd" bgcolor="#FFDFDF" >
					<td width="90" align="left"  bgcolor="#FFFFFF">&nbsp;&nbsp;<%If prev_medflag <> medflag Then response.write medflag Else response.write "&nbsp;"%> </td>
					<td width="260" align="left"  colspan="2"> <%=seqname%> 요약</td>
				<% ElseIf g_medflag = 1 and g_seqname = 1 And g_progname =1 then %>
				  <tr  class="trbd" bgcolor="#FFC1C1" >
					<td width="480" align="center" colspan="3"> 총합계</td>
				<% ElseIf g_medflag = 0 and g_seqname = 1 And g_progname =1 then %>
				  <tr  class="trbd" bgcolor="#FFFFC1" >
					<td width="480" align="left" colspan="3"> <%=medflag%> 요약 </td>
				<%End if%>
                        <td width="100" align="right" ><%If A01.value <> "0" Then response.write FormatNumber(A01,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A02.value <> "0" Then response.write FormatNumber(A02,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A03.value <> "0" Then response.write FormatNumber(A03,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A04.value <> "0" Then response.write FormatNumber(A04,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A05.value <> "0" Then response.write FormatNumber(A05,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A06.value <> "0" Then response.write FormatNumber(A06,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A07.value <> "0" Then response.write FormatNumber(A07,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A08.value <> "0" Then response.write FormatNumber(A08,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A09.value <> "0" Then response.write FormatNumber(A09,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A10.value <> "0" Then response.write FormatNumber(A10,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A11.value <> "0" Then response.write FormatNumber(A11,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If A12.value <> "0" Then response.write FormatNumber(A12,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                        <td width="100" align="right" ><%If total.value <> "0" Then response.write FormatNumber(total,0) Else  response.write "-"%>&nbsp;&nbsp;</td>
                      </tr>
				<%
						If g_seqname =1 And g_progname =1 Then
							prev_custname2 = ""
							prev_medflag = ""
						else
							prev_custname2 = custname2
							prev_medflag = medflag
						End If
						objrs.movenext
					loop
					objrs.close
					set objrs = nothing
				%>
              </table><!-- 부서코드로 표현하는 경우 -->

<% end if%>
</body>
<SCRIPT LANGUAGE="JavaScript">
<!--
	var custcode2 = parent.document.getElementById("custcode2") ;
	custcode2.innerHTML = "<%=str%>";
//-->
</SCRIPT>
