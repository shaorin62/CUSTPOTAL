
<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<body oncontextmenu='return false' >
<%
	dim cyear, cyear2, cmonth, cmonth2, yearmon, yearmon2
	cyear = request("cyear")			' ���۳⵵
	if cyear = "" then cyear = year(date)			' ���۳⵵�� ������ ���� �⵵�� �⺻ �⵵�� ����
	cmonth = request("cmonth")	' ���ۿ�
	if cmonth = "" then cmonth = month(date)' ���ۿ��� ������ ���� ���� �⺻ ���� ����
	If Len(cmonth) = 1 Then cmonth = "0"&cmonth			' ���ۿ��� 1�ڸ��� 0�� �ٿ��� 2�ڸ� ���� ����
	cyear2 = request("cyear2")		' ����⵵
	if cyear2 = "" then cyear2 = year(date)		' ����⵵ �⺻ ����
	cmonth2 = request("cmonth2")' �����
	if cmonth2 = "" then cmonth2 = month(date)			' ����� �⺻ ����
	If Len(cmonth2) = 1 Then cmonth2 = "0"&cmonth2	' ���� �ڸ��� ����

	yearmon = cyear & cmonth		' ���۳�� ����
	yearmon2 = cyear2 & cmonth2	' ������ ����

	Dim custcode : custcode = request("tcustcode")			'������ �ڵ�
	Dim custcode2 : custcode2 = request("tcustcode2")		'����� �ڵ�'����ڵ带 ���� ���ϸ� null �� ����

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


'	sql = "select case when m.medflag = 'B' then '�Ź�'		when m.medflag = 'C' then '����' end as medflag,c2.custname , isnull(sum(case when substring(m.yearmon, 5,2) = '01' then isnull(amt,0) else 0 end),0) as 'A01', isnull(sum(case when substring(m.yearmon, 5,2) = '02' then isnull(amt,0) else 0 end),0) as 'A02', isnull(sum(case when substring(m.yearmon, 5,2) = '03' then isnull(amt,0) else 0 end),0) as 'A03', isnull(sum(case when substring(m.yearmon, 5,2) = '04' then isnull(amt,0) else 0 end),0) as 'A04', isnull(sum(case when substring(m.yearmon, 5,2) = '05' then isnull(amt,0) else 0 end),0) as 'A05', isnull(sum(case when substring(m.yearmon, 5,2) = '06' then isnull(amt,0) else 0 end),0) as 'A06', isnull(sum(case when substring(m.yearmon, 5,2) = '07' then isnull(amt,0) else 0 end),0) as 'A07', isnull(sum(case when substring(m.yearmon, 5,2) = '08' then isnull(amt,0) else 0 end),0) as 'A08', isnull(sum(case when substring(m.yearmon, 5,2) = '09' then isnull(amt,0) else 0 end),0) as 'A09', isnull(sum(case when substring(m.yearmon, 5,2) = '10' then isnull(amt,0) else 0 end),0) as 'A10', isnull(sum(case when substring(m.yearmon, 5,2) = '11' then isnull(amt,0) else 0 end),0) as 'A11', isnull(sum(case when substring(m.yearmon, 5,2) = '12' then isnull(amt,0) else 0 end),0) as 'A12', sum(isnull(amt,0)) as 'TOTAL' from dbo.md_report_mst m  inner join dbo.sc_cust_temp c on m.clientcode = c.custcode inner join dbo.sc_cust_temp c2 on m.medcode = c2.custcode where m.medflag in ('b', 'C') and m.yearmon between '"&yearmon&"' and '"&yearmon2&"' and m.clientcode = '"&custcode2&"' group by case when m.medflag = 'B' then '�Ź�' when m.medflag = 'C' then '����' end,c2.custname with rollup"


	sql = " select case when m.med_flag = 'B' then '�Ź�'  when m.med_flag = 'C' then '����' end as medflag, c2.custname ,  isnull(sum(case when substring(m.yearmon, 5,2) = '01' then isnull(amt,0) else 0 end),0) as 'A01', isnull(sum(case when substring(m.yearmon, 5,2) = '02' then isnull(amt,0) else 0 end),0) as 'A02',  isnull(sum(case when substring(m.yearmon, 5,2) = '03' then isnull(amt,0) else 0 end),0) as 'A03',  isnull(sum(case when substring(m.yearmon, 5,2) = '04' then isnull(amt,0) else 0 end),0) as 'A04',  isnull(sum(case when substring(m.yearmon, 5,2) = '05' then isnull(amt,0) else 0 end),0) as 'A05',  isnull(sum(case when substring(m.yearmon, 5,2) = '06' then isnull(amt,0) else 0 end),0) as 'A06',  isnull(sum(case when substring(m.yearmon, 5,2) = '07' then isnull(amt,0) else 0 end),0) as 'A07',  isnull(sum(case when substring(m.yearmon, 5,2) = '08' then isnull(amt,0) else 0 end),0) as 'A08',  isnull(sum(case when substring(m.yearmon, 5,2) = '09' then isnull(amt,0) else 0 end),0) as 'A09',  isnull(sum(case when substring(m.yearmon, 5,2) = '10' then isnull(amt,0) else 0 end),0) as 'A10',  isnull(sum(case when substring(m.yearmon, 5,2) = '11' then isnull(amt,0) else 0 end),0) as 'A11',  isnull(sum(case when substring(m.yearmon, 5,2) = '12' then isnull(amt,0) else 0 end),0) as 'A12',  sum(isnull(amt,0)) as 'TOTAL' from dbo.md_report_mst_v m  inner join dbo.sc_cust_hdr c on m.clientcode = c.highcustcode  inner join dbo.sc_cust_dtl c2 on m.medcode = c2.custcode where m.med_flag in ('b', 'C')  and m.yearmon between '"&yearmon&"' and '"&yearmon2&"' and m.clientcode = '"&custcode2&"'  group by case when m.med_flag = 'B' then '�Ź�'  when m.med_flag = 'C' then '����' end,c2.custname  with rollup "



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
                        <td width="90" align="center">����</td>
                        <td width="150" align="center">��ü</td>
                        <td width="90" align="center" >1��</td>
                        <td width="90" align="center">2��</td>
                        <td width="90" align="center">3��</td>
                        <td width="90" align="center">4��</td>
                        <td width="90" align="center">5��</td>
                        <td width="90" align="center">6��</td>
                        <td width="90" align="center">7��</td>
                        <td width="90" align="center">8��</td>
                        <td width="90" align="center">9��</td>
                        <td width="90" align="center">10��</td>
                        <td width="90" align="center">11��</td>
                        <td width="90" align="center">12��</td>
                        <td width="90" align="center">��</td>
                      </tr>
				<!--  -->
				<% do until objrs.eof 	%>
				<% If IsNull(medflag) And IsNull(custname) Then %>
                  <tr  class="trbd" bgcolor="#FFFFC1" >
                        <td width="220" align="center" colspan="2">&nbsp;&nbsp;�հ�</td>
				  <% ElseIf Not IsNull(medflag) And IsNull(custname) then%>
                  <tr  class="trbd" bgcolor="#CCFFFF" >
                        <td width="220" align="center" colspan="2">&nbsp;&nbsp;<%=medflag%> �Ұ�</td>
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

	sql = "select case when m.medflag = 'B' then '�Ź�'		when m.medflag = 'C' then '����' end as medflag,c2.custname , isnull(sum(case when substring(m.yearmon, 5,2) = '01' then isnull(amt,0) else 0 end),0) as 'A01', isnull(sum(case when substring(m.yearmon, 5,2) = '02' then isnull(amt,0) else 0 end),0) as 'A02', isnull(sum(case when substring(m.yearmon, 5,2) = '03' then isnull(amt,0) else 0 end),0) as 'A03', isnull(sum(case when substring(m.yearmon, 5,2) = '04' then isnull(amt,0) else 0 end),0) as 'A04', isnull(sum(case when substring(m.yearmon, 5,2) = '05' then isnull(amt,0) else 0 end),0) as 'A05', isnull(sum(case when substring(m.yearmon, 5,2) = '06' then isnull(amt,0) else 0 end),0) as 'A06', isnull(sum(case when substring(m.yearmon, 5,2) = '07' then isnull(amt,0) else 0 end),0) as 'A07', isnull(sum(case when substring(m.yearmon, 5,2) = '08' then isnull(amt,0) else 0 end),0) as 'A08', isnull(sum(case when substring(m.yearmon, 5,2) = '09' then isnull(amt,0) else 0 end),0) as 'A09', isnull(sum(case when substring(m.yearmon, 5,2) = '10' then isnull(amt,0) else 0 end),0) as 'A10', isnull(sum(case when substring(m.yearmon, 5,2) = '11' then isnull(amt,0) else 0 end),0) as 'A11', isnull(sum(case when substring(m.yearmon, 5,2) = '12' then isnull(amt,0) else 0 end),0) as 'A12', sum(isnull(amt,0)) as 'TOTAL' from dbo.md_report_mst m  inner join dbo.sc_cust_temp c on m.clientcode = c.custcode inner join dbo.sc_cust_temp c2 on m.medcode = c2.custcode where m.medflag in ('b', 'C') and m.yearmon between '"&yearmon&"' and '"&yearmon2&"' and m.clientcode = '"&custcode&"' and m.clientsubcode = '" & custcode2 &"' group by case when m.medflag = 'B' then '�Ź�' when m.medflag = 'C' then '����' end,c2.custname with rollup"

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
                        <td width="90" align="center">����</td>
                        <td width="150" align="center">��ü</td>
                        <td width="90" align="center" >1��</td>
                        <td width="90" align="center">2��</td>
                        <td width="90" align="center">3��</td>
                        <td width="90" align="center">4��</td>
                        <td width="90" align="center">5��</td>
                        <td width="90" align="center">6��</td>
                        <td width="90" align="center">7��</td>
                        <td width="90" align="center">8��</td>
                        <td width="90" align="center">9��</td>
                        <td width="90" align="center">10��</td>
                        <td width="90" align="center">11��</td>
                        <td width="90" align="center">12��</td>
                        <td width="90" align="center">��</td>
                      </tr>
				<!--  -->
				<% do until objrs.eof 	%>
				<% If IsNull(medflag) And IsNull(custname) Then %>
                  <tr  class="trbd" bgcolor="#FFFFC1" >
                        <td width="220" align="center" colspan="2">&nbsp;&nbsp;�հ�</td>
				  <% ElseIf Not IsNull(medflag) And IsNull(custname) then%>
                  <tr  class="trbd" bgcolor="#CCFFFF" >
                        <td width="220" align="center" colspan="2">&nbsp;&nbsp;<%=medflag%> �Ұ�</td>
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
