
<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<body oncontextmenu='return false' >
<%
	dim cyear, cyear2, cmonth, cmonth2, yearmon, yearmon2
	cyear = request("cyear")											' ���۳⵵
	if cyear = "" then cyear = year(date)							' ���۳⵵�� ������ ���� �⵵�� �⺻ �⵵�� ����
	cmonth = request("cmonth")									' ���ۿ�
	if cmonth = "" then cmonth = month(date)				' ���ۿ��� ������ ���� ���� �⺻ ���� ����
	If Len(cmonth) = 1 Then cmonth = "0"&cmonth			' ���ۿ��� 1�ڸ��� 0�� �ٿ��� 2�ڸ� ���� ����
	cyear2 = request("cyear2")										' ����⵵
	if cyear2 = "" then cyear2 = year(date)						' ����⵵ �⺻ ����
	cmonth2 = request("cmonth2")								' �����
	if cmonth2 = "" then cmonth2 = month(date)			' ����� �⺻ ����
	If Len(cmonth2) = 1 Then cmonth2 = "0"&cmonth2	' ���� �ڸ��� ����

	yearmon = cyear & cmonth										' ���۳�� ����
	yearmon2 = cyear2 & cmonth2									' ������ ����

	Dim custcode : custcode = request("tcustcode")			'������ �ڵ�
	Dim custcode2 : custcode2 = request("tcustcode2")		'����� �ڵ�				'����ڵ带 ���� ���ϸ� null �� ����

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

	'if custcode = custcode2 then 	custcode2 = null
	'if custcode2 = "" then custcode2 = Null

	if request.cookies("class") = "D" or request.cookies("class") = "H"  then
		custcode2 = request.cookies("custcode2")
	end if

	if not isnull(custcode2) then

	'sql = " select distinct c.custname, j.seqname, c2.custname as custname4, case when m.med_flag = '01' then 'TV' when m.med_flag in ('02','03') then 'Radio' end as med_flag,isnull(sum(case when m.real_med_code = 'B00107' then isnull(amt,0) else 0 end ),0) as 'A01',isnull(sum(case when m.real_med_code = 'B00111' then isnull(amt,0) else 0 end ),0) as 'A02', isnull(sum(case when m.real_med_code = 'B00109' then isnull(amt,0) else 0 end ),0) as 'A03', isnull(sum(case when m.real_med_code = 'B00108' then isnull(amt,0) else 0 end ),0) as 'A04', isnull(sum(case when m.real_med_code = 'B00110' then isnull(amt,0) else 0 end ),0) as 'A05', isnull(sum(case when m.real_med_code = 'B00112' then isnull(amt,0) else 0 end ),0) as 'A06', isnull(sum(amt),0) as 'TOTAL' ,isnull(sum(amt*1.1),0) as 'VAT_TOTAL' from dbo.md_report_mst_v m inner join dbo.sc_cust_hdr c on m.exclientcode = c.highcustcode left outer join dbo.sc_subseq_dtl j on j.seqno = m.subseq left outer join dbo.sc_cust_dtl c2 on c2.custcode = m.exclientcode where m.med_flag in ('01', '02', '03') and m.yearmon = '"&yearmon&"' and m.clientcode = '"&custcode2&"' and m.amt <> 0 group by c.custname, j.seqname, c2.custname, case when m.med_flag = '01' then 'TV' when m.med_flag in ('02','03') then 'Radio' end with cube having c.custname is not null or (c.custname is null and seqname is null and c2.custname is null) order by c.custname desc, seqname desc, custname4 desc, med_flag desc "

'	sql = "select distinct c.custname, j.seqname, c2.custname as custname4, case when m.medflag = '01' then 'TV' 		when m.medflag in ('02','03') then 'Radio' 		end as medflag,isnull(sum(case when m.real_med_code = 'B00140' then isnull(amt,0) else 0 end ),0) as 'A01',isnull(sum(case when m.real_med_code = 'B00141' then isnull(amt,0) else 0 end ),0) as 'A02',isnull(sum(case when m.real_med_code = 'B00142' then isnull(amt,0) else 0 end ),0) as 'A03',isnull(sum(case when m.real_med_code = 'B00143' then isnull(amt,0) else 0 end ),0) as 'A04',isnull(sum(case when m.real_med_code = 'B00144' then isnull(amt,0) else 0 end ),0) as 'A05',isnull(sum(case when m.real_med_code = 'B00145' then isnull(amt,0) else 0 end ),0) as 'A06',isnull(sum(amt),0) as 'TOTAL'	,isnull(sum(amt*1.1),0) as 'VAT_TOTAL'	 from dbo.md_report_mst m inner join dbo.sc_cust_temp c on m.clientsubcode = c.custcode left outer join dbo.sc_jobcust j on j.seqno = m.subseq left outer join dbo.sc_cust_temp c2 on c2.custcode = m.exclientcode where m.medflag in ('01', '02', '03') and m.yearmon = '"&yearmon&"' and m.clientcode = '"&custcode&"' and m.amt <> 0 group by c.custname, j.seqname, c2.custname, case when m.medflag = '01' then 'TV' when m.medflag in ('02','03') then 'Radio' end with cube having c.custname is not null or (c.custname is null and seqname is null and c2.custname is null)  order by c.custname desc, seqname desc, custname4 desc, medflag desc"


	sql = " select distinct c.custname, j.seqname, c2.custname as custname4,  case when m.med_flag = '01' then 'TV'  when m.med_flag in ('02','03') then 'Radio' end as medflag, isnull(sum(case when m.real_med_code = 'B00107' then isnull(amt,0) else 0 end ),0) as 'A01', isnull(sum(case when m.real_med_code = 'B00111' then isnull(amt,0) else 0 end ),0) as 'A02', isnull(sum(case when m.real_med_code = 'B00109' then isnull(amt,0) else 0 end ),0) as 'A03', isnull(sum(case when m.real_med_code = 'B00108' then isnull(amt,0) else 0 end ),0) as 'A04', isnull(sum(case when m.real_med_code = 'B00110' then isnull(amt,0) else 0 end ),0) as 'A05', isnull(sum(case when m.real_med_code = 'B00112' then isnull(amt,0) else 0  end ),0) as 'A06', isnull(sum(amt),0) as 'TOTAL' , isnull(sum(amt*1.1),0) as 'VAT_TOTAL'  from dbo.md_report_mst_v m  inner join dbo.sc_cust_dtl c  on m.timcode = c.custcode  left outer join dbo.sc_subseq_dtl j  on j.seqno = m.subseq  left outer join dbo.sc_cust_hdr c2  on c2.highcustcode = m.exclientcode  where m.med_flag in ('01', '02', '03') and m.yearmon = '"&yearmon&"'  and m.clientcode = '"&custcode2&"' and m.amt <> 0  group by c.custname, j.seqname, c2.custname, case when m.med_flag = '01' then 'TV'  when m.med_flag in ('02','03') then 'Radio' end with cube  having c.custname is not null or (c.custname is null and seqname is null and c2.custname is null)  order by c.custname desc, seqname desc, custname4 desc, medflag desc "



	'response.write sql
'	response.end


	call get_recordset(objrs, sql)

	Dim custname, seqname,custname4,  medflag,  A01, A02, A03, A04, A05, A06, total, vat_total,  prev_seqname, prev_custname, prev_customer_total, customer_total, prev_custname4, str_total, prev_total, prev_seqname_total, seqname_total
	If Not objrs.eof Then
		Set custname = objrs("custname")
		Set seqname = objrs("seqname")
		Set custname4 = objrs("custname4")
		Set medflag = objrs("medflag")
		Set A01 = objrs("A01")
		Set A02 = objrs("A02")
		Set A03 = objrs("A03")
		Set A04 = objrs("A04")
		Set A05 = objrs("A05")
		Set A06 = objrs("A06")
		Set total = objrs("total")
		Set vat_total = objrs("vat_total")
	End if

%>
				  <table width="1300" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" >
                      <tr class="trhd">
                        <td width="100" align="center">����</td>
                        <td width="120" align="center">�귣��</td>
                        <td width="150" align="center"> Creative <br> agency </td>
                        <td width="70" align="center" >����</td>
                        <td width="100" align="center">����</td>
                        <td width="100" align="center">�λ�</td>
                        <td width="100" align="center">�뱸</td>
                        <td width="100" align="center">����</td>
                        <td width="100" align="center">����</td>
                        <td width="100" align="center">����</td>
                        <td width="100" align="center">�հ�</td>
                        <td width="100" align="center">(VAT����)</td>
                      </tr>
				<!--  custname, seqname,custname4,  medflag,  A01, A02, A03, A04, A05, A06, total, vat_total,  prev_seqname, prev_medflag -->
				<% do until objrs.eof 	%>
				<% 'If Not (Not IsNull(custname) And IsNull(seqname) And Not IsNull(custname4)) then%>
				<% If Not IsNull(custname) And Not IsNull(seqname) And Not IsNull(custname4) And IsNull(medflag) then%>
                      <tr class="trbd" bgcolor="#FFFFFF" >
                        <td align="left" >&nbsp;&nbsp;<%If prev_custname <> custname Then response.write custname %> </td>
                        <td align="left">&nbsp;&nbsp;<%If prev_seqname <> seqname Then response.write seqname%> </td>
                        <td align="left">&nbsp;&nbsp;<%If prev_custname4 <> custname4 Then response.write custname4 %></td>
                        <td align="left">&nbsp;&nbsp;�κ���</td>
				<% ElseIf Not IsNull(custname) And Not IsNull(seqname) And IsNull(custname4) And Not IsNull(medflag) Then
					seqname_total = seqname & " TOTAL"%>
                      <tr class="trbd" bgcolor="#CCFFFF" ><!-- �귣�� �� -->
                        <td align="left" >&nbsp;&nbsp;<%If prev_custname <> custname Then response.write custname %> </td>
                        <td align="left" colspan="2">&nbsp;&nbsp;<% If seqname_total <> prev_seqname_total Then response.write seqname_total %></td>
                        <td align="left">&nbsp;&nbsp;<%=medflag%></td>
				<% ElseIf Not IsNull(custname) And Not IsNull(seqname) And IsNull(custname4) And IsNull(medflag) then
					seqname_total = seqname & " TOTAL"%>
                      <tr class="trbd" bgcolor="#CCFFFF" ><!-- �귣�� �κ��� -->
                        <td align="left" >&nbsp;&nbsp;<%If prev_custname <> custname Then response.write custname %> </td>
                        <td align="left" colspan="2">&nbsp;&nbsp;<% If seqname_total <> prev_seqname_total Then response.write seqname_total %> </td>
                        <td align="left">&nbsp;&nbsp;�κ���</td>
				<% ElseIf Not IsNull(custname) and IsNull(seqname) And not IsNull(custname4) And IsNull(medflag) then
					seqname_total = seqname & " TOTAL"%>
                      <tr class="trbd" bgcolor="#CCFFFF" ><!-- �귣�� �κ��� -->
                        <td align="left" >&nbsp;&nbsp;<%If prev_custname <> custname Then response.write custname %> </td>
                        <td align="left" colspan="2">&nbsp;&nbsp;<% If seqname_total <> prev_seqname_total Then response.write seqname_total %> </td>
                        <td align="left">&nbsp;&nbsp;�κ���</td>
				<% ElseIf Not IsNull(custname) And IsNull(seqname) And IsNull(custname4) And Not IsNull(medflag) then
						customer_total = custname&" TOTAL" %>
                      <tr class="trbd" bgcolor="#FFFFC1" > <!-- ����� TV, Radio  ��-->

                        <td align="left" colspan="3">&nbsp;&nbsp;<%If customer_total <> prev_customer_total Then response.write customer_total%> </td>
                        <td align="left">&nbsp;&nbsp;<%=medflag%></td>
				<% ElseIf Not IsNull(custname) And IsNull(seqname) And IsNull(custname4) And  IsNull(medflag) then
						customer_total = custname&" TOTAL"%>
                      <tr class="trbd" bgcolor="#FFFFC1" > <!-- ����� TV, Radio .�κ��� -->
                        <td align="left" colspan="3">&nbsp;&nbsp;<%If customer_total <> prev_customer_total Then response.write customer_total %></td>
                        <td align="left">&nbsp;&nbsp;�κ���</td>
				<% ElseIf  IsNull(custname) And IsNull(seqname) And IsNull(custname4) And Not  IsNull(medflag) Then
					str_total = "TOTAL"%>
                      <tr class="trbd" bgcolor="#FFC1C1" >
                        <td  align="center" colspan="3">&nbsp;&nbsp;<%If prev_total <> str_total Then response.write str_total%>  </td>
                        <td align="left">&nbsp;&nbsp;<%=medflag%></td>
				<% ElseIf  IsNull(custname) And IsNull(seqname) And IsNull(custname4) And  IsNull(medflag) Then
					str_total = "TOTAL"%>
                      <tr class="trbd" bgcolor="#FFC1C1" >
                        <td  align="center" colspan="3">&nbsp;&nbsp;<%If prev_total <> str_total Then response.write str_total%> </td>
                        <td align="left">&nbsp;&nbsp;�κ���</td>
				<% Else %>
                      <tr class="trbd" bgcolor="#FFFFFF" >
                        <td align="left" >&nbsp;&nbsp;<%If prev_custname <> custname Then response.write custname %> </td>
                        <td align="left">&nbsp;&nbsp;<%If prev_seqname <> seqname Then response.write seqname%> </td>
                        <td align="left">&nbsp;&nbsp;<%If prev_custname4 <> custname4 Then response.write custname4 %></td>
                        <td align="left">&nbsp;&nbsp;<%=medflag%></td>
				<% End if%>
                        <td align="right"><% If A01 = "0" Then response.write "-" Else response.write FormatNumber(A01,0)%>&nbsp;&nbsp;</td>
                        <td align="right"><% If A02 = "0" Then response.write "-" Else response.write FormatNumber(A02,0)%>&nbsp;&nbsp;</td>
                        <td align="right"><% If A03 = "0" Then response.write "-" Else response.write FormatNumber(A03,0)%>&nbsp;&nbsp;</td>
                        <td align="right"><% If A04 = "0" Then response.write "-" Else response.write FormatNumber(A04,0)%>&nbsp;&nbsp;</td>
                        <td align="right"><% If A05 = "0" Then response.write "-" Else response.write FormatNumber(A05,0)%>&nbsp;&nbsp;</td>
                        <td align="right"><% If A06 = "0" Then response.write "-" Else response.write FormatNumber(A06,0)%>&nbsp;&nbsp;</td>
                        <td align="right"><% If total = "0" Then response.write "-" Else response.write FormatNumber(total,0)%>&nbsp;&nbsp;</td>
                        <td align="right"><% If vat_total = "0" Then response.write "-" Else response.write FormatNumber(vat_total,0)%>&nbsp;&nbsp;</td>
				<%
							If  (Not IsNull(custname) And IsNull(seqname) And IsNull(custname4) And  IsNull(medflag)) Or (Not IsNull(custname) And IsNull(seqname) And IsNull(custname4) And Not IsNull(medflag) ) Then
								prev_seqname = ""
								prev_customer_total = customer_total
								prev_seqname_total = seqname_total
							ElseIf (Not IsNull(custname) And Not IsNull(seqname) And IsNull(custname4) And Not IsNull(medflag)) Or (Not IsNull(custname) And Not IsNull(seqname) And IsNull(custname4) And IsNull(medflag) ) Then
								prev_custname4 = ""
								prev_seqname_total = seqname_total
							Else
								prev_customer_total = customer_total
								prev_custname = custname
								prev_seqname = seqname
								prev_custname4 = custname4
								prev_total = str_total
							End if
						'End if
						objrs.movenext
					loop
					objrs.close
					set objrs = nothing
				%>
              </table>
<% else

	sql = "select distinct c.custname, j.seqname, c2.custname as custname4, case when m.medflag = '01' then 'TV' 		when m.medflag in ('02','03') then 'Radio' 		end as medflag,isnull(sum(case when m.real_med_code = 'B00140' then isnull(amt,0) else 0 end ),0) as 'A01',isnull(sum(case when m.real_med_code = 'B00141' then isnull(amt,0) else 0 end ),0) as 'A02',isnull(sum(case when m.real_med_code = 'B00142' then isnull(amt,0) else 0 end ),0) as 'A03',isnull(sum(case when m.real_med_code = 'B00143' then isnull(amt,0) else 0 end ),0) as 'A04',isnull(sum(case when m.real_med_code = 'B00144' then isnull(amt,0) else 0 end ),0) as 'A05',isnull(sum(case when m.real_med_code = 'B00145' then isnull(amt,0) else 0 end ),0) as 'A06',isnull(sum(amt),0) as 'TOTAL'	,isnull(sum(amt*1.1),0) as 'VAT_TOTAL'	 from dbo.md_report_mst m inner join dbo.sc_cust_temp c on m.clientsubcode = c.custcode left outer join dbo.sc_jobcust j on j.seqno = m.subseq left outer join dbo.sc_cust_temp c2 on c2.custcode = m.exclientcode where m.medflag in ('01', '02', '03') and m.yearmon = '"&yearmon&"' and m.clientcode = '"&custcode&"' and m.clientsubcode like '"& custcode2 &"%' and  m.amt <> 0 group by c.custname, j.seqname, c2.custname, case when m.medflag = '01' then 'TV' when m.medflag in ('02','03') then 'Radio' end with cube having c.custname is not null or (c.custname is not null and seqname is null and c2.custname is null)  order by c.custname desc, seqname desc, custname4 desc, medflag desc"

	call get_recordset(objrs, sql)

	If Not objrs.eof Then
		Set custname = objrs("custname")
		Set seqname = objrs("seqname")
		Set custname4 = objrs("custname4")
		Set medflag = objrs("medflag")
		Set A01 = objrs("A01")
		Set A02 = objrs("A02")
		Set A03 = objrs("A03")
		Set A04 = objrs("A04")
		Set A05 = objrs("A05")
		Set A06 = objrs("A06")
		Set total = objrs("total")
		Set vat_total = objrs("vat_total")
	End if

%>
				  <table width="1200" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" >
                      <tr class="trhd">
                        <td width="120" align="center">�귣��</td>
                        <td width="150" align="center"> Creative <br> agency </td>
                        <td width="70" align="center" >����</td>
                        <td width="100" align="center">����</td>
                        <td width="100" align="center">�λ�</td>
                        <td width="100" align="center">�뱸</td>
                        <td width="100" align="center">����</td>
                        <td width="100" align="center">����</td>
                        <td width="100" align="center">����</td>
                        <td width="100" align="center">�հ�</td>
                        <td width="100" align="center">(VAT����)</td>
                      </tr>
				<!--  custname, seqname,custname4,  medflag,  A01, A02, A03, A04, A05, A06, total, vat_total,  prev_seqname, prev_medflag -->
				<% do until objrs.eof 	%>
				<%' If Not (IsNull(seqname) And Not IsNull(custname4)) then%>
				<% If  Not IsNull(seqname) And Not IsNull(custname4) And IsNull(medflag) then%>
                      <tr class="trbd" bgcolor="#FFFFFF" >
                        <td align="left">&nbsp;&nbsp;<%If prev_seqname <> seqname Then response.write seqname%> </td>
                        <td align="left">&nbsp;&nbsp;<%If prev_custname4 <> custname4 Then response.write custname4 %></td>
                        <td align="left">&nbsp;&nbsp;�κ���</td>
				<% ElseIf Not IsNull(seqname) And IsNull(custname4) And Not IsNull(medflag) Then
					seqname_total = seqname & " TOTAL"%>
                      <tr class="trbd" bgcolor="#CCFFFF" >
                        <td align="left" colspan="2">&nbsp;&nbsp;<% If seqname_total <> prev_seqname_total Then response.write seqname_total %></td>
                        <td align="left">&nbsp;&nbsp;<%=medflag%></td>
				<% ElseIf  Not IsNull(seqname) And IsNull(custname4) And IsNull(medflag) then
					seqname_total = seqname & " TOTAL"%>
                      <tr class="trbd" bgcolor="#CCFFFF" >
                        <td align="left" colspan="2">&nbsp;&nbsp;<% If seqname_total <> prev_seqname_total Then response.write seqname_total %></td>
                        <td align="left">&nbsp;&nbsp;�κ���</td>
				<% ElseIf   IsNull(seqname) And not IsNull(custname4) And IsNull(medflag) then
					seqname_total = seqname & " TOTAL"%>
                      <tr class="trbd" bgcolor="#CCFFFF" >
                        <td align="left" colspan="2">&nbsp;&nbsp;<% If seqname_total <> prev_seqname_total Then response.write seqname_total %></td>
                        <td align="left">&nbsp;&nbsp;�κ���</td>
				<% ElseIf   IsNull(seqname) And IsNull(custname4) And Not  IsNull(medflag) Then
					str_total = "TOTAL"%>
                      <tr class="trbd" bgcolor="#FFC1C1" >
                        <td  align="center" colspan="2">&nbsp;&nbsp;<%If prev_total <> str_total Then response.write str_total%>  </td>
                        <td align="left">&nbsp;&nbsp;<%=medflag%></td>
				<% ElseIf  IsNull(seqname) And IsNull(custname4) And  IsNull(medflag) Then
					str_total = "TOTAL"%>
                      <tr class="trbd" bgcolor="#FFC1C1" >
                        <td  align="center" colspan="2">&nbsp;&nbsp;<%If prev_total <> str_total Then response.write str_total%> </td>
                        <td align="left">&nbsp;&nbsp;�κ���</td>
				<% Else%>
                      <tr class="trbd" bgcolor="#FFFFFF" >
                        <td align="left">&nbsp;&nbsp;<%If prev_seqname <> seqname Then response.write seqname%> </td>
                        <td align="left">&nbsp;&nbsp;<%If prev_custname4 <> custname4 Then response.write custname4 %></td>
                        <td align="left">&nbsp;&nbsp; <%=medflag%></td>
				<% End if%>
                        <td align="right"><% If A01 = "0" Then response.write "-" Else response.write FormatNumber(A01,0)%>&nbsp;&nbsp;</td>
                        <td align="right"><% If A02 = "0" Then response.write "-" Else response.write FormatNumber(A02,0)%>&nbsp;&nbsp;</td>
                        <td align="right"><% If A03 = "0" Then response.write "-" Else response.write FormatNumber(A03,0)%>&nbsp;&nbsp;</td>
                        <td align="right"><% If A04 = "0" Then response.write "-" Else response.write FormatNumber(A04,0)%>&nbsp;&nbsp;</td>
                        <td align="right"><% If A05 = "0" Then response.write "-" Else response.write FormatNumber(A05,0)%>&nbsp;&nbsp;</td>
                        <td align="right"><% If A06 = "0" Then response.write "-" Else response.write FormatNumber(A06,0)%>&nbsp;&nbsp;</td>
                        <td align="right"><% If total = "0" Then response.write "-" Else response.write FormatNumber(total,0)%>&nbsp;&nbsp;</td>
                        <td align="right"><% If vat_total = "0" Then response.write "-" Else response.write FormatNumber(vat_total,0)%>&nbsp;&nbsp;</td>
					</tr>
				<%
							If  (IsNull(seqname) And IsNull(custname4) And  IsNull(medflag)) Or (IsNull(seqname) And IsNull(custname4) And Not IsNull(medflag) ) Then
								prev_seqname = ""
								prev_customer_total = customer_total
								prev_seqname_total = seqname_total
							ElseIf (Not IsNull(seqname) And IsNull(custname4) And Not IsNull(medflag)) Or (Not IsNull(seqname) And IsNull(custname4) And IsNull(medflag) ) Then
								prev_custname4 = ""
								prev_seqname_total = seqname_total
							Else
								prev_customer_total = customer_total
								prev_custname = custname
								prev_seqname = seqname
								prev_custname4 = custname4
								prev_total = str_total
							End if
						'End if
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