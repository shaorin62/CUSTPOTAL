
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

	Dim custcode : custcode = request("tcustcode")			'����� �ڵ�'����ڵ带 ���� ���ϸ� null �� ����
	Dim custcode2 : custcode2 = request("tcustcode2")		'������ �ڵ�

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


	sql = " select isnull(custname, 'Z') as custname, seqname, x.medflag, custpart,  sum(P01) as 'P01', sum(P02) as 'P02', sum(P03) as 'P03',  sum(P04) as 'P04', sum(P05) as 'P05', sum(P06) as 'P06', sum(P07) as 'P07', sum(P08) as 'P08', sum(P09) as 'P09', sum(P10) as 'P10', sum(P11) as 'P11', sum(P12) as 'P12', sum(TOTAl) as 'TOTAL' from dbo.sc_cust_hdr c inner join   (select isnull(j.seqname, 'Z') as seqname, isnull(case when m.med_flag = '01' then 'TV' when m.med_flag in ('02','03') then 'Radio' end,'A') as medflag, case when p.custpart = 'Z' then 'Others' else ' ' + p.custpart end as custpart, m.clientcode , isnull(sum(case when substring(m.yearmon, 5,2) = '01' then isnull(amt,0) else 0 end),0) as 'P01' , isnull(sum(case when substring(m.yearmon, 5,2) = '02' then isnull(amt,0) else 0 end),0) as 'P02' , isnull(sum(case when substring(m.yearmon, 5,2) = '03' then isnull(amt,0) else 0 end),0) as 'P03' , isnull(sum(case when substring(m.yearmon, 5,2) = '04' then isnull(amt,0) else 0 end),0) as 'P04' , isnull(sum(case when substring(m.yearmon, 5,2) = '05' then isnull(amt,0) else 0 end),0) as 'P05' , isnull(sum(case when substring(m.yearmon, 5,2) = '06' then isnull(amt,0) else 0 end),0) as 'P06' , isnull(sum(case when substring(m.yearmon, 5,2) = '07' then isnull(amt,0) else 0 end),0) as 'P07' , isnull(sum(case when substring(m.yearmon, 5,2) = '08' then isnull(amt,0) else 0 end),0) as 'P08' , isnull(sum(case when substring(m.yearmon, 5,2) = '09' then isnull(amt,0) else 0 end),0) as 'P09' , isnull(sum(case when substring(m.yearmon, 5,2) = '10' then isnull(amt,0) else 0 end),0) as 'P10' , isnull(sum(case when substring(m.yearmon, 5,2) = '11' then isnull(amt,0) else 0 end),0) as 'P11' , isnull(sum(case when substring(m.yearmon, 5,2) = '12' then isnull(amt,0) else 0 end),0) as 'P12' , sum(isnull(amt,0)) as 'TOTAL'  from dbo.md_report_mst_v m inner join dbo.vw_cust_part p on m.medcode = p.custcode left outer join dbo.sc_subseq_dtl j on j.seqno = m.subseq where m.med_flag in ('01', '02', '03') and substring(m.yearmon , 1, 4) = '"&cyear&"' and m.clientcode = '"&custcode2&"' and amt <> 0 group by j.seqname, m.clientcode , case when m.med_flag = '01' then 'TV' when m.med_flag in ('02','03') then 'Radio' end , custpart , m.clientcode ) as x on c.highcustcode = x.clientcode group by c.custname, seqname, x.medflag, custpart with rollup "

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

%>
<table width="1420" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" >
  <tr class="trhd">
    <td width="180" align="center">�����귣��</td>
    <td width="90" align="center">����</td>
    <td width="90" align="center">������ü</td>
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
<!--  seqname, medflag, custpart, P01, P02, P03, P04, P05, P06, P07, P08, P09, P10, P11, P12, total, prev_seqname, prev_medflag -->
<%
	do until objrs.eof

	If Not (custname <> "Z" And Not IsNull(seqname) And IsNull(medflag) And IsNull(trim(custpart)) ) then
		If custname = "Z"  And  IsNull(seqname) And  IsNull(medflag) And IsNull(trim(custpart)) Then %> <!-- ���հ� -->
		  <tr class="trbd" bgcolor="#FFFFC1" >
			<td width="360" align="center" colspan="3">&nbsp;&nbsp;���հ�</td>
		<% Elseif custname <> "Z" And Not IsNull(seqname) And Not IsNull(medflag) And IsNull(trim(custpart)) Then%>
		  <tr class="trbd" bgcolor="#FFDFDF" >
			<td width="180" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;<%If prev_seqname <> seqname Then if seqname ="Z" then response.write "&nbsp;" else response.write seqname end if  Else response.write "&nbsp;" end if%></td>
			<td width="180" align="center" colspan="2">&nbsp;&nbsp;<%=medflag%> ��� </td>
		<% ElseIf custname <> "Z" And IsNull(seqname) And IsNull(medflag) And IsNull(trim(custpart))Then%>
		  <tr class="trbd" bgcolor="#CCFFFF" >
			<td width="360" align="center" colspan="3">&nbsp;&nbsp;<%=custname%> �Ұ�</td>
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
<% else

	sql="select isnull(custname, 'Z') as custname, seqname, x.medflag, custpart, sum(P01) as 'P01', sum(P02) as 'P02', sum(P03) as 'P03', sum(P04) as 'P04', sum(P05) as 'P05', sum(P06) as 'P06', sum(P07) as 'P07', sum(P08) as 'P08', sum(P09) as 'P09', sum(P10) as 'P10', sum(P11) as 'P11', sum(P12) as 'P12', sum(TOTAl) as 'TOTAL' from dbo.sc_cust_hdr c inner join (select isnull(j.seqname, 'Z') as seqname, isnull(case when m.med_flag = '01' then 'TV' when m.med_flag in ('02','03') then 'Radio' end,'A') as medflag, case when p.custpart = 'Z' then 'Others' else p.custpart end as custpart, m.clientcode , isnull(sum(case when substring(m.yearmon, 5,2) = '01' then isnull(amt,0) else 0 end),0) as 'P01' , isnull(sum(case when substring(m.yearmon, 5,2) = '02' then isnull(amt,0) else 0 end),0) as 'P02' , isnull(sum(case when substring(m.yearmon, 5,2) = '03' then isnull(amt,0) else 0 end),0) as 'P03' ,  isnull(sum(case when substring(m.yearmon, 5,2) = '04' then isnull(amt,0) else 0 end),0) as 'P04' ,  isnull(sum(case when substring(m.yearmon, 5,2) = '05' then isnull(amt,0) else 0 end),0) as 'P05' , isnull(sum(case when substring(m.yearmon, 5,2) = '06' then isnull(amt,0) else 0 end),0) as 'P06' , isnull(sum(case when substring(m.yearmon, 5,2) = '07' then isnull(amt,0) else 0 end),0) as 'P07' , isnull(sum(case when substring(m.yearmon, 5,2) = '08' then isnull(amt,0) else 0 end),0) as 'P08' , isnull(sum(case when substring(m.yearmon, 5,2) = '09' then isnull(amt,0) else 0 end),0) as 'P09' ,  isnull(sum(case when substring(m.yearmon, 5,2) = '10' then isnull(amt,0) else 0 end),0) as 'P10' , isnull(sum(case when substring(m.yearmon, 5,2) = '11' then isnull(amt,0) else 0 end),0) as 'P11' ,  isnull(sum(case when substring(m.yearmon, 5,2) = '12' then isnull(amt,0) else 0 end),0) as 'P12' ,  sum(isnull(amt,0)) as 'TOTAL' from dbo.md_report_mst_v m inner join dbo.vw_cust_part p on m.medcode = p.custcode left outer join dbo.sc_jobcust j on j.seqno = m.subseq where m.med_flag in ('01', '02', '03') and substring(m.yearmon , 1, 4) = '"&cyear&"' and m.timcode like '%' and m.clientcode = '"&custcode2 &"' and amt <> 0 group by j.seqname, m.clientcode , case when m.med_flag = '01' then 'TV' when m.med_flag in ('02','03') then 'Radio' end , custpart , m.clientcode ) as x on c.highcustcode = x.clientcode group by c.custname, seqname, x.medflag, custpart with rollup order by custname, x.medflag desc "


	call get_recordset(objrs, sql)

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

%>
<table width="1420" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" >
  <tr class="trhd">
    <td width="180" align="center">�����귣��</td>
    <td width="90" align="center">����</td>
    <td width="90" align="center">������ü</td>
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
<!--  seqname, medflag, custpart, P01, P02, P03, P04, P05, P06, P07, P08, P09, P10, P11, P12, total, prev_seqname, prev_medflag -->
<%
	do until objrs.eof
	if  not (custname<>"Z"  And   IsNull(seqname) And  IsNull(medflag) And IsNull(custpart)) Then '

		If custname = "Z"  And  IsNull(seqname) And  IsNull(medflag) And IsNull(custpart) Then %> <!-- ���հ� -->
		  <tr class="trbd" bgcolor="#FFC1C1" >
			<td width="360" align="center" colspan="3">&nbsp;&nbsp;���հ�</td>
		<% Elseif custname <> "Z" And Not IsNull(seqname) And Not IsNull(medflag) And IsNull(custpart) Then%>
		  <tr class="trbd" bgcolor="#FFDFDF" >
			<td width="180" align="center" bgcolor="#FFFFFF">&nbsp;&nbsp;<%If prev_seqname <> seqname Then if seqname ="Z" then response.write "&nbsp;" else response.write seqname end if  Else response.write "&nbsp;" end if%></td>
			<td width="180" align="center" colspan="2">&nbsp;&nbsp;<%=medflag%> ��� </td>
		<% ElseIf not isnull(custname) and not isnull(seqname) and isnull(medflag) and isnull(custpart) Then%>
		  <tr class="trbd" bgcolor="#CCFFFF" >
			<td width="360" align="center" colspan="3">&nbsp;&nbsp;<%if seqname = "Z" then response.write "&nbsp;" else response.write seqname%> �Ұ�</td>
		<% ElseIf custname = "Z" And IsNull(seqname) And IsNull(medflag) And IsNull(custpart)Then%>
		  <tr class="trbd" bgcolor="#CCFFFF" >
			<td width="360" align="center" colspan="3">&nbsp;&nbsp;<%=custname%> �Ұ�</td>
		<% Else %>
		  <tr class="trbd" bgcolor="#FFFFFF" >
			<td width="180" align="left" >&nbsp;&nbsp;<%If prev_seqname <> seqname Then if seqname ="Z" then response.write "&nbsp;" else response.write seqname end if  Else response.write "&nbsp;" end if%> </td>
			<td width="90" align="left">&nbsp;&nbsp;<%=medflag%> </td>
			<td width="90" align="left">&nbsp;&nbsp;<%=custpart%></td>
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
			If custname <> "Z" And IsNull(seqname) And IsNull(medflag) And IsNull(custpart) Then
				prev_seqname = ""
				prev_medflag = ""
			Else
				prev_seqname = seqname
				prev_medflag = medflag
			End if				'f custname <> "Z" And IsNull(seqname) And IsNull(medflag) And IsNull(custpart) Then
		end if					'not(not isnull(custname)  And   IsNull(seqname) And  IsNull(medflag) And IsNull(custpart))
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