
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
'	If Len(cmonth) = 1 Then cmonth = "0"&cmonth			' ���ۿ��� 1�ڸ��� 0�� �ٿ��� 2�ڸ� ���� ����
	cyear2 = request("cyear2")		' ����⵵
	if cyear2 = "" then cyear2 = year(date)		' ����⵵ �⺻ ����
	cmonth2 = request("cmonth2")' �����
	if cmonth2 = "" then cmonth2 = month(date)			' ����� �⺻ ����
'	If Len(cmonth2) = 1 Then cmonth2 = "0"&cmonth2	' ���� �ڸ��� ����

	dim total_monthprice, total_expense, total_income, total_incomeratio, prev

	Dim custcode : custcode = request("tcustcode")			'������ �ڵ�
	Dim custcode2 : custcode2 = request("tcustcode2")		'����� �ڵ�'����ڵ带 ���� ���ϸ� null �� ����

	dim objrs, sql
	' ���õ� �����ֿ� �ش��ϴ� ����μ� ����
	sql = "select custcode, custname from dbo.sc_cust_temp where highcustcode = '" & custcode & "'  AND MEDFLAG = 'A'  and attr10 = 1 order by custname"
	call get_recordset(objrs, sql)

	dim str
	' �ش� ����θ� �޺��ڽ��� ����
	str = "<select name='tcustcode2'>"
	do until objrs.eof
		str = str & "<option value='" & objrs("custcode") & "'"
			if custcode2 = objrs("custcode") then str = str & " selected"'���õ� ����ΰ� �����ϸ� ����θ� ������Ų��.
		str = str & ">" & objrs("custname") & "</option>"
		objrs.movenext
	Loop
	str = str & "</select>"
	objrs.close

	if custcode2 = "" or custcode = custcode2 then custcode2 = Null

	if request.cookies("class") = "D" or request.cookies("class") = "H"  then
		custcode2 = request.cookies("custcode2")
	end if

	if isnull(custcode2) then

	sql = "select c.custcode, c.custname as custname2, m.contidx, m.title, m.firstdate, m.startdate, m.enddate, m.canceldate, isnull(t.monthprice,0) as totalprice, isnull(sum(d.monthprice),0) as monthprice, isnull(sum(d.expense),0) as expense from dbo.wb_contact_mst m inner join dbo.vw_contact_totalprice t on m.contidx = t.contidx left outer join dbo.wb_contact_md_dtl d on m.contidx = d.contidx inner join dbo.sc_cust_temp c on m.custcode = c.custcode inner join dbo.sc_cust_temp c2 on c.highcustcode = c2.custcode where d.cyear = '"&cyear&"' and d.cmonth = '"&cmonth&"' and c2.custcode = '"&custcode&"' group by c.custcode,m.contidx, m.title, m.firstdate, m.startdate, m.enddate, m.canceldate, t.monthprice, c.custname with rollup having (c.custcode is not null and c.custname is not null and m.contidx is not null) or (c.custcode is not null and c.custname is null and m.contidx is null and m.title is null and  m.firstdate is null and m.startdate is null and m.enddate is null)"
'	response.write sql
	call get_recordset(objrs, sql)

	dim cnt, contidx, title, firstdate, startdate, enddate, period, monthprice, expense, income, incomeratio, custname2, totalprice,canceldate, prev_custname2 ,grand_total

	cnt = objrs.recordcount

	if not objrs.eof Then
		set contidx = objrs("contidx")
		set title = objrs("title")
		set firstdate = objrs("firstdate")
		set startdate = objrs("startdate")
		set enddate = objrs("enddate")
		set totalprice = objrs("totalprice")
		set monthprice = objrs("monthprice")
		set expense = objrs("expense")
		set canceldate = objrs("canceldate")
		set custname2 = objrs("custname2")
	end if

%>
				  <table width="1015" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" >
                      <tr class="trhd">
                        <td width="220" align="center" >��ü��</td>
                        <td width="75" align="center">����<br>�������</td>
                        <td width="80" align="center">������</td>
                        <td width="80" align="center">������</td>
                        <td width="80" align="center">�ѱ����</td>
                        <td width="80" align="center">�������</td>
                        <td width="80" align="center">�����޾�</td>
                        <td width="80" align="center">������</td>
                        <td width="60" align="center">������</td>
                        <td width="100" align="center">����μ�</td>
                      </tr>
	     <%
			do until objrs.eof
			if day(startdate) = "1" then
				period = datediff("m", startdate, enddate)+1
			else
				period = datediff("m", startdate, enddate)
			end if


		%>
		<% if  isnull(custname2) then %>
                  <tr class="trbd" bgcolor="#FFFFC1">
                    <td align="left"  style="padding-left:5px;"><%=prev_custname2%> �Ұ�</td>
		<% else %>
                  <tr class="trbd" bgcolor="#FFFFFF">
                    <td  align="left"  style="padding-left:5px;"><%=title%> </td>
		<% end if %>
                    <td  align="center"><%=firstdate%></td>
                    <td align="center"><%=startdate%></td>
                    <td align="center"><%=enddate%></td>
                    <td align="right"><%If Not IsNull(monthprice + expense) Then response.write formatnumber(monthprice + expense,0) Else response.write "0"%>&nbsp;</td>
                    <td align="right"><%If Not IsNull(monthprice) or monthprice <> 0 Then response.write formatnumber(monthprice,0) Else response.write "0"%>&nbsp;</td>
                    <td  align="right"><%If Not IsNull(expense) Then response.write formatnumber(expense,0) Else response.write "0"%>&nbsp;</td>
                    <td align="right"><%If expense <> 0  Then response.write formatnumber(monthprice-expense,0) Else response.write "0"%>&nbsp;</td>
                    <td align="right"><%If monthprice <> 0 Then response.write formatnumber((monthprice-expense)/monthprice*100, 2) Else response.write "0.00"%>&nbsp;</td>
                    <td  align="center"><%=custname2%>&nbsp;</td>
                  </tr>
				<%
						if  not isnull(custname2) then
							grand_total = grand_total + monthprice + expense
							total_monthprice = total_monthprice + monthprice
							total_expense = total_expense + expense
						end if
						prev_custname2 = custname2
						objrs.movenext

					loop
					objrs.close
					set objrs = nothing

					total_income = total_monthprice - total_expense
					if total_income = 0 then
						total_incomeratio = "0.00"
					else
						total_incomeratio = total_monthprice - total_expense / total_monthprice * 100
					end if

					if total_income <> 0 then
				%>
                  <tr height="40" class="trbd"  bgcolor="#FFC1C1">
                    <td  align="center"  >���հ� </td>
                    <td  align="center"  > </td>
                    <td  align="center"  > </td>
                    <td  align="center"  > </td>
                    <td  align="center"  > <%If Not IsNull(grand_total) Then response.write formatnumber(grand_total,0) Else response.write "0"%>&nbsp;</td>
                    <td  align="right" ><%If Not IsNull(total_monthprice) Then response.write formatnumber(total_monthprice,0) Else response.write "0"%>&nbsp;</td>
                    <td  align="right" ><%If Not IsNull(total_expense) Then response.write formatnumber(total_expense,0) Else response.write "0"%>&nbsp;</td>
                    <td align="right" ><%If total_monthprice <> 0  Then response.write formatnumber(total_monthprice-total_expense,0) Else response.write "0"%>&nbsp;</td>
                    <td  align="right" ><%If total_monthprice <> 0 Then response.write formatnumber((total_monthprice-total_expense)/total_monthprice*100, 2) Else response.write "0.00"%>&nbsp;</td>
                    <td  align="center">&nbsp;</td>
                  </tr>
				  <% end if %>
              </table>
<%
	else


	sql = "select m.contidx, m.title, m.firstdate, m.startdate, m.enddate, m.canceldate, isnull(t.monthprice,0) as totalprice, isnull(sum(d.monthprice),0) as monthprice, isnull(sum(d.expense),0) as expense, c.custname as custname2 from dbo.wb_contact_mst m inner join dbo.vw_contact_totalprice t on m.contidx = t.contidx inner join dbo.wb_contact_md_dtl d on m.contidx = d.contidx  inner join dbo.sc_cust_temp c on m.custcode = c.custcode left outer join dbo.sc_cust_temp c2 on c.highcustcode = c2.custcode where m.custcode = '"&custcode2&"' and d.cyear =  '"&cyear&"' and d.cmonth = '"&cmonth&"' group by m.contidx, m.title, m.firstdate, m.startdate, m.enddate, m.canceldate, t.monthprice, c.custname order by m.title"
	call get_recordset(objrs, sql)

	cnt = objrs.recordcount

	if not objrs.eof Then
		set contidx = objrs("contidx")
		set title = objrs("title")
		set firstdate = objrs("firstdate")
		set startdate = objrs("startdate")
		set enddate = objrs("enddate")
		set totalprice = objrs("totalprice")
		set monthprice = objrs("monthprice")
		set expense = objrs("expense")
		set canceldate = objrs("canceldate")
		set custname2 = objrs("custname2")
	end if

%>
				  <table width="1015" border="0" cellspacing="1" cellpadding="0" bgcolor="#CCCCCC" >
                      <tr class="trhd" >
                        <td width="220" align="center" >��ü��</td>
                        <td width="75" align="center">����<br>�������</td>
                        <td width="75" align="center">������</td>
                        <td width="75" align="center">������</td>
                        <td width="80" align="center">�ѱ����</td>
                        <td width="80" align="center">�������</td>
                        <td width="80" align="center">�����޾�</td>
                        <td width="80" align="center">������</td>
                        <td width="50" align="center">������</td>
                      </tr>
	     <%
			do until objrs.eof
			if day(startdate) = "1" then
				period = datediff("m", startdate, enddate)+1
			else
				period = datediff("m", startdate, enddate)
			end if

		%>
                  <tr class="trbd" bgcolor="#FFFFFF">
                    <td width="220" align="left"  style="padding-left:5px;"><%=title%></td>
                    <td width="75" align="center"><%=firstdate%></td>
                    <td width="75" align="center"><%=startdate%></td>
                    <td width="75" align="center"><%=enddate%></td>
                    <td width="80" align="right"><%If Not IsNull(monthprice + expense) Then response.write formatnumber(monthprice + expense,0) Else response.write "0"%>&nbsp;</td>
                    <td width="80" align="right"><%If Not IsNull(monthprice) or monthprice <> 0 Then response.write formatnumber(monthprice,0) Else response.write "0"%>&nbsp;</td>
                    <td width="80" align="right"><%If Not IsNull(expense) Then response.write formatnumber(expense,0) Else response.write "0"%>&nbsp;</td>
                    <td width="80" align="right"><%If expense <> 0  Then response.write formatnumber(monthprice-expense,0) Else response.write "0"%>&nbsp;</td>
                    <td width="50" align="right"><%If monthprice <> 0 Then response.write formatnumber((monthprice-expense)/monthprice*100, 2) Else response.write "0.00"%>&nbsp;</td>
                  </tr>
				<%
						grand_total = grand_total + monthprice + expense
						total_monthprice = total_monthprice + monthprice
						total_expense = total_expense + expense
						objrs.movenext
					loop
					objrs.close
					set objrs = nothing

					total_income = total_monthprice - total_expense
					if total_income = 0 then
						total_incomeratio = "0.00"
					else
						total_incomeratio = total_monthprice - total_expense / total_monthprice * 100
					end if

					if total_income <> 0 then
				%>
                  <tr height="40" class="trbd"  bgcolor="#FFC1C1">
                    <td  align="center"  >���հ� </td>
                    <td  align="center"  > </td>
                    <td  align="center"  > </td>
                    <td  align="center"  > </td>
                    <td  align="center"  > <%If Not IsNull(grand_total) Then response.write formatnumber(grand_total,0) Else response.write "0"%>&nbsp;</td>
                    <td width="80" align="right" ><%If Not IsNull(total_monthprice) Then response.write formatnumber(total_monthprice,0) Else response.write "0"%>&nbsp;</td>
                    <td width="80" align="right" ><%If Not IsNull(total_expense) Then response.write formatnumber(total_expense,0) Else response.write "0"%>&nbsp;</td>
                    <td width="80" align="right" ><%If total_monthprice <> 0  Then response.write formatnumber(total_monthprice-total_expense,0) Else response.write "0"%>&nbsp;</td>
                    <td width="50" align="right" ><%If total_monthprice <> 0 Then response.write formatnumber((total_monthprice-total_expense)/total_monthprice*100, 2) Else response.write "0.00"%>&nbsp;</td>
                  </tr>
				 <% end if %>
              </table>
<%
					end if%>

</body>
<SCRIPT LANGUAGE="JavaScript">
<!--
	var custcode2 = parent.document.getElementById("custcode2") ;
	custcode2.innerHTML = "<%=str%>";
//-->
</SCRIPT>
