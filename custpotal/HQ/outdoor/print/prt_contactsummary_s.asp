<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<%
	Dim pcontidx : pcontidx = request("contidx")
	Dim pcyear : pcyear = request("cyear")
	Dim pcmonth : pcmonth = request("cmonth")
	' ���� ��� ���� ���� 
	Dim sql : sql = "select c.title, c.custcode, t.custname as teamname, c.firstdate, c.startdate, c.enddate, isnull(c.totalprice,0) as totalprice, isnull(e.monthly,0) as monthly, isnull(e.expense,0) as expense , c.comment, c.mediummemo, c.regionmemo  "
	sql = sql & " from wb_contact_mst c "
	sql = sql & "  left outer  join sc_cust_dtl t on c.custcode = t.custcode "
	sql = sql & "  left outer  join vw_contact_exe_monthly e on c.contidx = e.contidx and e.cyear = '" & pcyear & "' and e.cmonth = '" & pcmonth & "' "
	sql = sql & " where c.contidx = " & pcontidx 

'	response.write sql

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	Dim rs : Set rs = cmd.execute 
	custname = getcustname(rs("custcode"))
	teamname = getdeptname(rs("custcode"))

%>

<!-- // �˻� ���� �� ��� ���� �� -->
<input type="hidden" id="lastdate" value="<%=rs("enddate")%>" />
<table width="1024" align="center" style="margin-top:10px;">
	<tr>
		<th class="title">������</th>
		<td width="156" class="context"><%=getcustname(rs("custcode"))%></td>
		<th class="title">�����</th>
		<td width="156" class="context"><%=getdeptname(rs("custcode"))%></td>
		<th class="title">���</th>
		<td width="156" class="context"><%=rs("teamname")%></td>
		<th class="title">�Ѽ���</th>
		<td width="156" class="context"><%=FormatNumber(getmonthlyqty(pcontidx, pcyear, pcmonth),0)%></td>
	</tr>
	<tr>
		<th class="title">���Ⱓ</th>
		<td colspan="3" class="context"><%=rs("startdate")%>&nbsp;&nbsp; ~ &nbsp;&nbsp;<%=rs("enddate")%></td>
		<th class="title">���ʰ����</th>
		<td width="156" class="context"><%=rs("firstdate")%></td>
		<th class="title">�������</th>
		<td width="156" class="context"><%=FormatCurrency(rs("monthly"))%></td>
	</tr>
</table>
