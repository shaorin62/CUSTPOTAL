<!--#include virtual="/MP/outdoor/inc/Function.asp" -->
<%
	Dim pcontidx : pcontidx = request("contidx")
	Dim pcyear : pcyear = request("cyear")
	Dim pcmonth : pcmonth = request("cmonth")
	' ���� ��� ���� ����
	Dim sql : sql = "select c.title, c.custcode, t.custname as teamname, m.unit , c.firstdate, c.startdate, c.enddate, m.medcode, isnull(c.totalprice,0) as totalprice, isnull(e.monthly,0) as monthly, isnull(e.expense,0) as expense , m.locate, m.medclass, m.validclass, c.comment, c.mediummemo, c.regionmemo, m.map  "
	sql = sql & " from wb_contact_mst c "
	sql = sql & " left outer join wb_contact_md m on c.contidx = m.contidx "
	sql = sql & "  left outer  join sc_cust_dtl t on c.custcode = t.custcode "
	sql = sql & "  left outer  join vw_contact_exe_monthly e on c.contidx = e.contidx and e.cyear = '" & pcyear & "' and e.cmonth = '" & pcmonth & "' "
	sql = sql & " where c.contidx = " & pcontidx

'	response.write sql
'	response.end

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	Dim rs : Set rs = cmd.execute

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
		<td width="156" class="context"><%=FormatNumber(getmonthlyqty(pcontidx, pcyear, pcmonth),0)%>&nbsp;<%=rs("unit")%></td>
	</tr>
	<tr>
		<th class="title">���Ⱓ</th>
		<td colspan="3" class="context"><%=rs("startdate")%>&nbsp;&nbsp; ~ &nbsp;&nbsp;<%=rs("enddate")%></td>
		<th class="title">���ʰ����</th>
		<td width="156" class="context"><%=rs("firstdate")%></td>
		<th class="title">��ü��</th>
		<td width="156" class="context"><%=getmedname(rs("medcode"))%></td>
	</tr>
	<tr>
		<th class="title">��ü��ġ</th>
		<td colspan="3" class="context"><%=rs("locate")%></td>
		<th class="title">�������</th>
		<td width="156" class="context"><%=FormatCurrency(rs("monthly"))%></td>
		<th class="title">&nbsp;</th>
		<td width="156" class="context">&nbsp;</td>
	</tr>
</table>
