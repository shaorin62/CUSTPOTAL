<%@CODEPAGE=65001%>
<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
	On Error Resume Next
	Dim pcontidx : pcontidx = request("contidx")
	Dim pcyear : pcyear = request("cyear")
	Dim pcmonth : pcmonth = request("cmonth")
	' 광고 계약 기초 정보
	Dim sql : sql = "select c.title, c.custcode, t.custname as teamname, m.unit , c.firstdate, c.startdate, c.enddate, m.medcode, isnull(c.totalprice,0) as totalprice, isnull(e.monthly,0) as monthly, isnull(e.expense,0) as expense , m.locate, m.medclass, m.validclass, c.comment, c.mediummemo, c.regionmemo, m.map  "
	sql = sql & " from wb_contact_mst c "
	sql = sql & " left outer join wb_contact_md m on c.contidx = m.contidx "
	sql = sql & "  left outer  join sc_cust_dtl t on c.custcode = t.custcode "
	sql = sql & "  left outer  join vw_contact_exe_monthly e on c.contidx = e.contidx and e.cyear = '" & pcyear & "' and e.cmonth = '" & pcmonth & "' "
	sql = sql & " where c.contidx = " & pcontidx

'	response.write sql

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	Dim rs : Set rs = cmd.execute

%>

<!-- // 검색 조건 및 계약 관리 툴 -->
<input type="hidden" id="lastdate" value="<%=rs("enddate")%>" />
<table width="1024" align="center" style="margin-top:10px;">
	<tr>
		<th class="title">광고주</th>
		<td width="156" class="context"><%=getcustname(rs("custcode"))%></td>
		<th class="title">사업부</th>
		<td width="156" class="context"><%=getdeptname(rs("custcode"))%></td>
		<th class="title">운영팀</th>
		<td width="156" class="context"><%=rs("teamname")%></td>
		<th class="title">총수량</th>
		<td width="156" class="context"><%=FormatNumber(getmonthlyqty(pcontidx, pcyear, pcmonth),0)%>&nbsp;<%=rs("unit")%></td>
	</tr>
	<tr>
		<th class="title">계약기간</th>
		<td colspan="3" class="context"><%=rs("startdate")%>&nbsp;&nbsp; ~ &nbsp;&nbsp;<%=rs("enddate")%></td>
		<th class="title">최초계약일</th>
		<td width="156" class="context"><%=rs("firstdate")%></td>
		<th class="title">매체사</th>
		<td width="156" class="context"><%=getmedname(rs("medcode"))%></td>
	</tr>
	<tr>
		<th class="title">총광고료</th>
		<td width="156" class="context"><%=formatcurrency(rs("totalprice"))%></td>
		<th class="title">월광고료</th>
		<td width="156" class="context"><%=FormatCurrency(rs("monthly"))%></td>
		<th class="title"></th>
		<td width="156" class="context"><!-- 월지급액<%=FormatCurrency(rs("expense"))%> --></td>
		<th class="title"></th>
		<td width="156" class="context"><!--내수액(율) <%=FormatCurrency(rs("monthly") - rs("expense"))%> (<% If rs("monthly") <> 0 Then response.write FormatNumber(((rs("monthly")-rs("expense"))/rs("monthly")*100),2) Else response.write "0.00"%>) --></td>
	</tr>
	<tr>
		<th class="title">매체위치</th>
		<td colspan="3" class="context"><%=rs("locate")%></td>
		<th class="title">매체등급</th>
		<td width="156" class="context"><%=rs("medclass")%></td>
		<th class="title">효용성등급</th>
		<td width="156" class="context"><%=rs("validclass")%></td>
	</tr>
</table>
