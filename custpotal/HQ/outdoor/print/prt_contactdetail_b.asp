<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<%
	'On Error Resume Next

	Dim pcontidx : pcontidx = request("contidx")
	Dim pcyear : pcyear = request("cyear")
	Dim pcmonth : pcmonth = request("cmonth")
	Dim pside : pside = request("side")

	' 계약 매체 상세 정보
	Dim sql
	sql = "select a.mdidx, isnull(b.side,'') side, isnull(b.monthly,0) as monthly, isnull(b.expense,0) as expense, a.medcode, c.isHold from wb_contact_md a  inner join wb_contact_exe b on a.mdidx=b.mdidx and b.cyear='"&pcyear&"' and b.cmonth='"&pcmonth&"' left outer join wb_contact_trans c on a.medcode=c.medcode and c.cyear='"&pcyear&"' and c.cmonth='"&pcmonth&"' and a.contidx=c.contidx where a.contidx = " & pcontidx & " order by case when b.side<>'L' then ' ' +b.side else b.side end desc"
'	sql = "select distinct a.mdidx, b.side, b.standard, b.quality, isnull(c.monthly,0) as monthly , isnull(c.expense,0) as expense,
'
'	: sql = "select distinct a.mdidx, b.side, b.standard, b.quality, isnull(c.monthly,0) as monthly , isnull(c.expense,0) as expense, f.startdate, f.enddate, c.isHold "
'	sql = sql & " from wb_contact_md a  "
'	sql = sql & " inner join wb_contact_md_dtl b on a.mdidx = b.mdidx "
'	sql = sql & " left outer join wb_contact_exe c on b.mdidx=c.mdidx and b.side=c.side and c.cyear='"&pcyear&"' and c.cmonth='"&pcmonth&"'  "
'	sql = sql & " left outer join wb_subseq_exe d on c.cyear=d.cyear and c.cmonth=d.cmonth and c.mdidx=d.mdidx and c.side=d.side "
'	sql = sql & " left join wb_subseq_dtl e on d.thmno=e.thmno "
'	sql = sql & " left join wb_contact_mst f on a.contidx = f.contidx "
'	sql = sql & " where a.contidx =  " & pcontidx& " order by b.side desc"

'	response.write sql

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adcmdText
	Dim rs : Set rs = cmd.execute

	dim mdidx
	dim side
	dim monthly
	dim expense
	dim medcode
	dim isHold

	if not rs.eof then
		set mdidx = rs(0)
		set side = rs(1)
		set monthly = rs(2)
		set expense = rs(3)
		set medcode = rs(4)
		set isHold = rs(5)
	end if

'	Dim mdidx : Set mdidx = rs(0)
'	Dim side : Set side = rs(1)
'	Dim standard : Set standard = rs(2)
'	Dim quality : Set quality = rs(3)
'	Dim monthly : Set monthly = rs(4)
'	Dim expense : Set expense = rs(5)
'	Dim startdate : Set startdate = rs(6)
'	Dim enddate : Set enddate = rs(7)
'	Dim isHold : Set isHold = rs(8)
%>
<P>
<table width="1024" align="center" >
	<thead>
		<tr>
			<th width="50" class="detail">면</th>
			<th width="320" class="detail">규격 / 재질</th>
			<th width="110" class="detail">브랜드</th>
			<th width="110" class="detail">집행소재</th>
			<th width="100" class="detail">월광고료</th>
			<th width="100" class="detail">월지급액</th>
			<th width="75" class="detail">내수액</th>
			<th width="50" class="detail">내수율</th>
		</tr>
	</thead>
	<tbody>
<%
	If Not rs.eof Then
	Dim income, incomerate
	If pside = "" Then pside = side
	Do Until rs.eof

	income = monthly-expense
	If monthly = 0 Then incomerate = 0 Else incomerate = income/monthly*100
%>
		<tr>
			<td class="context2"  style='text-align:center;'><%=getside(Trim(side))%></td>
			<td class="context2"  style='text-align:center;'>  <%=getcurrentstandard(mdidx, pcyear, pcmonth, side, "standard")%> / <%=getcurrentstandard(mdidx, pcyear, pcmonth, side,"quality")%> </td>
			<td class="context2"  style="padding-left:3px;text-align:center;"><%=getcurrentbrandname(mdidx, pcyear, pcmonth, side)%></td>
			<td class="context2"  style="padding-left:3px;text-align:center;"><%=getcurrentthemename(mdidx, pcyear, pcmonth, side)%></td>
			<td class="context2"  style="padding-right:5px; text-align:right;"><%=FormatNumber(monthly,0)%></td>
			<td class="context2"  style="padding-right:5px; text-align:right;"><%=FormatNumber(expense,0)%></td>
			<td class="context2"  style="padding-right:5px; text-align:right;"><%=FormatNumber(income,0)%></td>
			<td class="context2"  style="padding-right:10px; text-align:right;"><%=FormatNumber(incomerate, 2)%></td>
		</tr>
<%
		rs.movenext
	Loop
	End If
%>
	</tbody>
</table>
<%
	Function  getside(side)
		Select Case side
			Case "F"
				getside = "정면"
			Case "B"
				getside = "후면"
			Case "L"
				getside = "좌측"
			Case "R"
				getside = "우측"
			Case Else
				getside = ""
		End Select
	End Function

	Function getbrand(thmno)
		Dim sql : sql = "select c.seqname from wb_subseq_dtl a inner join wb_subseq_mst b on a.subno=b.subno inner join sc_subseq_dtl c on b.seqno = c.seqno where a.thmno = '" & thmno &"' "
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		Set cmd = nothing
		If rs.eof Then getbrand = "" Else getbrand = rs(0)
	End Function

	Function getsubbrand(thmno)
		Dim sql : sql = "select b.subname from wb_subseq_dtl a inner join wb_subseq_mst b on a.subno=b.subno where a.thmno = '" & thmno &" "
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		Set cmd = nothing
		If rs.eof Then getsubbrand = "" Else getsubbrand = rs(0)
	End Function


	If Err.Number <> 0 Then
		Dim item
		For Each item In request.querystring
			response.write item & " : " & request.querystring(item) & "<br>"
		Next

		response.write "Err.Number : " & Err.number & "<br>"
		response.write "Err.Description : " & Err.Description & "<br>"
		response.write "Err.Source : " & Err.Source &"<br>"
		response.write sql
	End If
%>