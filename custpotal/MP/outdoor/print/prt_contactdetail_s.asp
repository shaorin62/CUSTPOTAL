<!--#include virtual="/MP/outdoor/inc/Function.asp" -->
<%
	'On Error Resume Next

	Dim pcontidx : pcontidx = request("contidx")
	Dim pcyear : pcyear = request("cyear")
	Dim pcmonth : pcmonth = request("cmonth")
	Dim pmdidx : pmdidx = request("mdidx")

	' 계약 매체 상세 정보
	dim sql
	sql = "select a.mdidx, b.side, isnull(monthly,0) monthly, isnull(expense,0) expense, a.locate, a.region, a.medcode, c.isHold, b.qty from wb_contact_md a inner join wb_contact_exe b on a.mdidx=b.mdidx and b.cyear='"&pcyear&"' and b.cmonth='"&pcmonth&"' left outer join wb_contact_trans c on a.contidx=c.contidx and a.medcode=c.medcode and c.cyear='"&pcyear&"' and c.cmonth = '"&pcmonth&"' where a.contidx = " & pcontidx &" order by case when b.side <> 'L' then ' ' +b.side else b.side end desc"

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
	dim locate
	dim region
	dim medcode
	dim isHold
	dim qty

	If Not rs.eof Then
		set mdidx = rs(0)
		set side = rs(1)
		set monthly = rs(2)
		set expense = rs(3)
		set locate = rs(4)
		set region = rs(5)
		set medcode = rs(6)
		set isHold = rs(7)
		set qty = rs(8)

		If pmdidx = "" then pmdidx = mdidx
	End If
%>
<P>
<table width="1024" align="center" cellpadding=0>
	<thead>
		<tr>
			<th width="200" class="detail">매체위치</th>
			<th width="140" class="detail">규격 / 재질</th>
			<th width="30" class="detail">수량</th>
			<th width="85" class="detail">브랜드</th>
			<th width="85" class="detail">소재</th>
			<th width="75" class="detail">월광고료</th>
			<th width="75" class="detail">월지급액</th>
			<th width="65" class="detail">내수액</th>
			<th width="50" class="detail">내수율</th>
			<th width="100" class="detail">매체사</th>
		</tr>
	</thead>
	<tbody>
<%
	If Not rs.eof Then
	Dim income, incomerate
	Do Until rs.eof

	income = monthly-expense
	If monthly = 0 Then incomerate = 0 Else incomerate = income/monthly*100
%>
		<tr>
			<td class="context2"  style="padding-left:3px;">[<%=region%>] <span title="<%=locate%>"><%=cutTitle(locate, 28)%></span> </td>
			<td class="context2"  style='text-align:center;' width="140" ><%=getcurrentstandard(mdidx, pcyear, pcmonth, side, "standard")%> <br><%=getcurrentstandard(mdidx, pcyear, pcmonth, side,"quality")%></td>
			<td class="context2"  style='text-align:center;'><%=FormatNumber(qty,0)%></td>
			<td class="context2"  style="padding-left:3px;text-align:center;"><%=getcurrentbrandname(mdidx, pcyear, pcmonth, side)%></td>
			<td class="context2"  style="padding-left:3px;text-align:center;"><%=getcurrentthemename(mdidx, pcyear, pcmonth, side)%></td>
			<td class="context2"  style="padding-right:5px; text-align:right;"><%=FormatNumber(monthly,0)%></td>
			<td class="context2"  style="padding-right:5px; text-align:right;"><%=FormatNumber(expense,0)%></td>
			<td class="context2"  style="padding-right:5px; text-align:right;"><%=FormatNumber(income,0)%></td>
			<td class="context2"  style="padding-right:10px; text-align:right;"><%=FormatNumber(incomerate, 2)%></td>
			<td class="context2"  style="padding-left:3px;" width="100"><%=getmedname(medcode)%></td>
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
				getside = "우측"
			Case "R"
				getside = "좌측"
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