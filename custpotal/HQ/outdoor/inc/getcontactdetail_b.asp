<%@CODEPAGE=65001%>
<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<%
	'On Error Resume Next

	Dim pcontidx : pcontidx = request("contidx")
	Dim pcyear : pcyear = request("cyear")
	Dim pcmonth : pcmonth = request("cmonth")
	Dim pside : pside = request("side")

	' 계약 매체 상세 정보
	Dim sql
	sql = "select a.mdidx, isnull(b.side,'') side, isnull(b.monthly,0) as monthly, isnull(b.expense,0) as expense, case when isnull(c.medcode,'') = '' then a.medcode else c.medcode end medcode, c.isHold from wb_contact_md a  inner join wb_contact_exe b on a.mdidx=b.mdidx and b.cyear='"&pcyear&"' and b.cmonth='"&pcmonth&"' left outer join wb_contact_trans c on a.contidx=c.contidx    and a.medcode = c.medcode and c.cyear='"&pcyear&"' and c.cmonth='"&pcmonth&"' and a.contidx=c.contidx where a.contidx = " & pcontidx & " order by case when b.side<>'L' then ' ' +b.side else b.side end desc"
'	sql = "select distinct a.mdidx, b.side, b.standard, b.quality, isnull(c.monthly,0) as monthly , isnull(c.expense,0) as expense, f.startdate, f.enddate, c.isHold "
'	sql = sql & " from wb_contact_md a  "
'	sql = sql & " left outer join wb_contact_md_dtl b on a.mdidx = b.mdidx and b.cyear+b.cyear<'" & pcyear&pcmonth& "'"
'	sql = sql & " left outer join wb_contact_exe c on b.mdidx=c.mdidx and b.side=c.side and c.cyear='"&pcyear&"' and c.cmonth='"&pcmonth&"'  "
'	sql = sql & " left outer join wb_subseq_exe d on c.cyear=d.cyear and c.cmonth=d.cmonth and c.mdidx=d.mdidx and c.side=d.side "
'	sql = sql & " left join wb_subseq_dtl e on d.thmno=e.thmno "
'	sql = sql & " left join wb_contact_mst f on a.contidx = f.contidx "
'	sql = sql & " where a.contidx =  " & pcontidx& " order by b.side desc"
'
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
<table width="1024" align="center" style="margin-top:10px;" border=0>
	<tr>
		<td width='712' height='25'><a href="#" onclick="getmedium('c'); return false;"><img src='/images/m_add.gif' width='14' height='15' alt="광고 매체 추가"></a>  면 추가 (*우측 상단의 매체관리 버튼을 클릭하여 매체기본 정보를 먼저 입력하신 후 면을 등록하세요) </td>
		<td width='312' align='right'><img src='/images/m_subseq.gif' width='16' height='17' alt="매체별 소재 관리"> 소재 <img src='/images/m_money.gif' width='16' height'='15' > 광고료 <img src='/images/m_photo.gif' width='16' height='15' > 사진 <img src='/images/m_edit.gif' width='16' height='15' > 수정 <img src='/images/m_delete.gif' width='16' height='15' alt="매체 정보 삭제" > 삭제 </td>
	</tr>
</table>
<table width="1024" align="center" >
	<thead>
		<tr>
			<th width="20" class="detail">&nbsp;</th>
			<th width="50" class="detail">면</th>
			<th width="320" class="detail">규격 / 재질</th>
			<th width="110" class="detail">브랜드</th>
			<th width="110" class="detail">집행소재</th>
			<th width="100" class="detail">월광고료</th>
			<th width="100" class="detail">월지급액</th>
			<th width="75" class="detail">내수액</th>
			<th width="50" class="detail">내수율</th>
			<th width="89" class="detail">&nbsp;</th>
		</tr>
	</thead>
	<tbody>
<%
	If Not rs.eof Then
	Dim income, incomerate
	If pside = "" Then pside = trim(side)
	Do Until rs.eof
	income = monthly-expense
	If monthly = 0 Then incomerate = 0 Else incomerate = income/monthly*100
%>
		<tr>
			<td class="context2"  style='text-align:center;'><input type='checkbox' name="side" value="<%=side%>"  onclick="setitem(this); getcontactphoto();" class="side" <%If CStr(pside) = CStr(side) Then response.write " checked"%>></td>
			<td class="context2"  style='text-align:center;'><%=getside(Trim(side))%></td>
			<td class="context2"  style='text-align:center;'> <%=getcurrentstandard(mdidx, pcyear, pcmonth, side, "standard")%> / <%=getcurrentstandard(mdidx, pcyear, pcmonth, side,"quality")%> </td>
			<td class="context2"  style="padding-left:3px;text-align:center;"><%=getcurrentbrandname(mdidx, pcyear, pcmonth, side)%></td>
			<td class="context2"  style="padding-left:3px;text-align:center;"><%=getcurrentthemename(mdidx, pcyear, pcmonth, side)%></td>
			<td class="context2"  style="padding-right:5px; text-align:right;"><%=FormatNumber(monthly,0)%></td>
			<td class="context2"  style="padding-right:5px; text-align:right;"><%=FormatNumber(expense,0)%></td>
			<td class="context2"  style="padding-right:5px; text-align:right;"><%=FormatNumber(income,0)%></td>
			<td class="context2"  style="padding-right:10px; text-align:right;"><%=FormatNumber(incomerate, 2)%></td>
			<td class="context2"  style=' text-align:center;'><a href="#" onclick="gettheme(<%=mdidx%>, '<%=side%>'); return false;"><img src='/images/m_subseq.gif' width='16' height='17' hspace=1 alt="매체별 소재 관리"></a><a href="#" onclick="getaccount(<%=mdidx%>, '<%=side%>'); return false;"><img src='/images/m_money.gif' width='16' height'='15' alt="광고 비용 관리" hspace=2></a><a href="#" onclick="getphoto(<%=mdidx%>, '<%=side%>'); return false;"><img src='/images/m_photo.gif' width='16' height='15' alt="매체 사진 관리" hspace=1></a><% If Not Len(isHold) Then %><A HREF="#" onclick="setmedium(); return false;"><img src='/images/m_edit.gif' width='16' height='15' alt="매체 정보 수정" hspace=2 class='<%=side%>'></A><img src='/images/m_delete.gif' width='16' height='15' alt="매체 정보 삭제" hspace=1 class='<%=side%>'><%Else%><a href="#" onclick="getmedium('u'); return false;"><img src='/images/m_edit.gif' width='16' height='15' alt="매체 정보 수정" hspace=2 class='<%=side%>'></a><a href="#" onclick="if (confirm('선택한 매체정보를 삭제하시겠습니까?\n\n매체에 등록된 광고비, 소재 정보도 모두 삭제됩니다.')) {getmedium('d');}; return false;"><img src='/images/m_delete.gif' width='16' height='15' alt="매체 정보 삭제" hspace=1 class='<%=side%>'></a><%End If%></td>
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