<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<object id="factory" style="display:none" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/hq/outdoor/inc/ScriptX.cab#Version=6,1,431,2">
</object>
<%	
	Dim pcontidx : pcontidx = request("contidx")
	Dim pcyear : pcyear = request("cyear")
	Dim pcmonth : pcmonth = request("cmonth")


'	response.write pcontidx
	' 광고 계약 기초 정보 
	Dim sql : sql = "select c.title, c.comment, c.mediummemo, c.regionmemo,  t.highcustcode, c.startdate, c.enddate  "
	sql = sql & " from wb_contact_mst c "
	sql = sql & "  left outer  join sc_cust_dtl t on c.custcode = t.custcode "
	sql = sql & "  left outer  join vw_contact_exe_monthly e on c.contidx = e.contidx and e.cyear = '" & pcyear & "' and e.cmonth = '" & pcmonth & "' "
	sql = sql & " where c.contidx = " & pcontidx 
'	response.write sql
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType =adCmdText
	Dim rs : Set rs = cmd.execute
	Set cmd = Nothing 
	If Not rs.eof Then 
		Dim title : title = rs("title")
		Dim comment : comment = rs("comment")
		Dim mediummemo : mediummemo = rs("mediummemo")
		Dim regionmemo : regionmemo = rs("regionmemo")
		Dim startdate : startdate = rs("startdate")
		Dim enddate : enddate = rs("enddate")
		If Not IsNull(comment) Then comment = Replace(comment, Chr(13)&Chr(10), "<br>")
		If Not IsNull(mediummemo) Then  mediummemo= Replace(mediummemo, Chr(13)&Chr(10), "<br>")
		If Not IsNull(regionmemo) Then  regionmemo= Replace(regionmemo, Chr(13)&Chr(10), "<br>")
	End If 

%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="http://10.110.10.86:6666/hq/outdoor/style.css" rel="stylesheet" type="text/css">
<script defer>
function PrintTest() {
  var factory = document.getElementById("factory");	
  factory.printing.header = "";   // Header에 들어갈 문장
  factory.printing.footer = "";   // Footer에 들어갈 문장
  factory.printing.portrait = false   // true 면 가로인쇄, false 면 세로 인쇄
  factory.printing.leftMargin = 1.0   // 왼쪽 여백 사이즈
  factory.printing.topMargin = 1.0   // 위 여백 사이즈
  factory.printing.rightMargin = 1.0  // 오른쪽 여백 사이즈
  factory.printing.bottomMargin = 0.1  // 아래 여백 사이즈
//  factory.printing.SetMarginMeasure(2); // 테두리 여백 사이즈 단위를 인치로 설정합니다.
 // factory.printing.printer = "HP DeskJet 870C";  // 프린트 할 프린터 이름
//  factory.printing.paperSize = "A4";   // 용지 사이즈
  factory.printing.paperSource = "Manual feed";   // 종이 Feed 방식
//  factory.printing.collate = true;   //  순서대로 출력하기
//  factory.printing.copies = 2;   // 인쇄할 매수
//  factory.printing.SetPageRange(false, 1, 3); // True로 설정하고 1, 3이면 1페이지에서 3페이지까지 출력
  factory.printing.Print(true) // 출력하기
}
</script>
<script type="text/javascript">
<!--
	window.onload = function () {
		PrintTest();
		//self.focus();
		//this.print();
		//this.close();
	}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  >
<!-- 계약 헤더 이미지 -->
<table width="1240"  class="title" align="center">
	<tr>
		<td><img src="/images/pop_top.gif" width="1240" height="60" align="absmiddle"></td>
	</tr>
</table>
<!-- // 계약 헤더 이미지 -->
<table width="1024"   align="center" style="margin-top:30px;">
	<tr>
		<td class="title"><img src="http://10.110.10.86:6666/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><%=title%> </td>
	</tr>
</table>
<% server.execute("/hq/outdoor/print/prt_contactsummary_s.asp") %>
<% server.execute("/hq/outdoor/print/prt_contactdetail_s.asp") %>
<% server.execute("/hq/outdoor/print/prt_reportphoto.asp") %>
<table width="1024" align="center" style="margin-top:10px;">
	<tr>
	  <th class="title" width='100' >매체특성</td>
	  <td width='684'  class="context" style="font-family:맑은 고딕;font-size:9px"><%=mediummemo%></td>
	</tr>
	<tr>
	  <th class="title" >지역특성</td>
	  <td  class="context" style="font-family:맑은 고딕;font-size:9px"><%=regionmemo %></td>
	</tr>
	<tr>
	  <th class="title" >특이사항</td>
	  <td  class="context" style="font-family:맑은 고딕;font-size:9px"><%=comment%></td>
	</tr>
</table>
</body>
</html>
