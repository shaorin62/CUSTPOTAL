<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
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
<link href="http://10.110.10.86:6666/cust/outdoor/style.css" rel="stylesheet" type="text/css">
<script type="text/javascript">
<!--
	window.onload = function () {
		self.focus();
		//this.print();
		//this.close();
	}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  >
<table width="1024"   align="center" style="margin-top:30px;">
	<tr>
		<td class="title"><img src="http://10.110.10.86:6666/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><%=title%> </td>
	</tr>
</table>
<% server.execute("/cust/outdoor/print/prt_contactsummary_s.asp") %>
<% server.execute("/cust/outdoor/print/prt_contactdetail_s.asp") %>
<% server.execute("/cust/outdoor/print/prt_reportphoto.asp") %>
<table width="1024" align="center" style="margin-top:10px;">
	<tr>
	  <th class="title" width='100' >매체특성</td>
	  <td width='684'  class="context"><%=mediummemo%></td>
	</tr>
	<tr>
	  <th class="title" >지역특성</td>
	  <td  class="context"><%=regionmemo %></td>
	</tr>
	<tr>
	  <th class="title" >특이사항</td>
	  <td  class="context"><%=comment%></td>
	</tr>
</table>
</body>
</html>
