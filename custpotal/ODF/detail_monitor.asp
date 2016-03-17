<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<%
'	Dim item
'	For Each item In request.querystring
'		response.write item &  " : " & request.querystring(item) & "<br>"
'	Next

	If request.cookies("userid") = "" Then response.redirect "/"
	Dim pmdidx : pmdidx = request("mdidx")
	Dim pside : pside = request("side")
	Dim pcustcode : pcustcode = request("custcode")
	Dim pteamcode : pteamcode = request("teamcode")
	Dim pcyear : pcyear = request("cyear")
	Dim pcmonth : pcmonth = request("cmonth")
	Dim pnum : pnum = request("num")
	If pnum="" Then pnum=0
	If pcyear = "" Then pcyear = Year(date)
	If pcmonth = "" Then pcmonth = Month(date)
	If Len(pcmonth) = 1 Then pcmonth = "0"&pcmonth

	Dim sql : sql = "select a.contidx, a.title, c.side, a.custcode, b.medcode,  a.startdate, a.enddate, b.region, b.locate, c.standard, e.codename quality, d.cdate, d.num, d.status, d.cname , d.comment, d.img01, d.img02, d.img03, d.img04 from wb_contact_mst a inner join wb_contact_md  b on a.contidx=b.contidx inner join wb_contact_md_dtl c on b.mdidx=c.mdidx and c.side=? left outer join wb_contact_monitor d on c.mdidx=d.mdidx and c.side=d.side and d.cyear=? and d.cmonth=? and num= ? inner join wb_code_library e on c.quality=e.codevalue where b.mdidx=? and e.category ='quality' "

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adcmdtext
	cmd.parameters.append cmd.createparameter("side", adchar, adparaminput, 1)
	cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4)
	cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2)
	cmd.parameters.append cmd.createparameter("num", adUnsignedTinyInt, adparaminput)
	cmd.parameters.append cmd.createparameter("mdidx", adinteger, adparaminput)
	cmd.parameters("side").value = pside
	cmd.parameters("cyear").value = pcyear
	cmd.parameters("cmonth").value = pcmonth
	cmd.parameters("num").value = pnum
	cmd.parameters("mdidx").value = pmdidx
	Dim rs : Set rs = cmd.execute
	Set cmd = Nothing

%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/hq/outdoor/style.css" rel="stylesheet" type="text/css">
<script type='text/javascript' src='/js/ajax.js'></script>
<script type='text/javascript' src='/js/script.js'></script>
<script type="text/javascript">
<!--
	var viewmonitor=null ;

	function preview() {
		var clickElement = event.srcElement ;
		var preimg = document.getElementById("preimg");
		preimg.src = clickElement.src;
		preimg.style.width="600";
		preimg.style.height="450";
	}

	function getmonitor(crud) {
		var url = "/odf/popup/view_monitor.asp?custcode=<%=pcustcode%>&teamcode=<%=pteamcode%>&mdidx=<%=pmdidx%>&side=<%=pside%>&cyear=<%=pcyear%>&cmonth=<%=pcmonth%>&num=<%=pnum%>&title=<%=rs("title")%>&crud="+crud;
		var name = "viewmonitor";
		var left = screen.width / 2 - 550 / 2;
		var top = screen.height / 2 - 442 / 2;
		var opt = "width=550, height=442, resizable=no, scrollbars=no, status=yes, left="+left+", top="+top;
		viewmonitor= window.open(url, name, opt);
	}

	window.onload = function () {
	}

	window.onunload = function () {
		if (viewmonitor) {viewmonitor.close();}
	}

//-->
</script>
</head>
<!-- <body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"   oncontextmenu="return false">
 --><body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"   >
<form action='list_monitor.asp' method='POST'>
<input type="hidden" id="orgnum" name='orgnum' value='<%=rs("num")%>' />
<table width="1240" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td height="24" background="/images/pop_top.gif" valign="top" >
	<% if request.cookies("class") = "F" then %><table width="700"  border="0" align="right" cellpadding="0" cellspacing="0" height="60">
      <tr style="padding-top:10">
        <td>&nbsp;</td>
        <td width="224" align="right" valign="top" ><span class="log">&nbsp;<%=request.cookies("custname")%></span> &nbsp;</td>
        <td width="1" valign="top"><img src="/images/top_vline_bg.gif" width="1" height="32"></td>
        <td width="104" align="right" valign="top" ><span class="log">&nbsp;<%=request.cookies("userid")%></span> &nbsp;</td>
        <td width="1" valign="top" ><img src="/images/top_vline_bg.gif" width="1" height="32"></td>
        <td width="184" align="right" valign="top" ><span class="log"><%=request.cookies("logtime")%>&nbsp;</span></td>
        <td width="1" valign="top"><img src="/images/top_vline_bg.gif" width="1" height="32"></td>
        <td width="85" align="center"valign="top" ><A HREF="/Log_out.asp"><img src="/images/btn_logout.gif" width="64" height="19" border="0"></A></td>
      </tr>
    </table>
	<% else %><table width="700"  border="0" align="right" cellpadding="0" cellspacing="0" height="60">
      <tr style="padding-top:10">
        <td height="24">&nbsp;</td>
      </tr>
    </table>
	<% end if %></td>
  </tr>
  <tr>
    <td height="24">&nbsp;</td>
  </tr>
  <tr>
    <td height="17"  align="center"><table width="1030" border="0" cellpadding="0" cellspacing="0" >
    <tr>
		<td class='title'><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle" > <%=rs("title")%> (<%=getside(rs("side"))%>)  모니터링 </td>
    </tr>
    </table></td>
  </tr>
  <tr>
    <td height="600"  align='center'  valign="top">
	<!-- -->
<table width="1024" align="center" style="margin-top:10px;">
	<tr>
		<td  colspan='8' height='25' align='right'><a href="list_monitor.asp?cyear=<%=pcyear%>&cmonth=<%=pcmonth%>&cmbcustcode=<%=pcustcode%>&cmbteamcode=<%=pteamcode%>"> 목록 </a> |  <% If Not IsNull(rs("cdate")) Then %><a href='#' onclick="getmonitor('u'); return false;">   수정</a> | <a href='#' onclick="if (confirm('등록된 모니터링 정보를 삭제하시겠습니까?')) {getmonitor('d')}; return false;"> 삭제 </a> | <a href="#" onclick="getmonitor('c'); return false;"> 등록 </a><% Else %><a href="#" onclick="getmonitor('c'); return false;"> 등록 </a><% End if%></td>
	</tr>
	<tr>
		<th class="title">광고주</th>
		<td width="156" class="context" colspan='7'><strong><%=getcustname(rs("custcode"))%></strong></td>
	</tr>
	<tr>
		<th class="title">계약기간</th>
		<td colspan="3" class="context"><%=rs("startdate")%>&nbsp;&nbsp; ~ &nbsp;&nbsp;<%=rs("enddate")%></td>
		<th class="title">매체규격</th>
		<td class="context" colspan='3'><%=rs("standard")%> </td>
	</tr>
	<tr>
		<th class="title">매체위치</th>
		<td class="context" colspan='3'>[<%=rs("region")%>] <%=rs("locate")%></td>
		<th class="title">매체재질</th>
		<td class="context" colspan='3'><%=rs("quality")%> </td>
	</tr>
	<tr>
		<th class="title">검수일자</th>
		<td width="156" class="context"><%=rs("cdate")%></td>
		<th class="title">검수횟수</th>
		<td width="156" class="context"><%=rs("num")%> <%If rs("num") <>"" Then response.write " (회차)"%></td>
		<th class="title">검수상태</th>
		<td width="156" class="context"><%=getmonitorstatus(rs("status"))%></td>
		<th class="title">검수자명</th>
		<td width="156" class="context"><%=rs("cname")%></td>
	</tr>
	<tr>
		<th class="title"  style="height:54px;">검수내용</th>
		<td class="context" colspan='7'><%=rs("comment")%></td>
	</tr>
</table>
<table width="1024" align="center" style="margin-top:10px;" >
	<tr>
		<td align='left'></td>
		<td align='right'></td>
	</tr>
</table>
<table width="1048" align="center" style="margin-top:5px;" >
	<tr valign='top'>
		<td width='24' valign='middle'><% Call getMonitorData(pmdidx, pside, pcyear, pcmonth, pnum, pcustcode, pteamcode, "P") %></td>
		<td width='250' ><%=getimage(rs("img01"))%></td>
		<td width='250'><%=getimage(rs("img02"))%></td>
		<td width='250'><%=getimage(rs("img03"))%></td>
		<td width='250'><%=getimage(rs("img04"))%></td>
		<td width='24' valign='middle'><% Call getMonitorData(pmdidx, pside, pcyear, pcmonth, pnum, pcustcode, pteamcode, "N") %></td>
	</tr>
</table>
<p>
<img id='preimg' width=0 height=0 align='center' style='margin-top:10px;' />

	<!--  -->
	</td>
  </tr>
  <tr>
    <td height="24">&nbsp; </td>
  </tr>
  <tr>
    <td height="24"><img src="/images/pop_bottom.gif" width="1240" height="71" align="absmiddle"></td>
  </tr>
</table>
</form>
</body>
</html>
<%
	Function getimage(photo)
		If IsNull(photo) Or photo = "" Then
			getimage = "<img src='/images/noimage.gif' width='240' height='180'  class='noimage' id='"&photo&"' >"
		Else
			getimage = "<img src='/pds/monitor/"&photo&"' width=240' height='180' class='photo' id='"&photo&"' onmouseover='preview();'>"
		End If
	End Function

	Sub getMonitorData(mdidx, side, cyear, cmonth, num, custcode, teamcode, direct)
		If UCase(direct) = "P" Then
			Dim sql : sql = "select num from wb_contact_monitor where mdidx=? and side=? and cyear=? and cmonth=? and num=(select max(num) from wb_contact_monitor where   mdidx=? and side=? and cyear=? and cmonth=? and num<?)"
		Else
			sql = "select num from wb_contact_monitor where mdidx=? and side=? and cyear=? and cmonth=? and num = (select min(num) from wb_contact_monitor where mdidx=? and side=? and cyear=? and cmonth=? and num>?)"
		End If
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		cmd.parameters.append	cmd.createparameter("mdidx", adinteger, adParaminPut)
		cmd.parameters.append	cmd.createparameter("side", adchar, adParaminPut, 1)
		cmd.parameters.append	cmd.createparameter("cyear", adchar, adParaminPut, 4)
		cmd.parameters.append	cmd.createparameter("cmonth", adchar, adParaminPut, 2)
		cmd.parameters.append	cmd.createparameter("mdidx2", adinteger, adParaminPut)
		cmd.parameters.append	cmd.createparameter("side2", adchar, adParaminPut, 1)
		cmd.parameters.append	cmd.createparameter("cyear2", adchar, adParaminPut, 4)
		cmd.parameters.append	cmd.createparameter("cmonth2", adchar, adParaminPut, 2)
		cmd.parameters.append	cmd.createparameter("num", adUnsignedTinyInt, adParaminPut)
		cmd.parameters("mdidx").value = mdidx
		cmd.parameters("side").value = side
		cmd.parameters("cyear").value = cyear
		cmd.parameters("cmonth").value = cmonth
		cmd.parameters("mdidx2").value = mdidx
		cmd.parameters("side2").value = side
		cmd.parameters("cyear2").value = cyear
		cmd.parameters("cmonth2").value = cmonth
		cmd.parameters("num").value = num
		Dim rs : Set rs = cmd.execute

		If Not rs.eof Then
			If UCase(direct) = "P" Then
				Response.write "<a href='detail_monitor.asp?mdidx="&mdidx&"&side="&side&"&cyear="&cyear&"&cmonth="&cmonth&"&num="&rs(0)&"&custcode="&custcode&"&teamcode="&teamcode&"'><img src='/images/btn_prev.gif' width='24' height='42' alt='"&rs(0)&" 회차'></a>"
			Else
				Response.write "<a href='detail_monitor.asp?mdidx="&mdidx&"&side="&side&"&cyear="&cyear&"&cmonth="&cmonth&"&num="&rs(0)&"&custcode="&custcode&"&teamcode="&teamcode&"'><img src='/images/btn_next.gif' width='24' height='42' alt='"&rs(0)&" 회차'></a>"
			End If
		End If
	End Sub
%>