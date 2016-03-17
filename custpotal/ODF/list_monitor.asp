<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<%
	If request.cookies("userid") = "" Then response.redirect "/"
	Dim pcustcode : pcustcode = request("cmbcustcode")
	Dim pteamcode : pteamcode = request("cmbteamcode")
	Dim cyear : cyear = request("cyear")
	Dim cmonth : cmonth = request("cmonth")
	If cyear = "" Then cyear = Year(date)
	If cmonth = "" Then cmonth = Month(date)
	If Len(cmonth) = 1 Then cmonth = "0"&cmonth
	dim sdate : sdate = dateserial(cyear, cmonth, "01")
	dim edate : edate = dateadd("d", -1, dateadd("m", 1,  sdate))

	Dim sql : sql = "select a.contidx, a.custcode , a.title, a.startdate, a.enddate from wb_contact_mst a inner join sc_cust_dtl b on a.custcode=b.custcode where  a.custcode like '"&pteamcode&"%' and b.highcustcode like '"&pcustcode&"%' and  a.startdate <= '"&edate&"' and a.enddate >= '" & sdate &"' order by a.contidx desc"

	Dim rs : Set rs = server.CreateObject("adodb.recordset")
	rs.activeconnection = application("connectionstring")
	rs.cursorLocation = aduseclient
	rs.cursortype = adopenstatic
	rs.locktype = adlockoptimistic
	rs.source = sql
	rs.open

	Dim totalrecord : totalrecord = rs.recordcount

	Dim contidx : Set contidx = rs(0)
	Dim custcode : Set custcode = rs(1)
	Dim title : Set title = rs(2)
	Dim startdate : Set startdate = rs(3)
	Dim enddate : Set enddate = rs(4)

	'response.write sql


	sql = "select b.contidx, region, locate, mdidx  from wb_contact_mst a inner join wb_contact_md b on a.contidx=b.contidx inner join sc_cust_dtl c on a.custcode=c.custcode where  a.custcode like '"&pteamcode&"%' and c.highcustcode like '"&pcustcode&"%' and  a.startdate <= '"&edate&"' and a.enddate >= '" & sdate &"' order by a.contidx desc"

'	response.write sql

	Dim rs2 : Set rs2 = server.CreateObject("adodb.recordset")
	rs2.activeconnection = application("connectionstring")
	rs2.cursorlocation = aduseclient
	rs2.cursortype = adopenstatic
	rs2.locktype = adLockOptimistic
	rs2.source = sql
	rs2.open

	If Not rs2.eof Then
		Dim region : Set region = rs2(1)
		Dim locate : Set locate = rs2(2)
		Dim mdidx : Set mdidx = rs2(3)
	End If

	sql = " select c.mdidx, c.side, d.cdate, d.num from wb_contact_mst a  inner join wb_contact_md b on a.contidx=b.contidx inner join vw_contact_md_dtl c on b.mdidx=c.mdidx left outer join  (select mdidx, cyear, cmonth, max(num) num, side, max(cdate) cdate from wb_contact_monitor group by mdidx, cyear, cmonth, side, cyear, cmonth)as d on c.mdidx=d.mdidx and c.side=d.side and d.cyear='"&cyear&"' and d.cmonth='"&cmonth&"' order by case when c.side <> 'L' then ' ' + c.side else c.side end desc"

'	response.write sql

	Dim rs3 : Set rs3 = server.CreateObject("adodb.recordset")
	rs3.activeconnection = application("connectionstring")
	rs3.cursorlocation = aduseclient
	rs3.cursortype = adopenstatic
	rs3.locktype = adLockOptimistic
	rs3.source = sql
	rs3.open

	If Not rs3.eof Then
		Dim side : Set side = rs3(1)
		Dim checkdate : Set checkdate = rs3(2)
		Dim num : Set num = rs3(3)
	End If

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
		function getcustcombo() {
			// 광고주 콤보 박스 가져오기
			var scope = null;
			var custcode = null;
			var params = "scope="+scope+"&custcode="+custcode;
			sendRequest("/inc/getcustcombo.asp", params, _getcustcombo, "GET");
		}

		function _getcustcombo() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
						var custcode = document.getElementById("custcode");
						custcode.innerHTML = xmlreq.responseText ;
//						getteamcombo();
				}
			}
		}

		function getteamcombo() {
			// 운영팀 콤보 박스 가져오기
			var custcode = document.getElementById("cmbcustcode").value;
			var teamcode = null;
			var params = "custcode="+custcode+"&teamcode="+teamcode;
			sendRequest("/inc/getteamcombo.asp", params, _getteamcombo, "GET");
		}

		function _getteamcombo() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
						var teamcode = document.getElementById("teamcode");
						teamcode.innerHTML = xmlreq.responseText ;
				}
			}
		}

		window.onload = function () {
			_sendRequest("/inc/getcustcombo.asp", "custcode=<%=pcustcode%>", _getcustcombo, "GET");
//			_sendRequest("/inc/getteamcombo.asp", "custcode=<%=pcustcode%>&teamcode=<%=pteamcode%>", _getteamcombo, "GET");
//			document.getElementById("cmbcustcode").attachEvent("onchange", getteamcombo);
		}
//-->
</script>
</head>
<!-- <body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"   oncontextmenu="return false">
 --><body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"   >
<form action='list_monitor.asp' method='POST'>
<table width="1240" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td height="24" background="/images/pop_top.gif" valign="top" >
	<table width="700"  border="0" align="right" cellpadding="0" cellspacing="0" height="60">
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
    </table></td>
  </tr>
  <tr>
    <td height="24">&nbsp;</td>
  </tr>
<!--
  <tr>
    <td height="17"  align="center">
	<table width="1030" border="0" cellpadding="0" cellspacing="0" >
    <tr>
		<td class='title'><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle" > <%'=request.cookies("custname")%>  옥외광고 모니터링 </td>
    </tr>
    </table>
	</td>
  </tr>
  <tr>
    <td height="27">&nbsp;</td>
  </tr>
 -->
  <tr>
    <td valign="top" class="bdpdd" align="center">

			<table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13" valign="top" ><img src="/images/bg_search_left.gif" width="13" height="35" ></td>
                  <td background="/images/bg_search.gif"> <img src="/images/icon_dot_search.gif" width="4" height="3" align="absmiddle"> 검색년월
				  <%call getyear(cyear)%> <%call getmonth(cmonth)%> &nbsp;    <span id="custcode">광고주 검색</span>  <!--<span id="teamcode">운영팀 검색</span>  <span id='medcode'> 매체사 검색 </span> <span id='empcode'> 담당자 검색 </span> --> <input type="image" src="/images/btn_search.gif" width="39" height="20" align="absmiddle" border=0></td>
				</td>
                  <td width="13" ><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table>
			<table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
			  <tr>
				<td height="35" align='right' style='padding-right:10px;'> <img src='/images/m_photo.gif' width='16' height='15' >  모니터링 관리 </td>
			  </tr>
            </table>
	</td>
  </tr>
  <tr>
    <td height="600"  align='center'  valign="top">
	<!-- -->
	  <table border="0"width="1030" cellpadding="0" cellspacing="0" bordercolor="#8D652B" id="contact">
	  <thead>
			<tr height='30' align='center'>
				<th  width="20" class="hd left">No</th >
				<th  width="250" class="hd center">매체명</th >
				<th  width="80" class="hd center">시작일자</th >
				<th  width="80" class="hd center">종료일자</th >
				<th  width="290" class="hd center">매체위치</th >
				<th  width="40" class="hd center">면</th >
				<th  width="30" class="hd center">&nbsp;</th >
				<th  width="80" class="hd center">최종검수일</th >
				<th  width="30" class="hd center">회차</th >
				<th  width="90" class="hd right">광고주</th >
			</tr>
		</thead>
		<tbody id='tbody'>
		<%
				Do Until rs.eof
		%>
			<tr height='32'>
				<td  class="hd none"  width="20" style='text-align:left;padding-top:9px;padding-left:5px;vertical-align:top;'><%=totalrecord%></td>
				<td  class="hd none" width="250" title='<%=title%>' style='text-align:left;padding-top:9px;padding-left:5px;vertical-align:top;' > <%=cutTitle(title, 42)%></a></td>
				<td  class="hd none" width="80" style='text-align:center;padding-top:9px;vertical-align:top;' ><%=startdate%></td>
				<td  class="hd none" width="80" style='text-align:center;padding-top:9px;vertical-align:top;'><%=enddate%></td>
				<td  class="hd none" colspan='5'><table  width='485' border=0>
				<%
					rs2.Filter = "contidx="&contidx
					rs2.sort = "mdidx desc"
					Do Until rs2.eof
				%>
					<tr height='32'>
						<td  width="290" title='<%=locate%>'  style='text-align:left;padding-top:9px;padding-left:5px;vertical-align:top;' > [<%=region%>] <%=locate%></td>
						<td width="180" colspan='4'><table  border=0 style="table-layout:fixed;">
							<%
								rs3.Filter = "mdidx="& mdidx
								rs3.sort = "side desc"
								Do Until rs3.eof
							%>
								<tr height='32'>
								<td  width="40" style='text-align:center;'><%=getside(side)%></td>
								<td  width="30" style='text-align:center;'><a href="detail_monitor.asp?custcode=<%=pcustcode%>&pteamcode=<%=pteamcode%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>&mdidx=<%=mdidx%>&side=<%=side%>&num=<%=num%>"><img src='/images/m_photo.gif' width='16' height='15' ></a></td>
								<td  width="80" style='text-align:center;'><%=checkdate%></td>
								<td  width="30" style='text-align:center;'><%=num%></td>
								</tr>
							<%
									rs3.movenext
								Loop
								rs3.Filter = ""
							%>
						</table></td>
					</tr>
				<%
					rs2.movenext
					Loop
				%>
			</table></td>
			<td  class="hd none" width="90" style='text-align:left;padding-top:9px;padding-left:5px;vertical-align:top;' ><%=getcustname(custcode)%></td>
			</tr>
			<%
						totalrecord = totalrecord - 1
						rs.movenext
					Loop
			%>
		</tbody>
        </table>
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