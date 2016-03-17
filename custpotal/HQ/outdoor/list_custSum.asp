<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<%
	Dim pcustcode : pcustcode = request("cmbcustcode")
	Dim pteamcode : pteamcode = request("cmbteamcode")
	Dim cyear : cyear = request("cyear")
	If cyear = "" Then cyear = Year(date)
	Dim sdate : sdate = DateSerial(cyear, "01", "01")
	Dim edate : edate = DateSerial(cyear, "12", "31")

	Dim sql : sql = "select d.highcustcode, a.custcode"
	sql = sql & ",sum(case when c.cmonth = '01' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '02' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '03' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '04' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '05' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '06' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '07' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '08' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '09' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '10' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '11' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(case when c.cmonth = '12' then isnull(c.monthly,0) else 0 end ) "
	sql = sql & ",sum(isnull(c.monthly,0)) "
	sql = sql & "from wb_contact_mst a inner join wb_contact_md b on a.contidx=b.contidx "
	sql = sql & "inner join wb_contact_exe c on c.mdidx=b.mdidx and cyear='"&cyear&"' "
	sql = sql & "inner join sc_cust_dtl d on a.custcode=d.custcode "
	sql = sql & "inner join wb_contact_trans e  "
	sql = sql & " on a.contidx=e.contidx and b.medcode=e.medcode and c.cyear = e.cyear and c.cmonth = e.cmonth and e.ishold in('Y','N') "
	sql = sql & "where d.highcustcode like '" & pcustcode & "%' and a.custcode like '" & pteamcode & "%' "
	sql = sql & "group by d.highcustcode, a.custcode with rollup"

'
'	response.write sql

	Dim rs : Set rs = server.CreateObject("adodb.recordset")
	rs.activeconnection = application("connectionstring")
	rs.cursorLocation = adUseClient
	rs.cursorType = adOpenStatic
	rs.lockType = adLockOptimistic
	rs.source = sql
	rs.open

	If Not rs.eof Then
		Dim custcode : Set custcode = rs(1)
		Dim jan : Set jan = rs(2)
		Dim feb : Set feb = rs(3)
		Dim mar : Set mar = rs(4)
		Dim apr : Set apr = rs(5)
		Dim may : Set may = rs(6)
		Dim jun : Set jun = rs(7)
		Dim jul : Set jul = rs(8)
		Dim aug : Set aug = rs(9)
		Dim sep : Set sep = rs(10)
		Dim oct_ : Set oct_ = rs(11)
		Dim nov : Set nov = rs(12)
		Dim dec : Set dec = rs(13)
		Dim sum : Set sum = rs(14)
	End If


	sql = "select distinct b.highcustcode, c.custname from wb_contact_mst a inner join sc_cust_dtl b on a.custcode=b.custcode inner join sc_cust_hdr c on b.highcustcode=c.highcustcode where a.custcode like '"&pteamcode&"%' and b.highcustcode like '"&pcustcode&"%' and a.startdate <= '"&edate&"' and a.enddate >= '"&sdate&"' "
'	response.write sql
	Dim cmd : set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adcmdText
	Dim rs2 : Set rs2 = cmd.execute
	If Not rs2.eof Then
		Dim highcustcode : Set highcustcode = rs2(0)
		Dim custname : Set custname = rs2(1)
	End If
	Set cmd = Nothing

	Dim s01 : s01 = 0
	Dim s02 : s02 = 0
	Dim s03 : s03 = 0
	Dim s04 : s04 = 0
	Dim s05 : s05 = 0
	Dim s06 : s06 = 0
	Dim s07 : s07 = 0
	Dim s08 : s08 = 0
	Dim s09 : s09 = 0
	Dim s10 : s10 = 0
	Dim s11 : s11 = 0
	Dim s12 : s12 = 0
	Dim total : total = 0
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
						getteamcombo();
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
						var teamcomboClick = document.getElementById("cmbteamcode");
						document.getElementById("teamcode").style.width = 100;
				}
			}
		}

		function getexcel() {
			// 엑셀전환
			var custcode = document.getElementById("cmbcustcode").value;
			var teamcode = document.getElementById("cmbteamcode").value;
			var cyear = document.getElementById("cyear").value;

			location.href = "/hq/outdoor/excel/xls_custSum.asp?custcode="+custcode+"&teamcode="+teamcode+"&cyear="+cyear;
		}

		window.onload = function () {
			_sendRequest("/inc/getcustcombo.asp", "custcode=<%=pcustcode%>", _getcustcombo, "GET");
			_sendRequest("/inc/getteamcombo.asp", "custcode=<%=pcustcode%>&teamcode=<%=pteamcode%>", _getteamcombo, "GET");
			document.getElementById("cmbcustcode").attachEvent("onchange", getteamcombo);
		}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  >
<form action="list_custsum.asp" method='post'>
<INPUT TYPE="hidden" NAME="menunum" value="<%=request("menunum")%>">
<!--#include virtual="/hq/top.asp" -->
  <table width="1240" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="Table_01">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/hq/left_outdoor_menu.asp"--></td>
      <td height="65" valign="top"><img src="/images/middle_navigater_outdoor.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td align="left" valign="top" colspan='2'>
	  <table width="1030" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="19" colspan='2'>&nbsp;</td>
          </tr>
          <tr>
            <td height="17" colspan='2'><TABLE  width="100%">
            <TR>
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="boldtitle"> 광고주별 집행현황 </span></TD>
				<TD width="50%" align="right"><span class="navigator" > 옥외광고 &gt;  옥외광고현황 &gt; 광고주별 집행현황 </span></TD>
            </TR>
            </TABLE>
			</td>
          </tr>
          <tr>
            <td height="15" colspan='2'>&nbsp;</td>
          </tr>
          <tr>
            <td colspan='2'>
			<table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13" valign="top" ><img src="/images/bg_search_left.gif" width="13" height="35" ></td>
                  <td background="/images/bg_search.gif"> <img src="/images/icon_dot_search.gif" width="4" height="3" align="absmiddle"> 검색년도
				  <%call getyear(cyear)%>     <span id="custcode">광고주 검색</span>  <span id="teamcode">운영팀 검색</span> <input type="image" src="/images/btn_search.gif" width="39" height="20" align="absmiddle" border=0></td>
                  <td  align="right" background="/images/bg_search.gif" ><!-- <img src="/images/btn_new.gif" width="30" height="30" align="absmiddle" border="0" alt="신규 계약 등록"></a><a href="#" onclick="fncPrint();"><img src="/images/btn_print.gif" width="30" height="30" align="absmiddle" border="0" alt="관리보고서 출력" hspace=2></a><a href="#" onclick="get_excel_sheet();"><img src="/images/btn_xls.gif" width="30" height="30" align="absmiddle" border="0" alt="엑셀 변환"></a> -->
				</td>
                  <td width="13" ><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="35" >&nbsp;(단위:천원)</td>
			<td align='right'><a href="#" onclick="getexcel(); return false;"><img src='/images/icon_xls.gif' width='17' height='16'> 엑셀 </a>  </td>
          </tr>
          <tr>
            <td valign="top" height="516" colspan='2'>

				  <table border="0"width="1650" cellpadding="0" cellspacing="0" bordercolor="#8D652B" id="contact">
					<tr height='30' align='center'>
						<th class="hd left" width="150">광고주</th>
						<th class="hd center" width="200">운영팀</th>
						<th class="hd center"width="100" >01월</th>
						<th class="hd center"width="100" >02월</th>
						<th class="hd center"width="100" >03월</th>
						<th class="hd center"width="100" >04월</th>
						<th class="hd center"width="100" >05월</th>
						<th class="hd center"width="100" >06월</th>
						<th class="hd center"width="100" >07월</th>
						<th class="hd center"width="100" >08월</th>
						<th class="hd center"width="100" >09월</th>
						<th class="hd center"width="100" >10월</th>
						<th class="hd center"width="100" >11월</th>
						<th class="hd center"width="100" >12월</th>
						<th class="hd right" width="100">합계</th>
					</tr>
					<%
						If Not rs2.eof Then
							Do Until rs2.eof
							rs.Filter = "highcustcode='"& highcustcode &"' "

							If rs.recordcount <> 0 Then
					%>
						<tr height='32'>
							<td  width='150' class="hd none" style='padding-left:3px;padding-top:9px;' valign='top'><%=custname%></td>
							<td  class="hd none" colspan='14' ><table  width='1400' border=0 style="table-layout:fixed;">
							<%
								Do Until rs.eof
							%>
								<tr height='30' <% If IsNull(custcode) Then response.write "bgcolor=#ececec" End If %>>
									<td  width='200'  style='padding-left:3px;'><%If IsNull(custcode) Then response.write "소계" Else response.write getteamname(custcode) %></td>
									<td  width="100"  style='padding-right:5px;text-align:right;'><%=FormatNumber(jan,0)%></td>
									<td  width="100"  style='padding-right:5px;text-align:right;'><%=FormatNumber(feb,0)%></td>
									<td  width="100"  style='padding-right:5px;text-align:right;'><%=FormatNumber(mar,0)%></td>
									<td  width="100"  style='padding-right:5px;text-align:right;'><%=FormatNumber(apr,0)%></td>
									<td  width="100"  style='padding-right:5px;text-align:right;'><%=FormatNumber(may,0)%></td>
									<td  width="100"  style='padding-right:5px;text-align:right;'><%=FormatNumber(jun,0)%></td>
									<td  width="100"  style='padding-right:5px;text-align:right;'><%=FormatNumber(jul,0)%></td>
									<td  width="100"  style='padding-right:5px;text-align:right;'><%=FormatNumber(aug,0)%></td>
									<td  width="100"  style='padding-right:5px;text-align:right;'><%=FormatNumber(sep,0)%></td>
									<td  width="100"  style='padding-right:5px;text-align:right;'><%=FormatNumber(oct_,0)%></td>
									<td  width="100"  style='padding-right:5px;text-align:right;'><%=FormatNumber(nov,0)%></td>
									<td  width="100"  style='padding-right:5px;text-align:right;'><%=FormatNumber(dec,0)%></td>
									<td  width='100' style='padding-right:10px;text-align:right;'><%=FormatNumber(sum,0)%></td>
								</tr>
							<%
								If IsNull(custcode) Then
									s01 = s01 + jan
									s02 = s02 + feb
									s03 = s03 + mar
									s04 = s04 + apr
									s05 = s05 + may
									s06 = s06 + jun
									s07 = s07 + jul
									s08 = s08 + aug
									s09 = s09 + sep
									s10 = s10 + oct_
									s11 = s11 + nov
									s12 = s12 + dec
									total = total + sum
								End If
								rs.movenext
								Loop

							%>
						</table></td>
						</tr>
						<%
							rs.movelast
							End If
								rs2.movenext
							Loop
							End If
						%>
					<tr height='32' align='center'>
						<th class="hd left" colspan=2 > 총합계</th>
						<th class="hd center"width="100"  align='right'><%=FormatNumber(s01,0)%>&nbsp;</th>
						<th class="hd center"width="100"  align='right'><%=FormatNumber(s02,0)%>&nbsp;</th>
						<th class="hd center"width="100"  align='right'><%=FormatNumber(s03,0)%>&nbsp;</th>
						<th class="hd center"width="100"  align='right'><%=FormatNumber(s04,0)%>&nbsp;</th>
						<th class="hd center"width="100"  align='right'><%=FormatNumber(s05,0)%>&nbsp;</th>
						<th class="hd center"width="100"  align='right'><%=FormatNumber(s06,0)%>&nbsp;</th>
						<th class="hd center"width="100"  align='right'><%=FormatNumber(s07,0)%>&nbsp;</th>
						<th class="hd center"width="100"  align='right'><%=FormatNumber(s08,0)%>&nbsp;</th>
						<th class="hd center"width="100"  align='right'><%=FormatNumber(s09,0)%>&nbsp;</th>
						<th class="hd center"width="100"  align='right'><%=FormatNumber(s10,0)%>&nbsp;</th>
						<th class="hd center"width="100"  align='right'><%=FormatNumber(s11,0)%>&nbsp;</th>
						<th class="hd center"width="100"  align='right'><%=FormatNumber(s12,0)%>&nbsp;</th>
						<th class="hd right" width="100" align='right'><%=FormatNumber(total,0)%></th>
					</tr>
              </table>
			  <p/>
			  <div id="debugConsole"></div>
			  </td>
          </tr>
      </table></td>
    </tr>
  </table>
</form>
</body>
</html>
<!--#include virtual="/bottom.asp" -->