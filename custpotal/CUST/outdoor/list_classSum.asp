<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<%
	Dim userid : userid = request.cookies("userid")
	Dim pcustcode : pcustcode = request("cmbcustcode")
	Dim pteamcode : pteamcode = request("cmbteamcode")
	Dim cyear : cyear = request("cyear")
	If cyear = "" Then cyear = Year(date)


	Dim sql
	Dim chkcount
	chkcount= 1
	Dim Custcodesql
	Dim Custcoderecord
	Dim Timcodesql
	Dim Timcoderecord
	Dim objrs_1
	Dim objrs
	Dim rs


	if pcustcode = "" or pcustcode = null then
		'=========================================================================================
		Custcodesql = "select clientcode from wb_account_cust where userid ='" & userid & "' "

		Set objrs_1 = server.CreateObject("adodb.recordset")
		objrs_1.activeconnection = application("connectionstring")
		objrs_1.cursorLocation = aduseclient
		objrs_1.cursortype = adopenstatic
		objrs_1.locktype = adlockoptimistic
		objrs_1.source = Custcodesql
		objrs_1.open

		Custcoderecord = objrs_1.recordcount
		'=========================================================================================



		if not objrs_1.eof then
			do until objrs_1.eof
				'=========================================================================================
				Timcodesql = "select timcode from wb_account_tim where userid ='" & userid & "' and clientcode = '" & objrs_1("clientcode") &"'"

				Set objrs = server.CreateObject("adodb.recordset")
				objrs.activeconnection = application("connectionstring")
				objrs.cursorLocation = aduseclient
				objrs.cursortype = adopenstatic
				objrs.locktype = adlockoptimistic
				objrs.source = Timcodesql
				objrs.open

				Timcoderecord = objrs.recordcount
				'=========================================================================================

				if chkcount > 1 then
					sql = sql  & " Union all "
				end if

				sql = sql & " select c.highclasscode, "
				sql = sql & " isnull(c.middleclassname, '소계') "
				sql = sql & " ,sum(case when b.cmonth='01' then isnull(b.monthly,0) else 0 end) "
				sql = sql & " ,sum(case when b.cmonth='02' then isnull(b.monthly,0) else 0 end) "
				sql = sql & " ,sum(case when b.cmonth='03' then isnull(b.monthly,0) else 0 end) "
				sql = sql & " ,sum(case when b.cmonth='04' then isnull(b.monthly,0) else 0 end) "
				sql = sql & " ,sum(case when b.cmonth='05' then isnull(b.monthly,0) else 0 end) "
				sql = sql & " ,sum(case when b.cmonth='06' then isnull(b.monthly,0) else 0 end) "
				sql = sql & " ,sum(case when b.cmonth='07' then isnull(b.monthly,0) else 0 end) "
				sql = sql & " ,sum(case when b.cmonth='08' then isnull(b.monthly,0) else 0 end) "
				sql = sql & " ,sum(case when b.cmonth='09' then isnull(b.monthly,0) else 0 end) "
				sql = sql & " ,sum(case when b.cmonth='10' then isnull(b.monthly,0) else 0 end) "
				sql = sql & " ,sum(case when b.cmonth='11' then isnull(b.monthly,0) else 0 end) "
				sql = sql & " ,sum(case when b.cmonth='12' then isnull(b.monthly,0) else 0 end) "
				sql = sql & " ,sum(isnull(b.monthly,0)) "
				sql = sql & " from wb_contact_md a "
				sql = sql & " inner join wb_contact_exe b on a.mdidx=b.mdidx and cyear='"&cyear&"' "
				sql = sql & " inner join vw_medium_class c on a.categoryidx=c.catcode "
				sql = sql & " inner join wb_contact_mst d on a.contidx=d.contidx "
				sql = sql  & " left outer join sc_cust_dtl e on e.custcode = d.custcode "
				sql = sql  & " inner  join wb_account_cust n on e.highcustcode  = n.clientcode and n.userid='"&userid&"' and n.clientcode =  '" & objrs_1("clientcode") &"' "
				If Timcoderecord > 0 then
					sql = sql  & " inner  join wb_account_tim t on d.custcode = t.timcode and t.userid='"&userid&"' and t.clientcode = '" & objrs_1("clientcode") &"' "
				End If
				sql = sql & " inner join wb_contact_trans f on f.cyear='"&cyear&"'  and a.contidx=f.contidx and b.cmonth = f.cmonth and a.medcode = f.medcode and (f.isHold='N' or f.isHold='Y')"
				sql = sql & " group by c.highclasscode, c.middleclassname with rollup "
				sql = sql & " having highclasscode is not null "


				chkcount = chkcount +1
				objrs_1.movenext
			Loop
			sql = sql & " order by highclasscode "
		end if

else

	'=========================================================================================
	Timcodesql = "select timcode from wb_account_tim where userid ='" & userid & "' and clientcode ='" & pcustcode & "'"

	Set objrs = server.CreateObject("adodb.recordset")
	objrs.activeconnection = application("connectionstring")
	objrs.cursorLocation = aduseclient
	objrs.cursortype = adopenstatic
	objrs.locktype = adlockoptimistic
	objrs.source = Timcodesql
	objrs.open

	Timcoderecord = objrs.recordcount
	'=========================================================================================


	sql = sql & " select c.highclasscode, "
	sql = sql & " isnull(c.middleclassname, '소계') "
	sql = sql & " ,sum(case when b.cmonth='01' then isnull(b.monthly,0) else 0 end) "
	sql = sql & " ,sum(case when b.cmonth='02' then isnull(b.monthly,0) else 0 end) "
	sql = sql & " ,sum(case when b.cmonth='03' then isnull(b.monthly,0) else 0 end) "
	sql = sql & " ,sum(case when b.cmonth='04' then isnull(b.monthly,0) else 0 end) "
	sql = sql & " ,sum(case when b.cmonth='05' then isnull(b.monthly,0) else 0 end) "
	sql = sql & " ,sum(case when b.cmonth='06' then isnull(b.monthly,0) else 0 end) "
	sql = sql & " ,sum(case when b.cmonth='07' then isnull(b.monthly,0) else 0 end) "
	sql = sql & " ,sum(case when b.cmonth='08' then isnull(b.monthly,0) else 0 end) "
	sql = sql & " ,sum(case when b.cmonth='09' then isnull(b.monthly,0) else 0 end) "
	sql = sql & " ,sum(case when b.cmonth='10' then isnull(b.monthly,0) else 0 end) "
	sql = sql & " ,sum(case when b.cmonth='11' then isnull(b.monthly,0) else 0 end) "
	sql = sql & " ,sum(case when b.cmonth='12' then isnull(b.monthly,0) else 0 end) "
	sql = sql & " ,sum(isnull(b.monthly,0)) "
	sql = sql & " from wb_contact_md a "
	sql = sql & " inner join wb_contact_exe b on a.mdidx=b.mdidx and cyear='"&cyear&"' "
	sql = sql & " inner join vw_medium_class c on a.categoryidx=c.catcode "
	sql = sql & " inner join wb_contact_mst d on a.contidx=d.contidx "
	sql = sql  & " left outer join sc_cust_dtl e on e.custcode = d.custcode "
	sql = sql  & " inner  join wb_account_cust n on e.highcustcode  = n.clientcode and n.userid='"&userid&"' "
	If Timcoderecord > 0 then
		sql = sql  & " inner  join wb_account_tim t on d.custcode = t.timcode and t.userid='"&userid&"' "
	End If
	sql = sql & " inner join wb_contact_trans f on f.cyear='"&cyear&"'  and a.contidx=f.contidx and b.cmonth = f.cmonth and a.medcode = f.medcode and (f.isHold='N' or f.isHold='Y')"
	sql = sql & " where d.custcode like '"&pteamcode&"%' and e.highcustcode like '"&pcustcode&"%' "
	sql = sql & " group by c.highclasscode, c.middleclassname with rollup "
	sql = sql & " having highclasscode is not null "
	sql = sql & " order by highclasscode "

end if



	 Set rs = server.CreateObject("adodb.recordset")
	rs.activeconnection = application("connectionstring")
	rs.cursorLocation = adUseClient
	rs.cursorType = adOpenStatic
	rs.lockType = adLockOptimistic
	rs.source = sql
	rs.open

	If Not rs.eof Then
		Dim middleclassname : Set middleclassname = rs(1)
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


	sql = "select categoryidx, categoryname from wb_category where categorylvl is null"
	Dim cmd : set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adcmdText
	Dim rs2 : Set rs2 = cmd.execute
	If Not rs2.eof Then
		Dim highclasscode : Set highclasscode = rs2(0)
		Dim highclassname : Set highclassname = rs2(1)
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
<link href="/cust/outdoor/style.css" rel="stylesheet" type="text/css">
<script type='text/javascript' src='/js/ajax.js'></script>
<script type='text/javascript' src='/js/script.js'></script>
<script type="text/javascript">
<!--
		function getcustcombo() {
			// 광고주 콤보 박스 가져오기
			var scope = null;
			var custcode = null;
			var params = "scope="+scope+"&custcode="+custcode;
			sendRequest("/inc/getcustcombo_cust.asp", params, _getcustcombo, "GET");
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
			sendRequest("/inc/getteamcombo_cust.asp", params, _getteamcombo, "GET");
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

			location.href = "/cust/outdoor/excel/xls_classSum.asp?custcode="+custcode+"&teamcode="+teamcode+"&cyear="+cyear;
		}

		window.onload = function () {
			_sendRequest("/inc/getcustcombo_cust.asp", "custcode=<%=pcustcode%>", _getcustcombo, "GET");
			_sendRequest("/inc/getteamcombo_cust.asp", "custcode=<%=pcustcode%>&teamcode=<%=pteamcode%>", _getteamcombo, "GET");
			document.getElementById("cmbcustcode").attachEvent("onchange", getteamcombo);
		}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  >
<form action="list_classsum.asp" method='post'>
<INPUT TYPE="hidden" NAME="menunum" value="<%=request("menunum")%>">
<!--#include virtual="/cust/top.asp" -->
  <table width="1240" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="Table_01">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/cust/left_outdoor_menu.asp"--></td>
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
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="boldtitle"> 매체분류 집행현황 </span></TD>
				<TD width="50%" align="right"><span class="navigator" > 옥외광고 &gt;  옥외광고현황 &gt; 매체분류 집행현황 </span></TD>
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
				  <%call getyear(cyear)%>     <span id="custcode">광고주 검색</span> <span id="teamcode">운영팀 검색</span> <input type="image" src="/images/btn_search.gif" width="39" height="20" align="absmiddle" border=0></td>
                  <td  align="right" background="/images/bg_search.gif" ><!-- <img src="/images/btn_new.gif" width="30" height="30" align="absmiddle" border="0" alt="신규 계약 등록"></a><a href="#" onclick="fncPrint();"><img src="/images/btn_print.gif" width="30" height="30" align="absmiddle" border="0" alt="관리보고서 출력" hspace=2></a><a href="#" onclick="get_excel_sheet();"><img src="/images/btn_xls.gif" width="30" height="30" align="absmiddle" border="0" alt="엑셀 변환"></a> -->
				</td>
                  <td width="13" ><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="35" >&nbsp;(단위:천원)</td>
			<td align='right'><a href="#" onclick="getexcel(); return false;"><img src='/images/icon_xls.gif' width='17' height='16' align='bottom'> 엑셀 </a>  </td>
          </tr>
          <tr>
            <td valign="top" height="516" colspan='2'>

				  <table border="0"width="1030" cellpadding="0" cellspacing="0" bordercolor="#8D652B" id="contact">
					<tr height='30' align='center'>
						<th class="hd left" width="100">대분류</th>
						<th class="hd center" width="100">중분류</th>
						<th class="hd center" width="100">01월</th>
						<th class="hd center" width="100">02월</th>
						<th class="hd center" width="100">03월</th>
						<th class="hd center" width="100">04월</th>
						<th class="hd center" width="100">05월</th>
						<th class="hd center" width="100">06월</th>
						<th class="hd center" width="100">07월</th>
						<th class="hd center" width="100">08월</th>
						<th class="hd center" width="100">09월</th>
						<th class="hd center" width="100">10월</th>
						<th class="hd center" width="100">11월</th>
						<th class="hd center" width="100">12월</th>
						<th class="hd right" width="100">합계</th>
					</tr>
					<%
							Dim highclassname_old
							Do Until rs2.eof
							rs.Filter = "highclasscode="& highclasscode
							If rs.recordcount <> 0 Then
					%>
						<tr height='32'>

							<td  class="hd none" colspan='15' >
							<table  width='1030' border="1" style="table-layout:fixed;">
							<%
								Do Until rs.eof
							%>
								<tr height='30' <% If middleclassname = "소계" Then response.write "bgcolor=#ececec" End If %>>
									<td  width='100' class="hd none" ><%if highclassname_old <> highclassname then response.write highclassname %></td>
									<%highclassname_old = highclassname %>
									<td  width='100'  style='padding-left:3px;'><%=middleclassname%></td>
									<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(jan,0)%></td>
									<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(feb,0)%></td>
									<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(mar,0)%></td>
									<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(apr,0)%></td>
									<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(may,0)%></td>
									<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(jun,0)%></td>
									<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(jul,0)%></td>
									<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(aug,0)%></td>
									<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(sep,0)%></td>
									<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(oct_,0)%></td>
									<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(nov,0)%></td>
									<td  width='100' style='padding-right:5px;text-align:right;'><%=FormatNumber(dec,0)%></td>
									<td  width='100' style='padding-right:10px;text-align:right;'><%=FormatNumber(sum,0)%></td>
								</tr>
							<%
								If middleclassname="소계" Then
									s01 = s01 + jan
									s02 = s02 + feb
									s03 = s03 + mar
									s04 = s04 + apr
									s05 = s05 + may
									s06 = s06 + jun
									s07 = s07 + jul
									s08 = s08 + aug
									s09 = s09 + oct_
									s10 = s10 + nov
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
							End If
								rs2.movenext
							Loop
						%>
					<tr height='32' align='center'>
						<th class="hd left" colspan=2 > 총합계</th>
						<th class="hd center" width="100" align='right'><%=FormatNumber(s01,0)%>&nbsp;</th>
						<th class="hd center" width="100" align='right'><%=FormatNumber(s02,0)%>&nbsp;</th>
						<th class="hd center" width="100" align='right'><%=FormatNumber(s03,0)%>&nbsp;</th>
						<th class="hd center" width="100" align='right'><%=FormatNumber(s04,0)%>&nbsp;</th>
						<th class="hd center" width="100" align='right'><%=FormatNumber(s05,0)%>&nbsp;</th>
						<th class="hd center" width="100" align='right'><%=FormatNumber(s06,0)%>&nbsp;</th>
						<th class="hd center" width="100" align='right'><%=FormatNumber(s07,0)%>&nbsp;</th>
						<th class="hd center" width="100" align='right'><%=FormatNumber(s08,0)%>&nbsp;</th>
						<th class="hd center" width="100" align='right'><%=FormatNumber(s09,0)%>&nbsp;</th>
						<th class="hd center" width="100" align='right'><%=FormatNumber(s10,0)%>&nbsp;</th>
						<th class="hd center" width="100" align='right'><%=FormatNumber(s11,0)%>&nbsp;</th>
						<th class="hd center" width="100" align='right'><%=FormatNumber(s12,0)%>&nbsp;</th>
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