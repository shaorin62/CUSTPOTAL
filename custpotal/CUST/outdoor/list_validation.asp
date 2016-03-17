<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<%
	Dim userid : userid = request.cookies("userid")
	Dim pcustcode : pcustcode = request("cmbcustcode")
	Dim pteamcode : pteamcode = request("cmbteamcode")
	Dim strstat : strstat = request("cmbSTAT")
	Dim cyear : cyear = request("cyear")
	Dim cmonth : cmonth = request("cmonth")
	Dim cyear2 : cyear2 = request("cyear2")
	Dim cmonth2 : cmonth2 = request("cmonth2")

	If cyear = "" Then cyear = Year(date)
	If cmonth = "" Then cmonth = Month(date)
	If cyear2 = "" Then cyear2 = Year(date)
	If cmonth2 = "" Then cmonth2 = Month(date)
	If Len(cmonth) = 1 Then cmonth = "0"&cmonth
	If Len(cmonth2) = 1 Then cmonth2 = "0"&cmonth2
	dim schdate : schdate = cyear&cmonth

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




	Dim strwhere
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

				if strstat = "9" then
					strwhere = " and isnull(md.categoryidx,'') = '9' "
				elseif strstat = "10" then
					strwhere = " and isnull(md.categoryidx,'') = '10' "
				elseif strstat = "11" then
					strwhere = " and isnull(md.categoryidx,'') = '11' "
				elseif strstat = "135" then
					strwhere = " and isnull(md.categoryidx,'') = '135' "
				else
					strwhere = " and isnull(md.categoryidx,'') = '9' "
				end if

				sql = sql  & " select c.contidx, c.title, c.firstdate, c.startdate, "
				sql = sql  & " c.enddate, isnull(sum(m.monthly),0) as monthly, "
				sql = sql  & " c.flag , "
				sql = sql  & " isnull(max(s.a_val),0) a_val, isnull(max(s.b_val),0) b_val, isnull(max(s.c_val),0) c_val, isnull(max(s.d_val),0) d_val, isnull(max(s.e_val),0) e_val, "
				sql = sql  & " isnull(max(s.a_val),0) + isnull(max(s.b_val),0) + isnull(max(s.c_val),0) + isnull(max(s.d_val),0) + isnull(max(s.e_val),0) tot, "
				sql = sql  & " max(s.class) totclass"
				sql = sql  & " from wb_contact_mst c  "
				sql = sql  & " left outer join sc_cust_dtl d on c.custcode = d.custcode  "
				sql = sql  & " inner  join wb_account_cust n on d.highcustcode  = n.clientcode and n.userid='"&userid&"' and n.clientcode =  '" & objrs_1("clientcode") &"' "
				If Timcoderecord > 0 then
					sql = sql  & " inner  join wb_account_tim t on c.custcode = t.timcode and t.userid='"&userid&"' and t.clientcode = '" & objrs_1("clientcode") &"' "
				End If
				sql = sql  & " left outer join vw_contact_exe_monthly m  "
				sql = sql  & " on m.contidx = c.contidx  "
				sql = sql  & " left outer join wb_validation_class s on c.contidx = s.contidx and s.isuse = 1 "
				sql = sql  & " left outer join  "
				sql = sql  & " ( "
				sql = sql  & " 	select contidx, max(DBO.WB_CATEGORYIDX_FUN(categoryidx)) categoryidx "
				sql = sql  & " 	from  wb_contact_md  "
				sql = sql  & " 	group by contidx "
				sql = sql  & " ) md on c.contidx = md.contidx "
				sql = sql  & " where  m.cyear+m.cmonth = '"&schdate&"' " & strwhere
				sql = sql  & " and isnull(DBO.WB_CATEGORYIDX_FUN(md.categoryidx) ,'') <> '' "
				sql = sql  & " and c.flag = 'B' "
				sql = sql  & " group by c.contidx, c.title, c.firstdate,  "
				sql = sql  & " c.startdate, c.enddate, isnull(c.totalprice,0), c.custcode ,c.flag  "

				chkcount = chkcount +1
				objrs_1.movenext
			Loop
			sql = sql  & " order by c.enddate,  c.title,  contidx desc  "
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


	if strstat = "9" then
		strwhere = " and isnull(md.categoryidx,'') = '9' "
	elseif strstat = "10" then
		strwhere = " and isnull(md.categoryidx,'') = '10' "
	elseif strstat = "11" then
		strwhere = " and isnull(md.categoryidx,'') = '11' "
	elseif strstat = "135" then
		strwhere = " and isnull(md.categoryidx,'') = '135' "
	else
		strwhere = " and isnull(md.categoryidx,'') = '9' "
	end if

	sql = sql  & " select c.contidx, c.title, c.firstdate, c.startdate, "
	sql = sql  & " c.enddate, isnull(sum(m.monthly),0) as monthly, "
	sql = sql  & " c.flag , "
	sql = sql  & " isnull(max(s.a_val),0) a_val, isnull(max(s.b_val),0) b_val, isnull(max(s.c_val),0) c_val, isnull(max(s.d_val),0) d_val, isnull(max(s.e_val),0) e_val, "
	sql = sql  & " isnull(max(s.a_val),0) + isnull(max(s.b_val),0) + isnull(max(s.c_val),0) + isnull(max(s.d_val),0) + isnull(max(s.e_val),0) tot, "
	sql = sql  & " max(s.class) totclass"
	sql = sql  & " from wb_contact_mst c  "
	sql = sql  & " left outer join sc_cust_dtl d on c.custcode = d.custcode  "
	sql = sql  & " inner  join wb_account_cust n on d.highcustcode  = n.clientcode and n.userid='"&userid&"' "
	If Timcoderecord > 0 then
		sql = sql  & " inner  join wb_account_tim t on c.custcode = t.timcode and t.userid='"&userid&"' "
	End If
	sql = sql  & " left outer join vw_contact_exe_monthly m  "
	sql = sql  & " on m.contidx = c.contidx  "
	sql = sql  & " left outer join wb_validation_class s on c.contidx = s.contidx and s.isuse = 1 "
	sql = sql  & " left outer join  "
	sql = sql  & " ( "
	sql = sql  & " 	select contidx, max(DBO.WB_CATEGORYIDX_FUN(categoryidx)) categoryidx "
	sql = sql  & " 	from  wb_contact_md  "
	sql = sql  & " 	group by contidx "
	sql = sql  & " ) md on c.contidx = md.contidx "
	sql = sql  & " where  m.cyear+m.cmonth = '"&schdate&"' "
	sql = sql  & " and d.highcustcode like '"&pcustcode&"%'  "
	sql = sql  & " and c.custcode like  '"&pteamcode&"%'   "& strwhere
	sql = sql  & " and isnull(DBO.WB_CATEGORYIDX_FUN(md.categoryidx) ,'') <> '' "
	sql = sql  & " and c.flag = 'B' "
	sql = sql  & " group by c.contidx, c.title, c.firstdate,  "
	sql = sql  & " c.startdate, c.enddate, isnull(c.totalprice,0), c.custcode ,c.flag  "
	sql = sql  & " order by c.enddate,  c.title,  contidx desc  "

end if


	Set rs = server.CreateObject("adodb.recordset")
	rs.activeconnection = application("connectionstring")
	rs.cursorLocation = adUseClient
	rs.cursorType = adOpenStatic
	rs.lockType = adLockOptimistic
	rs.source = sql
	rs.open

	Dim totalrecord : totalrecord = rs.recordcount
	Dim real_totalrecord : real_totalrecord = rs.recordcount

	Dim contidx : Set contidx = rs(0)
	Dim title : Set title = rs(1)
	Dim firstdate : Set firstdate = rs(2)
	Dim startdate : Set startdate = rs(3)
	Dim enddate : Set enddate = rs(4)
	Dim monthly : Set monthly = rs(5)
	Dim flag : Set flag = rs(6)
	Dim a_val : Set a_val = rs(7)
	Dim b_val : Set b_val = rs(8)
	Dim c_val : Set c_val = rs(9)
	Dim d_val : Set d_val = rs(10)
	Dim e_val : Set e_val = rs(11)
	Dim tot : Set tot = rs(12)
	Dim totclass : Set totclass = rs(13)

	Dim grandmonthly : grandmonthly = 0
	Dim granda_val : granda_val = 0
	Dim grandb_val : grandb_val = 0
	Dim grandc_val : grandc_val = 0
	Dim grandd_val : grandd_val = 0
	Dim grande_val : grande_val = 0
	Dim grandtot : grandtot = 0

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

		function getcontact(contidx, flag) {
			// 계약 상세 현황 팝업
			var cyear = "<%=cyear%>";
			var cmonth = "<%=cmonth%>";
			var url = "/cust/outdoor/popup/view_"+flag+"_contact.asp?contidx="+contidx+"&cyear="+cyear+"&cmonth="+cmonth;
			var name = "contactdetail";
			var left = screen.width / 2 - 1024 / 2;
			var top = 10;
			var opt = "width=1260,  resiable=yes, scrollbars=yes, left="+left+", top="+top;
			window.open(url, name, opt);
		}

		function getexcel() {
			// 엑셀전환

			var custcode = document.getElementById("cmbcustcode").value;
			var teamcode = document.getElementById("cmbteamcode").value;
			var strstat = document.getElementById("cmbSTAT").value;
			var cyear = document.getElementById("cyear").value;
			var cmonth = document.getElementById("cmonth").value;

			location.href = "/cust/outdoor/excel/xls_validation.asp?custcode="+custcode+"&teamcode="+teamcode+"&strstat="+strstat+"&cyear="+cyear+"&cmonth="+cmonth;
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
<form action="list_validation.asp" method='post'>
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
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="boldtitle"> 효용성평가 현황 </span></TD>
				<TD width="50%" align="right"><span class="navigator" > 옥외광고 &gt;  옥외광고현황 &gt; 효용성평가 현황 </span></TD>
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
                  <td background="/images/bg_search.gif"> <img src="/images/icon_dot_search.gif" width="4" height="3" align="absmiddle"> 검색년월
				  <%call getyear(cyear)%> <%call getmonth(cmonth)%>
				  <!-- ~ <%call getyear2(cyear2)%> <%call getmonth2(cmonth2)%>--><span id="custcode">광고주 검색</span> <span id="teamcode">운영팀 검색</span>  <SELECT id="cmbSTAT"  style="WIDTH: 50px" name="cmbSTAT">
										<OPTION value="9" <% if strstat = "9" then response.write "selected" end if %>>LED</OPTION>
										<OPTION value="10" <% if strstat = "10" then response.write "selected" end if %>>옥탑</OPTION>
										<OPTION value="11" <% if strstat = "11" then response.write "selected" end if %>>야립</OPTION>
										<OPTION value="135" <% if strstat = "135" then response.write "selected" end if %>>기타</OPTION>
									</SELECT><input type="image" src="/images/btn_search.gif" width="39" height="20" align="absmiddle" border=0></td>
                  <td  align="right" background="/images/bg_search.gif" ><!-- <img src="/images/btn_new.gif" width="30" height="30" align="absmiddle" border="0" alt="신규 계약 등록"></a><a href="#" onclick="fncPrint();"><img src="/images/btn_print.gif" width="30" height="30" align="absmiddle" border="0" alt="관리보고서 출력" hspace=2></a><a href="#" onclick="get_excel_sheet();"><img src="/images/btn_xls.gif" width="30" height="30" align="absmiddle" border="0" alt="엑셀 변환"></a> -->
				</td>
                  <td width="13" ><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="35" >&nbsp;</td>
			<td align='right'><a href="#" onclick="getexcel(); return false;"><img src='/images/icon_xls.gif' width='17' height='16' align='bottom'> 엑셀 </a>  </td>
          </tr>
          <tr>
            <td valign="top" height="516" colspan='2'>
<% If strstat = "9"  Then %>
				  <table border="3px"   width="1030" cellpadding="0" cellspacing="0" bordercolor="#8d652b" id="contact">
				  <thead>
					<tr height='20' align='center'>
						<th width="30" style=' text-align:center;' rowSpan="2">no</th>
						<th style=' text-align:center;' rowSpan="2">매체명</th>
						<th width="70" style=' text-align:center;' rowSpan="2">최초<br>계약일자</th>
						<th style=' text-align:center;' colSpan="2">계약기간</th>
						<th width="90" style=' text-align:center;' rowSpan="2">월광고료(원)</th>
						<th style=' text-align:center;' colSpan="5">평가항목</th>
						<th width="50" style=' text-align:center;' rowSpan="2">총점(100)</th>
						<th width="50" style=' text-align:center;' rowSpan="2">등급</th>
					</tr>
					<tr height='35'>
						<th width="70"style=' text-align:center;'>시작일</th>
						<th width="70"style=' text-align:center;'>종료일</th>
						<th width="60" style=' text-align:center;'>지역<br>(40)</th>
						<th width="60" style=' text-align:center;'>매체사양<br>(20)</th>
						<th width="60" style=' text-align:center;'>가시환경<br>(30)</th>
						<th width="60" style=' text-align:center;'>경쟁환경<br>(5)</th>
						<th width="60" style=' text-align:center;'>기타<br>(5)</th>
					</tr>
					</thead>
					</table>
					<table border="0"   width="1030" cellpadding="0" cellspacing="0" id="contact">
					<tbody id='tbody'>
					<%
						Do Until rs.eof
					%>
					<tr class="trbd" height='32'>
						<td width="30"  class="hd none" style='padding-left:10px; text-align:left;'><span name="totalrecord"><%=totalrecord%></span></td>
						<td width="300"  class="hd none" style="padding-left:5px;"><a href="#" onclick="getcontact(<%=contidx%>,'<%=flag%>'); return false;" title="<%=title%>" class='subject'><%=cutTitle(title, 44)%></a></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=firstdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=startdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=enddate%></td>
						<td width="90"  class="hd none" style='padding-right:3px; text-align:right;'><%=FormatNumber(monthly, 0)%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=a_val%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=b_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=c_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=d_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=e_val%></td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center'><%=tot%></td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center;'><%=totclass%></td>
					</tr>

				  <%
							totalrecord = totalrecord - 1
							grandmonthly = CDbl(grandmonthly) + CDbl(monthly)
							granda_val = CDbl(granda_val) + CDbl(a_val)
							grandb_val = CDbl(grandb_val) + CDbl(b_val)
							grandc_val = CDbl(grandc_val) + CDbl(c_val)
							grandd_val = CDbl(grandd_val) + CDbl(d_val)
							grande_val = CDbl(grande_val) + CDbl(e_val)
							grandtot = CDbl(grandtot) + CDbl(tot)
							rs.movenext
						Loop

						if cdbl(real_totalrecord ) <> 0 then
							grandtot =round(CDbl(grandtot) / real_totalrecord /4, 2)
						end if
				  %>
				  </tbody>
				  <tfoot>
                <tr height="30">
                    <td class="hd left"  colspan='5' style="text-align:center;"><strong>평균</strong> </td>
                    <td class="hd center" style=' text-align:right; font-weight:bold'><%=formatnumber(grandmonthly,0)%>&nbsp;</td>
                    <td  class="hd center" style='padding-right:3px; text-align:right;'><%=granda_val%></td>
					<td  class="hd center" style='padding-right:3px; text-align:right;'><%=grandb_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandc_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandd_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grande_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:center;'><%=grandtot%></td>
                    <td class="hd right"></td>
                  </tr>
				  </tfoot>
              </table>
<% elseIf strstat = "10"  Then %>
				  <table border="3px"   width="1030" cellpadding="0" cellspacing="0" bordercolor="#8d652b" id="contact">
				  <thead>
					<tr height='20' align='center'>
						<th width="30" style=' text-align:center;' rowSpan="2">no</th>
						<th style=' text-align:center;' rowSpan="2">매체명</th>
						<th width="70" style=' text-align:center;' rowSpan="2">최초<br>계약일자</th>
						<th style=' text-align:center;' colSpan="2">계약기간</th>
						<th width="90" style=' text-align:center;' rowSpan="2">월광고료(원)</th>
						<th style=' text-align:center;' colSpan="5">평가항목</th>
						<th width="50" style=' text-align:center;' rowSpan="2">총점(100)</th>
						<th width="50" style=' text-align:center;' rowSpan="2">등급</th>
					</tr>
					<tr height='35'>
						<th width="70"style=' text-align:center;'>시작일</th>
						<th width="70"style=' text-align:center;'>종료일</th>
						<th width="60" style=' text-align:center;'>지역<br>(30)</th>
						<th width="60" style=' text-align:center;'>매체사양<br>(25)</th>
						<th width="60" style=' text-align:center;'>가시환경<br>(25)</th>
						<th width="60" style=' text-align:center;'>경쟁환경<br>(10)</th>
						<th width="60" style=' text-align:center;'>기타<br>(10)</th>
					</tr>
					</thead>
					</table>
					<table border="0"   width="1030" cellpadding="0" cellspacing="0" id="contact">
					<tbody id='tbody'>
					<%
						Do Until rs.eof
					%>
					<tr class="trbd" height='32'>
						<td width="30"  class="hd none" style='padding-left:10px; text-align:left;'><span name="totalrecord"><%=totalrecord%></span></td>
						<td width="300"  class="hd none" style="padding-left:5px;"><a href="#" onclick="getcontact(<%=contidx%>,'<%=flag%>'); return false;" title="<%=title%>" class='subject'><%=cutTitle(title, 44)%></a></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=firstdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=startdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=enddate%></td>
						<td width="90"  class="hd none" style='padding-right:3px; text-align:right;'><%=FormatNumber(monthly, 0)%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=a_val%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=b_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=c_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=d_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=e_val%></td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center'><%=tot%></td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center;'><%=totclass%></td>
					</tr>

				  <%
							totalrecord = totalrecord - 1
							grandmonthly = CDbl(grandmonthly) + CDbl(monthly)
							granda_val = CDbl(granda_val) + CDbl(a_val)
							grandb_val = CDbl(grandb_val) + CDbl(b_val)
							grandc_val = CDbl(grandc_val) + CDbl(c_val)
							grandd_val = CDbl(grandd_val) + CDbl(d_val)
							grande_val = CDbl(grande_val) + CDbl(e_val)
							grandtot = CDbl(grandtot) + CDbl(tot)
							rs.movenext
						Loop

						if cdbl(real_totalrecord ) <> 0 then
							grandtot =round(CDbl(grandtot) / real_totalrecord /4, 2)
						end if
				  %>
				  </tbody>
				  <tfoot>
                <tr height="30">
                    <td class="hd left"  colspan='5' style="text-align:center;"><strong>평균</strong> </td>
                    <td class="hd center" style=' text-align:right; font-weight:bold'><%=formatnumber(grandmonthly,0)%>&nbsp;</td>
                    <td  class="hd center" style='padding-right:3px; text-align:right;'><%=granda_val%></td>
					<td  class="hd center" style='padding-right:3px; text-align:right;'><%=grandb_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandc_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandd_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grande_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:center;'><%=grandtot%></td>
                    <td class="hd right"></td>
                  </tr>
				  </tfoot>
              </table>
<% elseIf strstat = "11"  Then %>
				  <table border="3px"   width="1030" cellpadding="0" cellspacing="0" bordercolor="#8d652b" id="contact">
				  <thead>
					<tr height='20' align='center'>
						<th width="30" style=' text-align:center;' rowSpan="2">no</th>
						<th style=' text-align:center;' rowSpan="2">매체명</th>
						<th width="70" style=' text-align:center;' rowSpan="2">최초<br>계약일자</th>
						<th style=' text-align:center;' colSpan="2">계약기간</th>
						<th width="90" style=' text-align:center;' rowSpan="2">월광고료(원)</th>
						<th style=' text-align:center;' colSpan="5">평가항목</th>
						<th width="50" style=' text-align:center;' rowSpan="2">총점(100)</th>
						<th width="50" style=' text-align:center;' rowSpan="2">등급</th>
					</tr>
					<tr height='35'>
						<th width="70"style=' text-align:center;'>시작일</th>
						<th width="70"style=' text-align:center;'>종료일</th>
						<th width="60" style=' text-align:center;'>지역<br>(30)</th>
						<th width="60" style=' text-align:center;'>매체사양<br>(25)</th>
						<th width="60" style=' text-align:center;'>가시환경<br>(25)</th>
						<th width="60" style=' text-align:center;'>기타<br>(10)</th>
						<th width="60" style=' text-align:center;'></th>
					</tr>
					</thead>
					</table>
					<table border="0"   width="1030" cellpadding="0" cellspacing="0" id="contact">
					<tbody id='tbody'>
					<%
						Do Until rs.eof
					%>
					<tr class="trbd" height='32'>
						<td width="30"  class="hd none" style='padding-left:10px; text-align:left;'><span name="totalrecord"><%=totalrecord%></span></td>
						<td width="300"  class="hd none" style="padding-left:5px;"><a href="#" onclick="getcontact(<%=contidx%>,'<%=flag%>'); return false;" title="<%=title%>" class='subject'><%=cutTitle(title, 44)%></a></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=firstdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=startdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=enddate%></td>
						<td width="90"  class="hd none" style='padding-right:3px; text-align:right;'><%=FormatNumber(monthly, 0)%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=a_val%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=b_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=c_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=d_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'>0</td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center'><%=tot%></td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center;'><%=totclass%></td>
					</tr>

				  <%
							totalrecord = totalrecord - 1
							grandmonthly = CDbl(grandmonthly) + CDbl(monthly)
							granda_val = CDbl(granda_val) + CDbl(a_val)
							grandb_val = CDbl(grandb_val) + CDbl(b_val)
							grandc_val = CDbl(grandc_val) + CDbl(c_val)
							grandd_val = CDbl(grandd_val) + CDbl(d_val)
							grande_val = CDbl(grande_val) + CDbl(e_val)
							grandtot = CDbl(grandtot) + CDbl(tot)
							rs.movenext
						Loop
						if cdbl(real_totalrecord ) <> 0 then
							grandtot =round(CDbl(grandtot) / real_totalrecord /4, 2)
						end if
				  %>
				  </tbody>
				  <tfoot>
                <tr height="30">
                    <td class="hd left"  colspan='5' style="text-align:center;"><strong>평균</strong> </td>
                    <td class="hd center" style=' text-align:right; font-weight:bold'><%=formatnumber(grandmonthly,0)%>&nbsp;</td>
                    <td  class="hd center" style='padding-right:3px; text-align:right;'><%=granda_val%></td>
					<td  class="hd center" style='padding-right:3px; text-align:right;'><%=grandb_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandc_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandd_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grande_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:center;'><%=grandtot%></td>
                    <td class="hd right"></td>
                  </tr>
				  </tfoot>
              </table>
<% elseIf strstat = "135"  Then %>
				  <table border="3px"   width="1030" cellpadding="0" cellspacing="0" bordercolor="#8d652b" id="contact">
				  <thead>
					<tr height='20' align='center'>
						<th width="30" style=' text-align:center;' rowSpan="2">no</th>
						<th style=' text-align:center;' rowSpan="2">매체명</th>
						<th width="70" style=' text-align:center;' rowSpan="2">최초<br>계약일자</th>
						<th style=' text-align:center;' colSpan="2">계약기간</th>
						<th width="90" style=' text-align:center;' rowSpan="2">월광고료(원)</th>
						<th style=' text-align:center;' colSpan="5">평가항목</th>
						<th width="50" style=' text-align:center;' rowSpan="2">총점(100)</th>
						<th width="50" style=' text-align:center;' rowSpan="2">등급</th>
					</tr>
					<tr height='35'>
						<th width="70"style=' text-align:center;'>시작일</th>
						<th width="70"style=' text-align:center;'>종료일</th>
						<th width="60" style=' text-align:center;'>지역<br>(30)</th>
						<th width="60" style=' text-align:center;'>매체사양<br>(40)</th>
						<th width="60" style=' text-align:center;'>가시환경<br>(15)</th>
						<th width="60" style=' text-align:center;'>기타<br>(15)</th>
						<th width="60" style=' text-align:center;'></th>
					</tr>
					</thead>
					</table>
					<table border="0"   width="1030" cellpadding="0" cellspacing="0" id="contact">
					<tbody id='tbody'>
					<%
						Do Until rs.eof
					%>
					<tr class="trbd" height='32'>
						<td width="30"  class="hd none" style='padding-left:10px; text-align:left;'><span name="totalrecord"><%=totalrecord%></span></td>
						<td width="300"  class="hd none" style="padding-left:5px;"><a href="#" onclick="getcontact(<%=contidx%>,'<%=flag%>'); return false;" title="<%=title%>" class='subject'><%=cutTitle(title, 44)%></a></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=firstdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=startdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=enddate%></td>
						<td width="90"  class="hd none" style='padding-right:3px; text-align:right;'><%=FormatNumber(monthly, 0)%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=a_val%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=b_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=c_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=d_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'>0</td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center'><%=tot%></td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center;'><%=totclass%></td>
					</tr>

				  <%
							totalrecord = totalrecord - 1
							grandmonthly = CDbl(grandmonthly) + CDbl(monthly)
							granda_val = CDbl(granda_val) + CDbl(a_val)
							grandb_val = CDbl(grandb_val) + CDbl(b_val)
							grandc_val = CDbl(grandc_val) + CDbl(c_val)
							grandd_val = CDbl(grandd_val) + CDbl(d_val)
							grande_val = CDbl(grande_val) + CDbl(e_val)
							grandtot = CDbl(grandtot) + CDbl(tot)
							rs.movenext
						Loop

						if cdbl(real_totalrecord ) <> 0 then
							grandtot =round(CDbl(grandtot) / real_totalrecord /4, 2)
						end if
				  %>
				  </tbody>
				  <tfoot>
                <tr height="30">
                    <td class="hd left"  colspan='5' style="text-align:center;"><strong>평균</strong> </td>
                    <td class="hd center" style=' text-align:right; font-weight:bold'><%=formatnumber(grandmonthly,0)%>&nbsp;</td>
                    <td  class="hd center" style='padding-right:3px; text-align:right;'><%=granda_val%></td>
					<td  class="hd center" style='padding-right:3px; text-align:right;'><%=grandb_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandc_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandd_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grande_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:center;'><%=grandtot%></td>
                    <td class="hd right"></td>
                  </tr>
				  </tfoot>
              </table>
<% Else %>
				  <table border="3px"   width="1030" cellpadding="0" cellspacing="0" bordercolor="#8d652b" id="contact">
				  <thead>
					<tr height='20' align='center'>
						<th width="30" style=' text-align:center;' rowSpan="2">no</th>
						<th style=' text-align:center;' rowSpan="2">매체명</th>
						<th width="70" style=' text-align:center;' rowSpan="2">최초<br>계약일자</th>
						<th style=' text-align:center;' colSpan="2">계약기간</th>
						<th width="90" style=' text-align:center;' rowSpan="2">월광고료(원)</th>
						<th style=' text-align:center;' colSpan="5">평가항목</th>
						<th width="50" style=' text-align:center;' rowSpan="2">총점(100)</th>
						<th width="50" style=' text-align:center;' rowSpan="2">등급</th>
					</tr>
					<tr height='35'>
						<th width="70"style=' text-align:center;'>시작일</th>
						<th width="70"style=' text-align:center;'>종료일</th>
						<th width="60" style=' text-align:center;'>지역<br>(40)</th>
						<th width="60" style=' text-align:center;'>매체사양<br>(20)</th>
						<th width="60" style=' text-align:center;'>가시환경<br>(30)</th>
						<th width="60" style=' text-align:center;'>경쟁환경<br>(5)</th>
						<th width="60" style=' text-align:center;'>기타<br>(5)</th>
					</tr>
					</thead>
					</table>
					<table border="0"   width="1030" cellpadding="0" cellspacing="0" id="contact">
					<tbody id='tbody'>
					<%
						Do Until rs.eof
					%>
					<tr class="trbd" height='32'>
						<td width="30"  class="hd none" style='padding-left:10px; text-align:left;'><span name="totalrecord"><%=totalrecord%></span></td>
						<td width="300"  class="hd none" style="padding-left:5px;"><a href="#" onclick="getcontact(<%=contidx%>,'<%=flag%>'); return false;" title="<%=title%>" class='subject'><%=cutTitle(title, 44)%></a></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=firstdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=startdate%></td>
						<td width="70"  class="hd none" style=' text-align:center;'><%=enddate%></td>
						<td width="90"  class="hd none" style='padding-right:3px; text-align:right;'><%=FormatNumber(monthly, 0)%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=a_val%></td>
						<td width="60"  class="hd none" style='padding-right:3px; text-align:right;'><%=b_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=c_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=d_val%></td>
						<td width="60"  class="hd none" style='padding-right:10px; text-align:right;'><%=e_val%></td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center'><%=tot%></td>
						<td width="50"  class="hd none" style='padding-left:3px; text-align:center;'><%=totclass%></td>
					</tr>

				  <%
							totalrecord = totalrecord - 1
							grandmonthly = CDbl(grandmonthly) + CDbl(monthly)
							granda_val = CDbl(granda_val) + CDbl(a_val)
							grandb_val = CDbl(grandb_val) + CDbl(b_val)
							grandc_val = CDbl(grandc_val) + CDbl(c_val)
							grandd_val = CDbl(grandd_val) + CDbl(d_val)
							grande_val = CDbl(grande_val) + CDbl(e_val)
							grandtot = CDbl(grandtot) + CDbl(tot)
							rs.movenext
						Loop

						if cdbl(real_totalrecord ) <> 0 then
							grandtot =round(CDbl(grandtot) / real_totalrecord /4, 2)
						end if
				  %>
				  </tbody>
				  <tfoot>
                <tr height="30">
                    <td class="hd left"  colspan='5' style="text-align:center;"><strong>평균</strong> </td>
                    <td class="hd center" style=' text-align:right; font-weight:bold'><%=formatnumber(grandmonthly,0)%>&nbsp;</td>
                    <td  class="hd center" style='padding-right:3px; text-align:right;'><%=granda_val%></td>
					<td  class="hd center" style='padding-right:3px; text-align:right;'><%=grandb_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandc_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grandd_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:right;'><%=grande_val%></td>
					<td  class="hd center" style='padding-right:10px; text-align:center;'><%=grandtot%></td>
                    <td class="hd right"></td>
                  </tr>
				  </tfoot>
              </table>
<% End If %>




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