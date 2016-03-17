<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<%
	Dim userid : userid = request.cookies("userid")
	Dim pcustcode : pcustcode = request("cmbcustcode")
	Dim pteamcode : pteamcode = request("cmbteamcode")
	Dim cmbseqno : cmbseqno = request("cmbseqno")
	Dim cmbthmno : cmbthmno = request("cmbthmno")

	Dim cyear : cyear = request("cyear")
	Dim cmonth : cmonth = request("cmonth")
	If cyear = "" Then cyear = Year(date)
	If cmonth = "" Then cmonth = Month(date)
	If Len(cmonth) = 1 Then cmonth = "0"&cmonth
	dim sdate : sdate = dateserial(cyear, cmonth, "01")
	dim edate : edate = dateadd("d", -1, dateadd("m", 1,  sdate))

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



				if cmbseqno = "" then
					sql = sql  & " select a.contidx, a.custcode, a.title, a.startdate, a.enddate, a.flag from wb_contact_mst a    "
					sql = sql  & " left outer join sc_cust_dtl b on a.custcode = b.custcode "
					sql = sql  & " inner  join wb_account_cust n on b.highcustcode  = n.clientcode and n.userid='"&userid&"' and n.clientcode =  '" & objrs_1("clientcode") &"' "
					If Timcoderecord > 0 then
						sql = sql  & " inner  join wb_account_tim t on a.custcode = t.timcode and t.userid='"&userid&"' and t.clientcode = '" & objrs_1("clientcode") &"' "
					End If
					sql = sql  & " inner join wb_contact_md c on a.contidx=c.contidx "
					sql = sql  & " inner join vw_contact_md_dtl d on c.mdidx=d.mdidx "
					sql = sql  & " left outer join vw_subseq_exe e on e.mdidx=d.mdidx and d.side=e.side and e.cyear = '" & cyear &"' and e.cmonth = '"&cmonth &"' "
					sql = sql  & " left outer join tmp_subseq_mtx f on e.thmno=f.thmno and seqno like '"&cmbseqno&"%' and e.thmno like '"&cmbthmno&"%' "
					sql = sql  & " where a.startdate <= '"&edate&"' and a.enddate >= '"&sdate&"'  "
					sql = sql  & " group by a.contidx, a.custcode, a.title, a.startdate, a.enddate, a.flag "

				Else

					sql = sql  & " select a.contidx, a.custcode, a.title, a.startdate, a.enddate, a.flag from wb_contact_mst a    "
					sql = sql  & " left outer join sc_cust_dtl b on a.custcode = b.custcode "
					sql = sql  & " inner  join wb_account_cust n on b.highcustcode  = n.clientcode and n.userid='"&userid&"' and n.clientcode =  '" & objrs_1("clientcode") &"' "
					If Timcoderecord > 0 then
						sql = sql  & " inner  join wb_account_tim t on a.custcode = t.timcode and t.userid='"&userid&"' and t.clientcode = '" & objrs_1("clientcode") &"' "
					End If
					sql = sql  & " inner join wb_contact_md c on a.contidx=c.contidx  "
					sql = sql  & " inner join vw_contact_md_dtl d on c.mdidx=d.mdidx "
					sql = sql  & " inner join vw_subseq_exe e on e.mdidx=d.mdidx and d.side=e.side and e.cyear = '" & cyear &"' and e.cmonth = '"&cmonth &"'   "
					sql = sql  & " inner join tmp_subseq_mtx f on e.thmno=f.thmno and seqno like '"&cmbseqno&"%' and e.thmno like '"&cmbthmno&"%'  "
					sql = sql  & " where a.startdate <= '"&edate&"' and a.enddate >= '"&sdate&"'  "
					sql = sql  & " group by a.contidx, a.custcode, a.title, a.startdate, a.enddate, a.flag "
				end if


				chkcount = chkcount +1
				objrs_1.movenext
			Loop

			sql = sql  & " order by contidx desc "
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

	if cmbseqno = "" then
		sql = "select a.contidx, a.custcode, a.title, a.startdate, a.enddate, a.flag from wb_contact_mst a    "
		sql = sql  & " left outer join sc_cust_dtl b on a.custcode = b.custcode "
		sql = sql  & " inner  join wb_account_cust n on b.highcustcode  = n.clientcode and n.userid='"&userid&"' "
		If Timcoderecord > 0 then
			sql = sql  & " inner  join wb_account_tim t on a.custcode = t.timcode and t.userid='"&userid&"' "
		End If
		sql = sql  & " inner join wb_contact_md c on a.contidx=c.contidx "
		sql = sql  & " inner join vw_contact_md_dtl d on c.mdidx=d.mdidx "
		sql = sql  & " left outer join vw_subseq_exe e on e.mdidx=d.mdidx and d.side=e.side and e.cyear = '" & cyear &"' and e.cmonth = '"&cmonth &"' "
		sql = sql  & " left outer join tmp_subseq_mtx f on e.thmno=f.thmno and seqno like '"&cmbseqno&"%' and e.thmno like '"&cmbthmno&"%' "
		sql = sql  & " where a.startdate <= '"&edate&"' and a.enddate >= '"&sdate&"'  "
		sql = sql  & " and a.custcode like '"&pteamcode&"%' and b.highcustcode like '"&pcustcode&"%' "
		sql = sql  & " group by a.contidx, a.custcode, a.title, a.startdate, a.enddate, a.flag "

	Else

		sql = "select a.contidx, a.custcode, a.title, a.startdate, a.enddate, a.flag from wb_contact_mst a    "
		sql = sql  & " left outer join sc_cust_dtl b on a.custcode = b.custcode "
		sql = sql  & " inner  join wb_account_cust n on b.highcustcode  = n.clientcode and n.userid='"&userid&"' "
		If Timcoderecord > 0 then
			sql = sql  & " inner  join wb_account_tim t on a.custcode = t.timcode and t.userid='"&userid&"' "
		End If
		sql = sql  & " inner join wb_contact_md c on a.contidx=c.contidx  "
		sql = sql  & " inner join vw_contact_md_dtl d on c.mdidx=d.mdidx "
		sql = sql  & " inner join vw_subseq_exe e on e.mdidx=d.mdidx and d.side=e.side and e.cyear = '" & cyear &"' and e.cmonth = '"&cmonth &"'   "
		sql = sql  & " inner join tmp_subseq_mtx f on e.thmno=f.thmno and seqno like '"&cmbseqno&"%' and e.thmno like '"&cmbthmno&"%'  "
		sql = sql  & " where a.startdate <= '"&edate&"' and a.enddate >= '"&sdate&"'  "
		sql = sql  & " and a.custcode like '"&pteamcode&"%' and b.highcustcode like '"&pcustcode&"%' "
		sql = sql  & " group by a.contidx, a.custcode, a.title, a.startdate, a.enddate, a.flag "


	end if


end if

	Set rs = server.CreateObject("adodb.recordset")
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
	Dim flag : Set flag = rs(5)

	sql = "select a.contidx,  b.side, c.monthly, d.thmno ,  a.mdidx from wb_contact_md a inner join vw_contact_md_dtl b on a.mdidx=b.mdidx inner join wb_contact_exe c on b.mdidx=c.mdidx and b.side=c.side and c.cyear='"&cyear&"' and c.cmonth='"&cmonth&"' left outer join  vw_subseq_exe d on c.mdidx=d.mdidx and c.side=d.side and d.cyear='"&cyear&"' and d.cmonth='"&cmonth&"' left outer join tmp_subseq_mtx e on d.thmno=e.thmno where seqno like '"&cmbseqno&"%' and d.thmno like '"&cmbthmno&"%' order by a.contidx desc, case when  b.side <> 'L' then ' ' +b.side else b.side end  desc"

	Dim rs2 : Set rs2 = server.CreateObject("adodb.recordset")
	rs2.activeconnection = application("connectionstring")
	rs2.cursorlocation = aduseclient
	rs2.cursortype = adopenstatic
	rs2.locktype = adLockOptimistic
	rs2.source = sql
	rs2.open

	If Not rs2.eof Then
		Dim side : Set side = rs2(1)
		Dim monthly : Set monthly = rs2(2)
		Dim thmno : Set thmno = rs2(3)
	End If

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
						document.getElementById("cmbcustcode").style.width="220px";
						document.getElementById("cmbcustcode").attachEvent("onchange", getteamcombo);
						document.getElementById("cmbcustcode").attachEvent("onchange", getbrandcode);
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
						document.getElementById("cmbteamcode").style.width="180px";
				}
			}
		}

		function getbrandcode() {
		// 광고주를 선택 했을때 실행

			var custcode = document.getElementById("cmbcustcode").value;
			var seqno = "<%=cmbseqno%>" ;
			if (custcode == "") seqno = null;
			var params = "custcode="+custcode+"&seqno="+seqno+ "&Custcodesql="+"<%=Custcodesql%>"  ;
			//alert(params);
			_sendRequest("/cust/outdoor/inc/getbrandcombo.asp", params, _getbrandcode, "GET");
			_sendRequest("/cust/outdoor/inc/getthemecombo.asp", params, _getthemecode, "GET");
		}

		function _getbrandcode() {
			
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
					var displayseqno = document.getElementById("displayseqno");
					if (displayseqno) {
						displayseqno.innerHTML = xmlreq.responseText ;
						document.getElementById("cmbseqno").attachEvent("onchange", getthemecode);
					}
				}
			}
		}

		function getthemecode() {
			//소재를 선택 했을때 실행
			var highseqno = document.getElementById("cmbseqno").value;
			var thmno = "<%=cmbthmno%>" ;
			var params = "highseqno="+highseqno+"&thmno="+thmno ;
			sendRequest("/cust/outdoor/inc/getthemecombo.asp", params, _getthemecode, "GET");
		}

		function _getthemecode() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
					var displaythmno = document.getElementById("displaythmno");
						displaythmno.innerHTML = xmlreq.responseText ;
				}
			}
		}


		function getexcel() {
			// 엑셀전환
			var custcode = document.getElementById("cmbcustcode").value;
			var teamcode = document.getElementById("cmbteamcode").value;
			var cmbseqno = document.getElementById("cmbseqno").value;
			var cmbthmno = document.getElementById("cmbthmno").value;
			var cyear = document.getElementById("cyear").value;
			var cmonth = document.getElementById("cmonth").value;

			location.href = "/cust/outdoor/excel/xls_brand.asp?custcode="+custcode+"&teamcode="+teamcode+"&cyear="+cyear+"&cmonth="+cmonth+"&cmbseqno="+cmbseqno+"&cmbthmno="+cmbthmno;
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

	function debug() {
		var debug = document.getElementById("debugConsole");
		debug.innerHTML =  xmlreq.responseText ;
	}

		window.onload = function () {
			_sendRequest("/inc/getcustcombo_cust.asp", "custcode=<%=pcustcode%>", _getcustcombo, "GET");
			_sendRequest("/inc/getteamcombo_cust.asp", "custcode=<%=pcustcode%>&teamcode=<%=pteamcode%>", _getteamcombo, "GET");
			_sendRequest("/cust/outdoor/inc/getbrandcombo.asp", "custcode=<%=pcustcode%>&seqno=<%=cmbseqno%>&Custcodesql=<%=Custcodesql%>", _getbrandcode, "GET");
			//_sendRequest("/cust/outdoor/inc/getsubbrandcombo.asp",  "highseqno=<%=cmbseqno%>&subno=<%=cmbsubno%>", _getsubbrandcode, "GET");
			_sendRequest("/cust/outdoor/inc/getthemecombo.asp", "custcode=<%=pcustcode%>&highseqno=<%=cmbseqno%>&thmno=<%=cmbthmno%>&Custcodesql=<%=Custcodesql%>", _getthemecode, "GET");
			document.getElementById("cmbcustcode").attachEvent("onchange", getteamcombo);
		}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  >
<form action="list_brand.asp" method='post'>
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
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="boldtitle"> 브랜드별 집행현황 </span></TD>
				<TD width="50%" align="right"><span class="navigator" > 옥외관리 &gt;  옥외광고현황 &gt; 브랜드별 집행현황 </span></TD>
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
				  <%call getyear(cyear)%> <%call getmonth(cmonth)%> &nbsp;    <span id="custcode">광고주 검색</span> <span id="teamcode">운영팀 검색</span>  <span id='displayseqno'> 브랜드 검색 </span><!--  <span id='displaysubno'> 서브 브랜드 검색 </span> --> <span id='displaythmno'> 소재명 검색 </span> &nbsp;<input type="image" src="/images/btn_search.gif" width="39" height="20" align="absmiddle" border=0></td>
				</td>
                  <td width="13" ><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="35" ></td>
			<td align='right'><a href="#" onclick="getexcel(); return false;"><img src='/images/icon_xls.gif' width='17' height='16' align='bottom'> 엑셀 </a> </td>
          </tr>
          <tr>
            <td valign="top" height="516" colspan='2'>

	  <table border="0"width="1030" cellpadding="0" cellspacing="0" bordercolor="#8D652B" id="contact">
	  <thead>
			<tr height='30' align='center'>
				<th  width="30" class="hd left">No</th >
				 <th  width="210" class="hd center">매체명</th >
				 <th  width="80" class="hd center">시작일자</th >
				 <th  width="80" class="hd center">종료일자</th >
				 <th  width="40" class="hd center">면</th >
				 <th  width="110" class="hd center">브랜드</th >
				 <th  width="110" class="hd center">서브브랜드</th >
				 <th  width="110" class="hd center">소재명</th >
				 <th  width="80" class="hd center">월광고료</th >
				 <th  width="90" class="hd center">광고주</th >
				 <th  width="90" class="hd right">운영팀</th >
			</tr>
		</thead>
		<tbody id='tbody'>
		<%
				Do Until rs.eof
		%>
			<tr height='32'>
				<td  class="hd none" style='text-align:center;padding-top:9px;padding-left:11px;vertical-align:top;'  width="30"><%=totalrecord%> </td>
				<td  class="hd none" style='text-align:left;padding-top:9px;vertical-align:top;' width="210" title='<%=title%>' ><a href="#" onclick="getcontact(<%=contidx%>, '<%=flag%>'); return false;"><%=cutTitle(title, 30)%></a></td>
				<td  class="hd none"style='text-align:center;padding-top:9px;vertical-align:top;' width="80"><%=startdate%></td>
				<td  class="hd none" style='text-align:center;padding-top:9px;vertical-align:top;' width="80"><%=enddate%></td>
				<td  class="hd none" colspan='5'><table  width='450' border=0 style="table-layout:fixed;">
				<%
					rs2.Filter = "contidx="&contidx
					Do Until rs2.eof
				%>
					<tr height='32'>
						<td  width="45" style='text-align:center;'><%=side%></td>
						<td  width="110" style='padding-left:5px;'><%=getbrand(thmno)%></td>
						<td  width="110" style='padding-left:5px;'><%=getsubbrand(thmno)%></td>
						<td  width="110" style='padding-left:5px;'><%=getthmname(thmno)%></td>
						<td  width="80" style='text-align:right;padding-right:10px;'><%=FormatNumber(monthly,0)%></td>
					</tr>
				<%
					rs2.movenext
					Loop
				%>
			</table></td>
			<td  class="hd none" width="90" style='text-align:left;padding-top:9px;padding-left:5px;vertical-align:top;' ><%=getcustname(custcode)%></td>
			<td  class="hd none" width="90" style='text-align:left;padding-top:9px;padding-left:5px;vertical-align:top;'  ><%=getteamname(custcode)%></td>
			</tr>
			<%
						totalrecord = totalrecord-1
						rs.movenext
					Loop
			%>
		</tbody>
        </table>
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