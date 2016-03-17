<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<%
	Dim userid : userid = request.cookies("userid")
	Dim pcustcode : pcustcode = request("cmbcustcode")
	Dim pteamcode : pteamcode = request("cmbteamcode")
	If pcustcode = "" Then pcustcode = "%"
	If pteamcode = "" Then pteamcode = "%"
	Dim cyear : cyear = request("cyear")
	Dim cmonth : cmonth = request("cmonth")
	If cyear = "" Then cyear = Year(date)
	If cmonth = "" Then cmonth = Month(date)
	If Len(cmonth) = 1 Then cmonth = "0"&cmonth


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



	if pcustcode = "%" or pcustcode = null then
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


				sql = sql & " select distinct a.contidx, a.custcode, a.title, a.startdate, a.enddate, a.flag, e.report from wb_contact_mst a  "
				sql = sql & " left outer join wb_contact_md b on a.contidx=b.contidx "
				sql = sql & " inner join wb_contact_exe c on b.mdidx=c.mdidx  "
				sql = sql & " inner join sc_cust_dtl d on a.custcode = d.custcode "
				sql = sql & " inner  join wb_account_cust n on d.highcustcode  = n.clientcode and n.userid='"&userid&"' and n.clientcode =  '" & objrs_1("clientcode") &"' "

				If Timcoderecord > 0 then
					sql = sql  & " inner  join wb_account_tim t on a.custcode = t.timcode and t.userid='"&userid&"' and t.clientcode = '" & objrs_1("clientcode") &"' "
				End If

				sql = sql & "  left outer join wb_report_mst e on e.contidx=a.contidx and e.cyear='" & cyear & "' and e.cmonth='" & cmonth &"' "
				sql = sql & " where   c.cyear='"&cyear&"' and c.cmonth='"&cmonth&"'  "


				chkcount = chkcount +1
				objrs_1.movenext
			Loop
			sql = sql & " order by a.contidx desc "
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


	sql = sql & " select distinct a.contidx, a.custcode, a.title, a.startdate, a.enddate, a.flag, e.report from wb_contact_mst a  "
	sql = sql & " left outer join wb_contact_md b on a.contidx=b.contidx "
	sql = sql & " inner join wb_contact_exe c on b.mdidx=c.mdidx  "
	sql = sql & " inner join sc_cust_dtl d on a.custcode = d.custcode "
	sql = sql & " inner  join wb_account_cust n on d.highcustcode  = n.clientcode and n.userid='"&userid&"' "

	If Timcoderecord > 0 then
		sql = sql  & " inner  join wb_account_tim t on a.custcode = t.timcode and t.userid='"&userid&"' "
	End If

	sql = sql & "  left outer join wb_report_mst e on e.contidx=a.contidx and e.cyear='" & cyear & "' and e.cmonth='" & cmonth &"' "
	sql = sql & " where  a.custcode like '"&pteamcode&"' and d.highcustcode like '"&pcustcode&"' and c.cyear='"&cyear&"' and c.cmonth='"&cmonth&"'  "
	sql = sql & " order by a.contidx desc "


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
	Dim teamcode : Set teamcode = rs(1)
	Dim title : Set title = rs(2)
	Dim startdate : Set startdate = rs(3)
	Dim enddate : Set enddate = rs(4)
	Dim flag : Set flag = rs(5)
	Dim report : Set report = rs(6)

'	response.write sql


	sql = "select sum(isnull(c.expense,0))  expense, a.region, a.locate, sum(isnull(c.qty,0)) qty, a.contidx, filename, a.mdidx, reportname, a.unit, z.custcode, a.medcode, a.empid   "
	sql = sql & " from wb_contact_mst z inner join wb_contact_md a on a.contidx=z.contidx  "
	sql = sql & " inner join wb_contact_exe c on a.mdidx=c.mdidx "
	sql = sql & " left outer join wb_report_dtl d on c.mdidx=d.mdidx and c.cyear = d.cyear and c.cmonth = d.cmonth "
	sql = sql & " inner join sc_cust_dtl e on z.custcode=e.custcode  "
	sql = sql  & " inner  join wb_account_cust n on e.highcustcode  = n.clientcode and n.userid='"&userid&"' "
	If Timcoderecord > 0 then
		sql = sql  & " inner  join wb_account_tim t on z.custcode = t.timcode and t.userid='"&userid&"' "
	End If
	sql = sql & " where c.cyear='"&cyear&"' and c.cmonth='" &cmonth&"' and z.custcode like '"&pteamcode&"' and e.highcustcode like '"&pcustcode&"'  "
	sql = sql & " group by a.region, a.locate, a.contidx, filename, a.mdidx, reportname, unit, z.custcode, a.medcode, a.empid "


'response.write sql

	Dim rs2 : Set rs2 = server.CreateObject("adodb.recordset")
	rs2.activeconnection = application("connectionstring")
	rs2.cursorlocation = aduseclient
	rs2.cursortype = adopenstatic
	rs2.locktype = adLockOptimistic
	rs2.source = sql
	rs2.open

	If Not rs2.eof Then
		Dim expense : Set expense = rs2(0)
		Dim region : Set region = rs2(1)
		Dim locate : Set locate = rs2(2)
		Dim qty : Set qty = rs2(3)
		Dim filename : Set filename = rs2(5)
		Dim mdidx : Set mdidx = rs2(6)
		Dim reportname : Set reportname = rs2(7)
		Dim Unit : Set unit = rs2(8)
		Dim custcode : Set custcode = rs2(9)
		Dim medcode : Set medcode = rs2(10)
		Dim empid : Set empid = rs2(11)
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
		var rows = 0;
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
				}
			}
		}

		function getprint() {
			// 관리 보고서 출력
			var frm = document.forms[0];
			frm.target = "_blank";
			frm.method = "post";
			frm.action = "/cust/outdoor/process/print_report.asp";
			frm.submit();
			cleanToggle();

		}

		function getDonwload() {
			var frm = document.forms[0];
			frm.method = "post";
			frm.target = "process";
			frm.action = "/cust/outdoor/process/down_html.asp";
			//getToggle();
			//getToggle2();
			frm.submit();
		}

		function goSearch() {
			var frm = document.forms[0];
			frm.method = "post";
			frm.action = "/cust/outdoor/list_report.asp";
			frm.submit();
		}

		function getconvertfile() {
			var frm = document.forms[0];
			frm.target = "process";
			frm.method = "post";
			frm.action = "/cust/outdoor/process/create_html.asp";
			frm.submit();
			cleanToggle();
		}

		function getdeletefile() {
			if (confirm("선택한 계약의 파일을 삭제하시겠습니까?")) {
				var frm = document.forms[0];
				frm.target = "process";
				frm.method = "post";
				frm.action = "/cust/outdoor/process/delete_html.asp";
				frm.submit();
			cleanToggle();
			}
		}

		function download(file) {
			location.href = '/med/download.asp?filename='+file;
		}
		function download2(file) {
			location.href = '/med/download2.asp?filename='+file;
		}

		function cleanToggle() {
			document.getElementById("toggle").checked = false;
			getToggle();

		}

		function getToggle() {
			var toggle =  document.getElementById("toggle").checked ;
			var elem = document.getElementsByTagName("input") ;
			for (var i=0 ; i < elem.length ; i++) {
				if (elem[i].getAttribute("className") == "check") {
					elem[i].checked = toggle;
				}
			}
		}

		function getToggle2() {
			var elem = document.getElementsByTagName("input") ;
			for (var i=0 ; i < elem.length ; i++) {
				if (elem[i].getAttribute("className") == "check") {
					if (elem[i].checked == true)  {
						elem[i].checked = true
					} else {
						elem[i].checked = false
					}
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

		window.onload = function () {
			_sendRequest("/inc/getcustcombo_cust.asp", "custcode=<%=pcustcode%>", _getcustcombo, "GET");
			_sendRequest("/inc/getteamcombo_cust.asp", "custcode=<%=pcustcode%>&teamcode=<%=pteamcode%>", _getteamcombo, "GET");
			document.getElementById("cmbcustcode").attachEvent("onchange", getteamcombo);
			document.getElementById("toggle").attachEvent("onclick", getToggle);
		}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  >
<form >
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
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="boldtitle"> 옥외광고 관리보고 </span></TD>
				<TD width="50%" align="right"><span class="navigator" > 옥외관리 &gt;  옥외광고현황 &gt; 옥외광고 관리보고 </span></TD>
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
				  <%call getyear(cyear)%> <%call getmonth(cmonth)%> &nbsp;    <span id="custcode">광고주 검색</span> <span id="teamcode">운영팀 검색</span>  <a href="#" onclick="goSearch(); return false;"><img src="/images/btn_search.gif" width="39" height="20" align="absmiddle" border=0></a></td>
				</td>
                  <td width="13" ><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="35" > <a href="#" onclick="getprint(); return false;"><img src='/images/m_print.gif' width='16' height='16' alt='선택한 계약의 관리보고서를 인쇄' align='bottom'> 관리보고서 인쇄  </a> <a href="#" onclick="getDonwload(); return false;"><img src='/images/m_download.gif' width='19' height='16' alt='선택된 ppt파일, 관리보고서 파일을 다운로드' align='bottom'> 파일 다운로드</a>   </td>
			<!--<td align='right'> <a href="#" onclick="getconvertfile();return false;"> <img src='/images/icon_report.gif' width='14' height='16' alt='선택한 계약의 파일생성'  align='bottom'> 파일생성 </a>  <a href="#" onclick="getdeletefile(); return false;"><img src='/images/m_remove.gif' width='13' height='12' alt='선택한 계약의 파일삭제' align='bottom'>  파일삭제 </a>

			</td>-->
          </tr>
          <tr>
            <td valign="top" height="516" colspan='2'>

	  <table border="0"width="1030" cellpadding="0" cellspacing="0" bordercolor="#8D652B" id="contact">
	  <thead>
			<tr height='30' align='center'>
				<th  width="30" class="hd left"><input type="checkbox" id="toggle" /></th >
				 <th  width="200" class="hd center">매체명</th >
				 <th  width="80" class="hd center">시작일자</th >
				 <th  width="80" class="hd center">종료일자</th >
				 <th  width="50" class="hd center">수량</th >
				 <th  width="200" class="hd center">매체위치</th >
				 <th  width="100" class="hd center">광고주</th >
				 <th  width="90" class="hd center">운영팀</th >
				 <th  width="100" class="hd center">매체사</th >
				 <th  width="60" class="hd center">담당자</th >
				 <th  width="30" class="hd right">파일</th >
			</tr>
		</thead>
		<tbody id='tbody'>
		<%
				Dim intLoop : intLoop = 1
				Do Until rs.eof
		%>
			<tr height='32'>
				<td  class="hd none" style='text-align:right;padding-right:15px;'  width="30"> <input type="checkbox" class='check' name='contidx' value='<%=contidx%>' /></td>
				<td  class="hd none" style='text-align:left;' width="200" title='<%=title%>' ><%If Not IsNull(report) Then %><a href="#" onclick="download2('<%=report%>'); return false;"><img src='/images/webpage.gif' width='18' height='18' alt='<%=report%>'></a><% End If %><a href="#" onclick="getcontact(<%=contidx%>, '<%=flag%>'); return false;"><%=cutTitle(title, 28)%></a></td>
				<td  class="hd none" style='text-align:center;' width="80"><%=startdate%></td>
				<td  class="hd none" style='text-align:center;' width="80"><%=enddate%></td>
				<td  class="hd none" colspan='8'><table  width='620' border=0 style="table-layout:fixed;">
				<%
					rs2.Filter = "contidx="&contidx
					Do Until rs2.eof
				%>
					<tr height='32'>
						<td  width="50" ><%=FormatNumber(qty,0)%> <%=unit%></td>
						<td  width="200" style='padding-left:3px;' title='<%=locate%>'><%=cutTitle(locate, 30)%></td>
						<td  width="100" style='padding-left:3px;'><%=getcustname(custcode)%></td>
						<td  width="90" style='padding-left:3px;'><%=getteamname(custcode)%></td>
						<td  width="100" style='padding-left:3px;'><%=getmedname(medcode)%></td>
						<td  width="60" style='padding-left:3px'><%=getempname(empid)%></td>
						<td  width="30" style='text-align:center;'><% If Not IsNull(filename) Then response.write "<a href='#' onclick=""download('"&server.URLEncode(filename)&"'); return false;""><img src='/images/m_ppt.gif' width='16' height='16' title='"&reportname&"'>" End If %></td>
					</tr>
				<%
					rs2.movenext
					Loop
				%>
			</table></td>
			</tr>
			<%
						intLoop = intLoop + 1
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
<iframe src="about:blank" name="process" width=0 height=0 frameborder=0></iframe>