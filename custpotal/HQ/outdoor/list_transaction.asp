<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->
<%
	Dim atag : atag = ""
	Dim pcustcode : pcustcode = request("cmbcustcode")
	Dim pteamcode : pteamcode = request("cmbteamcode")
	Dim cyear : cyear = request("cyear")
	Dim cmonth : cmonth = request("cmonth")
	If cyear = "" Then cyear = Year(date)
	If cmonth = "" Then cmonth = Month(date)
	If Len(cmonth) = 1 Then cmonth = "0"&cmonth
	dim sdate : sdate = dateserial(cyear, cmonth, "01")
	dim edate : edate = dateadd("d", -1, dateadd("m", 1,  sdate))
	pcustcode = clearXSS(pcustcode, atag)
	pteamcode = clearXSS(pteamcode, atag)
	cyear = clearXSS(cyear, atag)
	cmonth = clearXSS(cmonth, atag)

	Dim strstat : strstat = request("cmbSTAT")
	Dim strwhere

	if strstat = "" then
		strwhere = ""
	elseif strstat = "Y" then
		strwhere = " and isnull(e.ishold,'') = 'Y' "
	elseif strstat = "N" then
		strwhere = " and isnull(e.ishold,'') = 'N' "
	elseif strstat = "M" then
		strwhere = " and isnull(e.ishold,'') = '' "
	else
		strwhere = ""
	end if

'	Dim sql : sql = "select a.contidx, a.title, a.firstdate, a.startdate, a.enddate , case isnull(sum(monthly),0) when 0 THEN 0 ELSE  isnull(totalprice, 0) END totalprice, isnull(sum(monthly),0) monthly, isnull(sum(expense), 0) expense, a.custcode, a.flag, case isnull(e.medcode,'') when '' then c.medcode else e.medcode end medcode, e.isHold from wb_contact_mst a inner join sc_cust_dtl b on a.custcode=b.custcode inner join wb_contact_md c on a.contidx=c.contidx inner join wb_contact_exe d on c.mdidx = d.mdidx and d.cyear = '"&cyear&"' and d.cmonth = '"&cmonth&"'  left outer join wb_contact_trans e on a.contidx=e.contidx and c.contidx=e.contidx and e.cyear='"&cyear&"' and e.cmonth='"&cmonth&"'  where a.startdate  <= '"&edate&"' and a.enddate >= '"&sdate&"' and b.highcustcode like '"&pcustcode&"%' and a.custcode like '"&pteamcode&"%' "& strwhere &" group by a.contidx, a.title, a.firstdate, a.startdate, a.enddate ,a.totalprice ,a.custcode ,a.flag, c.medcode, e.medcode, e.isHold order by a.contidx desc"
	
	Dim sql : sql = "select a.contidx, a.title, a.firstdate, a.startdate, a.enddate, "
	sql = sql  & " case isnull(sum(monthly),0) when 0 THEN 0 ELSE  isnull(totalprice, 0) END totalprice, isnull(sum(monthly),0) monthly, isnull(sum(expense), 0) expense, a.custcode, a.flag,"
	sql = sql  & " case isnull(e.medcode,'') when '' then c.medcode else e.medcode end medcode, e.isHold"
	sql = sql  & " from wb_contact_mst a inner join sc_cust_dtl b on a.custcode=b.custcode inner join wb_contact_md c on a.contidx=c.contidx inner join wb_contact_exe d on c.mdidx = d.mdidx"
	sql = sql  & " and d.cyear = '"&cyear&"' and d.cmonth = '"&cmonth&"'  left outer join wb_contact_trans e on a.contidx=e.contidx and c.contidx=e.contidx"
	sql = sql  & " and e.cyear='"&cyear&"' and e.cmonth='"&cmonth&"'  where a.startdate  <= '"&edate&"' and a.enddate >= '"&sdate&"' and b.highcustcode like '"&pcustcode&"%' and a.custcode like '"&pteamcode&"%' "& strwhere &" "
	sql = sql  & " group by a.contidx, a.title, a.firstdate, a.startdate, a.enddate ,a.totalprice ,a.custcode ,a.flag, c.medcode, e.medcode, e.isHold order by a.contidx desc"


'	response.write sql
'
	Dim rs : Set rs = server.CreateObject("adodb.recordset")
	rs.activeconnection = application("connectionstring")
	rs.cursorLocation = adUseClient
	rs.cursorType = adOpenStatic
	rs.lockType = adLockOptimistic
	rs.source = sql
	rs.open

	Dim totalrecord : totalrecord = rs.recordcount

	Dim contidx : Set contidx = rs(0)
	Dim title : Set title = rs(1)
	Dim firstdate : Set firstdate = rs(2)
	Dim startdate : Set startdate = rs(3)
	Dim enddate : Set enddate = rs(4)
	Dim totalprice : Set totalprice = rs(5)
	Dim monthly : Set monthly = rs(6)
	Dim expense : Set expense = rs(7)
	Dim teamcode : Set teamcode = rs(8)
	Dim flag : Set flag = rs(9)
	Dim medcode: Set medcode = rs(10)
	Dim isHold : Set isHold = rs(11)
	'Dim seq : Set seq = rs(12)
	Dim income : income = 0
	Dim incomerate : incomerate = "0.00"

	Dim grandtotalprice : grandtotalprice =  0
	Dim grandmonthly : grandmonthly = 0
	Dim grandexpense : grandexpense = 0
	Dim grandincome : grandincome = 0
	Dim grandincomerate : grandincomerate = 0
	Dim grandprice : grandprice = 0

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

		function getcontact(contidx, flag) {
			// 계약 상세 현황 팝업
			var cyear = "<%=cyear%>";
			var cmonth = "<%=cmonth%>";
			var url = "/hq/outdoor/popup/view_"+flag+"_contact.asp?contidx="+contidx+"&cyear="+cyear+"&cmonth="+cmonth;
			var name = "contactdetail";
			var left = screen.width / 2 - 1024 / 2;
			var top = 10;
			var opt = "width=1260,  resiable=yes, scrollbars=yes, left="+left+", top="+top;
			window.open(url, name, opt);
		}

		function gettoggle() {
			var bln = document.getElementById("toggle").checked;
			var checkElement = document.getElementsByTagName("input");
			for (var i=0; i<checkElement.length;i++) {
				if (checkElement[i].getAttribute("type") == "checkbox") checkElement[i].checked = bln;
			}
		}

		function getAccept() {
			var check = true;
			var hold = false ;
			var accept = false;
			var clickElement = document.getElementsByTagName("input");
			for (var i=0; i<clickElement.length; i++) {
				if (clickElement[i].getAttribute("name") == "contidx") {
						switch (clickElement[i].getAttribute("className")) {
							case "Y":
							case "C":
								if (clickElement[i].checked) hold = true;
								break;
							case "N":
								if (clickElement[i].checked) accept = true;
								clickElement[i].checked = false ;
								break;
							default :
								if (clickElement[i].checked)	check = false;
						}
					}
				}
			if (accept) {alert("이미 정산대기 중인 계약은 정산할 수 없습니다.."); return false;}
			if (hold) {alert("승인 또는 취소 항목은 정산할 수 없습니다."); return false;}
			if (check) {alert('정산할 계약을 선택하세요'); return false;}
			document.getElementById("crud").value = "U"

			submitchange();
		}

		function getAcceptCancel() {
			var check = true;
			var hold = false ;
			var clickElement = document.getElementsByTagName("input");
			for (var i=0; i<clickElement.length; i++) {
				if (clickElement[i].getAttribute("name") == "contidx") {
						switch (clickElement[i].getAttribute("className")) {
							case "Y":
							case "C":
								if (clickElement[i].checked) hold = true;
								break;
							case "A":
								if (clickElement[i].checked) check = true;
								clickElement[i].checked = false;
								break;
							default :
								if (clickElement[i].checked) check = false;
						}
					}
				}
			if (hold) {alert("승인 또는 취소 항목은 정산취소할 수 없습니다."); return false;}
			if (check) {alert('정산취소할  계약을 선택하세요'); return false;}
			document.getElementById("crud").value = "C"

			if (confirm("선택하신 계약을 정산취소 신청하시겠습니까?")) {
				submitchange();
			}
		}

		function getHoldCancel() {
			var check = false;
			var hold = true ;
			var clickElement = document.getElementsByTagName("input");
			for (var i=0; i<clickElement.length; i++) {
				if (clickElement[i].getAttribute("name") == "contidx") {
						switch (clickElement[i].getAttribute("className").toString()) {
							case "Y":
							case "C":
								if (clickElement[i].checked) hold = false;
								break;

							default :
								if (clickElement[i].checked) check = true;
						}
					}
				}
			if (check) {alert('RMS에서 승인되지 않은 계약은 취소할 수 없습니다.'); return false;}
			if (hold) {alert("승인 취소할  계약을 선택하세요."); return false;}
			document.getElementById("crud").value = "H"

			if (confirm("선택한 계약을 RMS 승인취소 신청하시겠습니까?")) {
				submitchange();
			}
		}

		function submitchange() {
			var frm = document.forms[0];
			frm.action = "/hq/outdoor/process/db_transaction.asp";
			frm.method = "post";
			frm.submit();
		}

		function getexcel() {
			// 엑셀전환
		}

		window.onload = function () {
			_sendRequest("/inc/getcustcombo.asp", "custcode=<%=pcustcode%>", _getcustcombo, "GET");
			_sendRequest("/inc/getteamcombo.asp", "custcode=<%=pcustcode%>&teamcode=<%=pteamcode%>", _getteamcombo, "GET");
			document.getElementById("cmbcustcode").attachEvent("onchange", getteamcombo);
		}
//-->
//관리 보고서 출력(추가사항)
function getprint() {

			// 관리 보고서 출력
			var frm = document.forms[0];
			frm.target = "_blank";
			frm.method = "post";
			frm.action = "/hq/outdoor/process/print_report2.asp";
			frm.submit();
			cleanToggle();

		}
function cleanToggle() {
			document.getElementById("toggle").checked = false;
			gettoggle();

		}

</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  >
<form action="list_transaction.asp" method='post' name="form">
<INPUT TYPE="hidden" NAME="menunum" value="<%=request("menunum")%>">
<INPUT TYPE="hidden" NAME="crud" ID='crud' >
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
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="boldtitle"> 광고비용 집행현황 </span></TD>
				<TD width="50%" align="right"><span class="navigator" > 옥외광고 &gt;  옥외광고현황 &gt; 광고비용 집행현황 </span></TD>
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
				  <%call getyear(cyear)%> <%call getmonth(cmonth)%>   <span id="custcode">광고주 검색</span> <span id="teamcode">운영팀 검색</span> <SELECT id="cmbSTAT"  style="WIDTH: 50px" name="cmbSTAT">
										<OPTION value="" <% if strstat = "" then response.write "selected" end if %>>전체</OPTION>
										<OPTION value="M" <% if strstat = "M" then response.write "selected" end if %>>미결</OPTION>
										<OPTION value="N" <% if strstat = "N" then response.write "selected" end if %>>대기</OPTION>
										<OPTION value="Y" <% if strstat = "Y" then response.write "selected" end if %>>승인</OPTION>
									</SELECT> <input type="image" src="/images/btn_search.gif" width="39" height="20" align="absmiddle" border=0></td>
                  <td  align="right" background="/images/bg_search.gif" > 
				  <A HREF="#" onclick="getAccept();  return false;"><img src="/images/btn_execution.gif" width="78" height="18" align="absmiddle" border="0" ></A> 
				  <A HREF="#" onclick='getAcceptCancel();  return false;'><img src="/images/btn_execution_cancel.gif" width="78" height="18" align="absmiddle" border="0" ></A> <!-- <A HREF="#" onclick="getHoldCancel(); return false;"><img src="/images/btn_hold_cancel.gif" width="78" height="18" align="absmiddle" border="0" ></A> --></td>
                  <td width="13" ><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="35" ><img src='/images/unlock.gif' width='16' height='16' alt='정산되지 않은 상태'> 미결 <img src='/images/lock.gif' width='16' height='16' alt='정산 요청 중 RMS에서 정산승인 처리전 상태'> 대기 <img src='/images/hold.gif' width='16' height='16' alt='RMS에서 정산승인 완료' hspace=2> 승인 <!-- <img src='/images/unhold.gif' width='16' height='16' alt='RMS에서 정산취소 대기' hspace=2> 취소 --> &nbsp;&nbsp;&nbsp;&nbsp;
			<a href="#" onclick="getprint(); return false;"><img src='/images/m_print.gif' width='16' height='16' alt='선택한 계약의 관리보고서를 인쇄' align='bottom' > 관리보고서 인쇄  </a></td>
			
			<td align='right'><a href="#"><!-- <img src='/images/icon_xls.gif' width='17' height='16'></a> 엑셀 --> </td>
          </tr>
          <tr>
            <td valign="top" height="516" colspan='2'>

				  <table border="0"width="1120" cellpadding="0" cellspacing="0" bordercolor="#8D652B" id="contact">
				  <thead>
					<tr height='30' align='center'>
						<th class="hd left" width="20">&nbsp;</th>
						<th class="hd center" width="20"><INPUT TYPE="checkbox" NAME="toggle" id='toggle' onclick='gettoggle();'></th>
						<th class="hd center" width="240">매체명</th>
						<th class="hd center" width="70">최초계약</th>
						<th class="hd center" width="70">시작일자</th>
						<th class="hd center" width="70">종료일자</th>
						<th class="hd center" width="90">총광고료</th>
						<th class="hd center" width="90">월광고료</th>
						<th class="hd center" width="90">월지급액</th>
						<th class="hd center" width="80">내수액</th>
						<th class="hd center" width="50">내수율</th>
						<th class="hd center" width="115">광고주</th>
						<th class="hd right" width="110">매체사</th>
					</tr>
					</thead>
					<tbody id='tbody'>
					<%
						Do Until rs.eof
							income = monthly-expense
							If monthly = 0 Then incomerate = "0.00" Else 	incomerate = income/monthly*100
					%>
					<tr height='32'>
						<td  class="hd none" style=' text-align:center;'><% If isHold = "Y" Then response.write "<img src='/images/hold.gif' width='16' height='16' alt='승인완료'>" Else If isHold = "N" Then response.write "<img src='/images/lock.gif' width='16' height='16' alt='정산요청'>" Else If isHolde ="C" Then response.write "<img src='/images/unhold.gif' width='16' height='16' alt='승인취소'>" Else  response.write "<img src='/images/unlock.gif' width='16' height='16' alt='미정산'>" End If %></td>
						<td  class="hd none" style='padding-left:5px; text-align:left;'><INPUT TYPE="checkbox" NAME="contidx"  value="<%=contidx&","&medcode%>" <%If isHold = "Y" Then response.write "class='Y' " Else If isHold="N" Then response.write "class='N'" Else If isHold="C" Then response.write "class='C'" Else response.write "class='A'" End If %>></td>
						<td  class="hd none" style="padding-left:5px;"><a href="#" onclick="getcontact(<%=contidx%>,'<%=flag%>'); return false;" title="<%=title%>" class='subject'><%=cutTitle(title, 38)%></a></td>
						<td  class="hd none" style=' text-align:center;'><%=firstdate%></td>
						<td  class="hd none" style=' text-align:center;'><%=startdate%></td>
						<td  class="hd none" style=' text-align:center;'><%=enddate%></td>
						<td  class="hd none" style='text-align:right;'><%=FormatNumber(totalprice, 0)%>&nbsp;&nbsp;</td>
						<td  class="hd none" style='text-align:right;'><%=FormatNumber(monthly, 0)%>&nbsp;&nbsp;</td>
						<td  class="hd none" style='text-align:right;'><%=formatnumber(expense,0)%>&nbsp;&nbsp;</td>
						<td  class="hd none" style='text-align:right;'><%=formatnumber(income,0)%>&nbsp;&nbsp;</td>
						<td  class="hd none" style='text-align:right;'><%=formatnumber(incomerate,2)%>&nbsp;&nbsp;</td>
						<td  class="hd none" style=''><%=getcustname(teamcode)%></td>
						<td  class="hd none" style=''><%=getmedname(medcode)%></td>
					</tr>
				  <%
							totalrecord = totalrecord - 1
							grandmonthly = CDbl(grandmonthly) + CDbl(monthly)
							grandexpense = CDbl(grandexpense) + CDbl(expense)
							grandtotalprice = CDbl(grandtotalprice) + CDbl(totalprice)
							rs.movenext
						Loop

						grandincome = CDbl(grandmonthly) - CDbl(grandexpense)
						if grandincome = 0 Then grandincomerate = "0.00" else	if grandmonthly = 0 then grandincomerate = "0.00" else grandincomerate = grandincome/grandmonthly *100 end if end if
				  %>
				  </tbody>
				  <tfoot>
                  <tr height="30">
                    <td class="hd left"  colspan='6' style="text-align:center;"><strong>총합계</strong> </td>
                    <td class="hd center" style=' text-align:right; ;font-size:11px;font-family:돋움'><%=formatnumber(grandtotalprice,0)%>&nbsp;</td>
                    <td class="hd center" style=' text-align:right; '><%=formatnumber(grandmonthly,0)%>&nbsp;</td>
                    <td class="hd center" style=' text-align:right; '><%=formatnumber(grandexpense,0)%>&nbsp;</td>
                    <td class="hd center" style=' text-align:right; '><%=formatnumber(grandincome,0)%>&nbsp;</td>
                    <td class="hd center" style=' text-align:right; '><%=formatnumber(grandincomerate,2)%>&nbsp;</td>
                    <td class="hd right" colspan='2'></td>
                  </tr>
				  </tfoot>
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
