<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<%
	' iframe 을 이용하여 프로세스 처리 framename = processFrame

	Dim pcustcode : pcustcode = request("cmbcustcode")
	Dim pteamcode : pteamcode = request("cmbteamcode")

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

	dim sdate : sdate = dateserial(cyear, cmonth, "01")
	dim sdate2 : sdate2 = dateserial(cyear2, cmonth2, "01")
	dim edate : edate = dateadd("d", -1, dateadd("m", 1,  sdate2))

'	response.write cyear&cmonth &"====="
'	response.write cyear2&cmonth2 &"====="
'	response.write sdate&"=====" & edate
	Dim sql : sql = "select c.contidx, c.title, c.firstdate, c.startdate, c.enddate, isnull(c.totalprice,0) as totalprice, isnull(m.monthly,0) as monthly,"
	sql = sql  & " isnull(m.expense,0) as expense, c.custcode as teamcode, d.custname as teamname, d2.custname as custname, c.flag "
	sql = sql  & " from wb_contact_mst c "
	sql = sql  & " left outer join sc_cust_dtl d on c.custcode = d.custcode "
	sql = sql  & " left outer join sc_cust_hdr d2 on d.highcustcode = d2.highcustcode "
	sql = sql  & " left outer join vw_contact_exe_monthly m on m.contidx = c.contidx "
	sql = sql  & " and m.cyear >='"&cyear&"' and m.cyear <='"&cyear2&"' and m.cmonth >='"&cmonth&"' and m.cmonth <='"&cmonth2&"' "
	sql = sql  & " where c.startdate >= '"&sdate&"' and c.enddate <= '"&edate&"' and d.highcustcode like '"&pcustcode&"%' and c.custcode like '"&pteamcode&"%' "
	sql = sql  & " order by contidx desc "

	'response.write sql

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
	Dim teamname : Set teamname = rs(9)
	Dim custname : Set custname = rs(10)
	Dim flag : Set flag = rs(11)
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
		var rows = 0;
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
			if (flag == "" || flag == null) flag = "S";
			var url = "/hq/outdoor/popup/view_"+flag+"_contact.asp?contidx="+contidx+"&cyear="+cyear+"&cmonth="+cmonth;
			var name = "contactdetail";
			var left = screen.width / 2 - 1024 / 2;
			var top = 10;
			var opt = "width=1260,  resiable=yes, scrollbars=yes, left="+left+", top="+top;
			window.open(url, name, opt);
		}

		function _getcontactedit(ary) {
			var tableElement = document.getElementById("contact");
			var rowElement = tableElement.insertRow(1);

			for (var i = 0 ; i < ary.length; i++) {
				var cellElement = rowElement.insertCell(-1);
				cellElement.appendChild(document.createTextNode(ary[i]));
			}
		}

		function getcontactview(contidx, crud) {
			//  계약 팝업
			var custcode = document.getElementById("cmbcustcode").value;
			var teamcode = document.getElementById("cmbteamcode").value;
			var cyear = document.getElementById("cyear").value;
			var cmonth = document.getElementById("cmonth").value;
			var url = "/hq/outdoor/popup/view_contact.asp?contidx="+contidx+"&custcode="+custcode+"&teamcode="+teamcode+"&cyear="+cyear+"&cmonth="+cmonth+"&crud="+crud;
			var name = "contactpop";
			var left = screen.width / 2 - 550 / 2;
			var top = 10;
			var opt = "width=550, height=560, resizable=no, scrollbars=no, status=yes, left="+left+", top="+top;
			var win = window.open(url, name, opt);
			win.focus();
		}

		function getcontactdelete(arg) {
			// 계약 삭제
			if (confirm("선택한 계약에 해당 하는 모든 데이터가 삭제됩니다.\n\n계약을 삭제하시겠습니까?")) {
				processFrame.location.href = "/hq/outdoor/process/db_contact.asp?contidx="+arg+"&custcode=<%=pcustcode%>&teamcode=<%=pteamcode%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>&crud=d";
//				rows = event.srcElement.parentElement.parentElement.parentElement.rowIndex;
//				var params = "contidx="+arg+"&custcode=<%=pcustcode%>&teamcode=<%=pteamcode%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
//				sendRequest("/hq/outdoor/process/db_delete_contact.asp", params, _getcontactdelete, "GET");
			}
		}

		function _getcontactdelete() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
					var debugConsole = document.getElementById("debugConsole");
					debugConsole.innerHTML = xmlreq.responseText;
				}
			}
		}


		function getprint() {
			// 관리 보고서 출력
		}

		function getexcel() {
			// 엑셀전환
			var custcode = document.getElementById("cmbcustcode").value;
			var teamcode = document.getElementById("cmbteamcode").value;
			var cyear = document.getElementById("cyear").value;
			var cmonth = document.getElementById("cmonth").value;
			var cyear2 = document.getElementById("cyear2").value;
			var cmonth2 = document.getElementById("cmonth2").value;

			location.href = "/hq/outdoor/excel/xls_contact_month.asp?custcode="+custcode+"&teamcode="+teamcode+"&cyear="+cyear+"&cmonth="+cmonth+"&cyear2="+cyear2+"&cmonth2="+cmonth2;
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
<form action="list_contact_month.asp" method='post'>
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
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="boldtitle"> 옥외광고 집행현황 </span></TD>
				<TD width="50%" align="right"><span class="navigator" > 옥외광고 &gt;  옥외광고현황 &gt; 옥외광고 집행현황 </span></TD>
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
				  <%call getyear(cyear)%> <%call getmonth(cmonth)%>&nbsp;~&nbsp;<%call getyear2(cyear2)%> <%call getmonth2(cmonth2)%>    <span id="custcode">광고주 검색</span> <span id="teamcode">운영팀 검색</span> <input type="image" src="/images/btn_search.gif" width="39" height="20" align="absmiddle" border=0></td>
                  <td  align="right" background="/images/bg_search.gif" ><img src="/images/btn_excel.gif" width="85" height="22" align="absmiddle" vspace="5" class="stylelink" onclick="getexcel();"></td>
                  <td width="13" ><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="20" > </td>

          </tr>
          <tr>
            <td valign="top" height="516" colspan='2'>

				  <table border="0"width="1030" cellpadding="0" cellspacing="0" bordercolor="#8D652B" id="contact">
				  <thead>
					<tr height='30' align='center'>
						<th class="hd left" width="20">No</th>
						<th class="hd center" width="242">매체명</th>
						<th class="hd center" width="70">최초계약</th>
						<th class="hd center" width="70">시작일자</th>
						<th class="hd center" width="70">종료일자</th>
						<th class="hd center" width="70">총광고료</th>
						<th class="hd center" width="70">월광고료</th>
						<th class="hd center" width="70">월지급액</th>
						<th class="hd center" width="67">내수액</th>
						<th class="hd center" width="47">내수율</th>
						<th class="hd center" width="75">광고주</th>
						<th class="hd right" width="100">운영팀</th>
					</tr>
					</thead>
					<tbody id='tbody'>
					<%
						Do Until rs.eof
							income = monthly-expense
							If monthly = 0 Then incomerate = "0.00" Else 	incomerate = income/monthly*100
					%>
					<tr height='32'>
						<td  class="hd none" style='padding-left:3px; text-align:left;'><span name="totalrecord"><%=totalrecord%></span></td>
						<td  class="hd none" style="padding-left:5px;"><%=cutTitle(title, 40)%></td>
						<td  class="hd none" style=' text-align:center;'><%=firstdate%></td>
						<td  class="hd none" style=' text-align:center;'><%=startdate%></td>
						<td  class="hd none" style=' text-align:center;'><%=enddate%></td>
						<td  class="hd none" style='padding-right:3px; text-align:right;'><%=FormatNumber(totalprice, 0)%></td>
						<td  class="hd none" style='padding-right:3px; text-align:right;'><%=FormatNumber(monthly, 0)%></td>
						<td  class="hd none" style='padding-right:3px; text-align:right;'><%=formatnumber(expense,0)%></td>
						<td  class="hd none" style='padding-right:3px; text-align:right;'><%=formatnumber(income,0)%></td>
						<td  class="hd none" style='padding-right:10px; text-align:right;'><%=formatnumber(incomerate,2)%></td>
						<td  class="hd none" style='padding-left:3px;'><%=custname%></td>
						<td  class="hd none" style='padding-left:3px;'><%=teamname%></td>
					</tr>
				  <%
							totalrecord = totalrecord - 1
							grandmonthly = CDbl(grandmonthly) + CDbl(monthly)
							grandexpense = CDbl(grandexpense) + CDbl(expense)
							grandtotalprice = CDbl(grandtotalprice) + CDbl(totalprice)
							rs.movenext
						Loop

						grandincome = CDbl(grandmonthly) - CDbl(grandexpense)
						if grandincome = 0 Then grandincomerate = "0.00" else	grandincomerate = grandincome/grandmonthly *100
				  %>
				  </tbody>
				  <tfoot>
                  <tr height="30">
                    <td class="hd left"  colspan='6' style="text-align:center;"><strong>총합계</strong> </td>
<!--                     <td class="hd center" style=' text-align:right; font-weight:bold;font-size:11px;'><%=formatnumber(grandtotalprice/1000,0)%>&nbsp;</td> -->
                    <td class="hd center" style=' text-align:right; font-weight:bold'><%=formatnumber(grandmonthly,0)%>&nbsp;</td>
                    <td class="hd center" style=' text-align:right; font-weight:bold'><%=formatnumber(grandexpense,0)%>&nbsp;</td>
                    <td class="hd center" style=' text-align:right; font-weight:bold'><%=formatnumber(grandincome,0)%>&nbsp;</td>
                    <td class="hd center" style=' text-align:right; font-weight:bold'><%=formatnumber(grandincomerate,2)%>&nbsp;</td>
                    <td class="hd right" colspan='2'>&nbsp;</td>
                  </tr>
				  </tfoot>
              </table>
			  <div id="debugConsole"> &nbsp;</div>
			  </td>
          </tr>
      </table></td>
    </tr>
  </table>
</form>
<iframe src='about:blank' name='processFrame' frameborder=0 width='0' height='0'></iframe>
</body>
</html>
<!--#include virtual="/bottom.asp" -->