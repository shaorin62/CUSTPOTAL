<!--#include virtual="/MP/outdoor/inc/Function.asp" -->
<%
	Dim pcontidx : pcontidx = Request("contidx")
	If pcontidx = "" Then pcontidx = 0
	Dim custcode : custcode = Request("custcode")
	Dim teamcode : teamcode = Request("teamcode")
	Dim orgcustcode : orgcustcode = custcode
	Dim orgteamcode : orgteamcode = teamcode
	Dim pcyear : pcyear = request("cyear")
	Dim pcmonth : pcmonth =request("cmonth")
	Dim crud : crud = request("crud")

	Dim sql : sql = " select c.title, c.firstdate, c.startdate, c.enddate, c.custcode, d.highcustcode, c.comment, c.regionmemo, c.mediummemo, c.flag, d2.custcode deptcode  from wb_contact_mst c inner join  sc_cust_dtl d on c.custcode = d.custcode left outer join sc_cust_dtl d2 on d.clientsubcode = d2. custcode where c.contidx = " & pcontidx

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	Dim rs : Set rs = cmd.Execute
	Set cmd = Nothing

	If Not rs.eof Then
		Dim title : title = rs(0)
		Dim firstdate : firstdate = Replace(rs(1),"-","")
		Dim startdate : startdate = Replace(rs(2),"-","")
		Dim enddate : enddate = Replace(rs(3),"-","")
		teamcode = rs(4)
		custcode = rs(5)
		Dim comment : comment = rs(6)
		Dim regionmemo : regionmemo = rs(7)
		Dim mediummemo : mediummemo = rs(8)
		Dim flag : flag = rs(9)
		Dim deptcode : deptcode = rs(10)

	End If

	If title = "" Then
		toptitle = "신규 계약 등록"
	Else
		If UCase(crud)="E" Then
			toptitle = title & " 재계약"
		Else
			toptitle = title
		End If
	End If
%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<link href="/MP/outdoor/style.css" rel="stylesheet" type="text/css">
	<title>▒▒ SK M&C | Media Management System ▒▒  </title>
	<script type='text/javascript' src='/js/ajax.js'></script>
	<script type='text/javascript' src='/js/script.js'></script>
    <script type="text/javascript" src="/js/calendar.js"></script>
	<script type="text/javascript">
	<!--
		var scope = "global";

		function getcustcombo() {
			var crud = "<%=crud%>";
			if (crud == "U") scope = null;

			var custcode = "<%=custcode%>" ;
			var params = "scope="+scope+"&custcode="+custcode;
			sendRequest("/inc/getcustcombo.asp", params, _getcustcombo, "GET");
		}

		function _getcustcombo() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
						var custcode = document.getElementById("custcode");
						custcode.innerHTML = xmlreq.responseText ;
						document.getElementById("cmbcustcode").attachEvent("onchange", getdeptcombo);
						getdeptcombo();
				}
			}
		}

		function getdeptcombo() {
			var custcode = document.getElementById("cmbcustcode").value;
			var deptcode = "<%=deptcode%>";
			var params = "custcode="+custcode+"&deptcode="+deptcode;
			sendRequest("/inc/getdeptcombo.asp", params, _getdeptcombo, "GET");
		}

		function _getdeptcombo() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
					var deptcode = document.getElementById("deptcode");
					deptcode.innerHTML = xmlreq.responseText ;
					document.getElementById("cmbdeptcode").attachEvent("onchange", getteamcombo);
					getteamcombo();
				}
			}
		}

		function getteamcombo() {
			var custcode = document.getElementById("cmbcustcode").value;
			var deptcode = document.getElementById("cmbdeptcode").value;
			var teamcode = "<%=teamcode%>" ;
			var params = "custcode="+custcode+"&deptcode="+deptcode+"&teamcode="+teamcode;
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

		function getmedemployee() {
			var medcode = document.getElementById("cmbmed").value;
			var params = "medcode="+medcode ;
			sendRequest("/inc/getmedemployee.asp", params, _getmedemployee, "GET");
		}

		function _getmedemployee() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
						var employeeview = document.getElementById("employeeview");
						employeeview.innerHTML = xmlreq.responseText ;
				}
			}
		}


		function submitchange() {
			var frm = document.forms[0];

			if (frm.txttitle.value.replace(/\s/g, "") == "") {
				alert("계약매체명을 입력하세요");
				frm.txttitle.focus();
				return false;
			}
			if (frm.txtfirstdate.value.replace(/\s/g, "") == "") {
				alert("최초계약일을 입력하세요");
				frm.txtfirstdate.focus();
				return false;
			}
			if (frm.txtfirstdate.value.length != 8) {
				alert("일자형식은 8자리(20090101)로 입력하세요");
				frm.txtfirstdate.focus();
				return false;
			}
			if (frm.txtstartdate.value.replace(/\s/g, "") == "") {
				alert("계약시작일을 입력하세요");
				frm.txtstartdate.focus();
				return false;
			}
			if (frm.txtstartdate.value.length != 8) {
				alert("일자형식은 8자리(20090101)로 입력하세요");
				frm.txtstartdate.focus();
				return false;
			}
			if (frm.txtenddate.value.replace(/\s/g, "") == "") {
				alert("계약종료일을 입력하세요");
				frm.txtenddate.focus();
				return false;
			}
			if (frm.txtenddate.value.length != 8) {
				alert("일자형식은 8자리(20090101)로 입력하세요");
				frm.txtenddate.focus();
				return false;
			}
			if (frm.cmbteamcode.selectedIndex == 0) {
				alert("운영팀을 선택하세요");
				frm.cmbteamcode.focus();
				return false;
			}
			frm.target = "processFrm";
			frm.action = "/MP/outdoor/process/db_contact.asp";
			frm.method = "post";
			frm.submit();
		}

		function checkNumber(p) {
			if (isNaN(p.value)) {
				alert("숫자만 입력하실수 있습니다.");
				p.value = "";
			}
		}

		window.onload = function () {
			self.focus();
			_sendRequest("/inc/getcustcombo.asp", "scope="+scope+"&custcode=<%=custcode%>", _getcustcombo, "GET");
			_sendRequest("/inc/getdeptcombo.asp","custcode=<%=custcode%>&deptcode=<%=deptcode%>", _getdeptcombo, "GET");
			_sendRequest("/inc/getteamcombo.asp", "custcode=<%=custcode%>&deptcode=<%=deptcode%>&teamcode=<%=teamcode%>", _getteamcombo, "GET");
			document.getElementById("cmbcustcode").attachEvent("onchange", getteamcombo);
			var crud = "<%=crud%>";
			var flag = "<%=flag%>";
			if (flag == "S") document.getElementById("small").checked = true;
			else document.getElementById("big").checked = true;

			if (crud!="c") {
				document.getElementById("big").disabled = true;
				document.getElementById("small").disabled = true;
			}
		}
	//-->
	</script>
</head>

<body>
<form>
<input type="hidden" id="contidx" name="contidx" value="<%=pcontidx%>" />
<input type="hidden" id="cyear" name="cyear" value="<%=pcyear%>" />
<input type="hidden" id="cmonth" name="cmonth" value="<%=pcmonth%>" />
<input type="hidden" id="hdnflage" name="hdnflage" value="<%=flag%>" />
<input type="hidden" id="crud" name="crud" value="<%=crud%>" />
<input type="hidden" id="orgcustcode" name="orgcustcode" value="<%=orgcustcode%>" />
<input type="hidden" id="orgteamcode" name="orgteamcode" value="<%=orgteamcode%>" />
<table width="550" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="absmiddle"> <%=toptitle%>   </td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="550" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="22"><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td>
<!--  -->
			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td class="hdr h" width='95'>계약명</td>
				<td class="sc" width='400'><input name="txttitle" type="text" id="txttitle"maxlength="100" style="width:370px" style="ime-mode:active;" value="<%=title%>"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">최초계약일</td>
				<td class="sc"><input name="txtfirstdate" type="text" id="txtfirstdate" maxlength="8"  class="dt" value="<%=firstdate%>" onkeyup="checkNumber(this); return false;"> <a href="#"  onclick="Calendar_D(document.all.txtfirstdate);  return false;" ><img src="/images/calendar.gif" width="39" height="20"  align="absmiddle" ></a></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">계약기간</td>
				<td class="sc"><input name="txtstartdate" type="text" id="txtstartdate" class="dt" maxlength="8"  value="<%=startdate%>"><%If crud="U" Then%><img src="/images/calendar.gif" width="39" height="20"  align="absmiddle"><% Else %> <a href="#" onclick="Calendar_D(document.all.txtstartdate); return false;"><img src="/images/calendar.gif" width="39" height="20"  align="absmiddle"></a><% End If %> ~ <input name="txtenddate" type="text" id="txtenddate"  class="dt" maxlength="8"  value="<%=enddate%>"> <a href="#"  onclick="Calendar_D(document.all.txtenddate);return false;" ><img src="/images/calendar.gif" width="39" height="20"  align="absmiddle"></a></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">광고주</td>
				<td class="sc"><div id="custcode"></div></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">사업부 </td>
				<td class="sc"> <div id='deptcode'></div> </td>
			</tr>
			<tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h"> 운영팀</td>
				<td class="sc"> <div id='teamcode'></div></td>
			</tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
 			<tr>
				<td class="hdr h">지역특성</td>
				<td class="sc" style="padding-top:3px; padding-bottom:3px;"><textarea name="txtregionmemo"  id="txtregionmemo" onclick="checktextlength(this, 1000); return false;" onkeyup="checktextlength(this, 1000); return false;" class="reg"><%=regionmemo%></textarea></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">매체특성</td>
				<td class="sc" style="padding-top:3px; padding-bottom:3px;"><textarea name="txtmediummemo" id="txtmediummemo"  onclick="checktextlength(this, 1000); return false;" onkeyup="checktextlength(this, 1000); return false;" class="reg"><%=mediummemo%></textarea></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">특이사항</td>
				<td class="sc" style="padding-top:3px; padding-bottom:3px;" ><textarea name="txtcomment" id="txtcomment" onclick="checktextlength(this, 200); return false;" onkeyup="checktextlength(this, 200); return false;" class="reg"><%=comment%></textarea></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">매체구분 </td>
				<td class="sc"> <input type="radio" id="big" name='rdoflag' value="B" checked/> <span title="매체면으로 관리하는 광고">대형 (야립,옥탑) </span><input type="radio" id="small" name='rdoflag' value="S" /> <span title="매체단위로 관리하는 광고">소형 </span></td>
			</tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
                  <td  align="right" valign="bottom" width='495' height='50'>  <a href="#" onclick="submitchange(); return false;"><img src="/images/btn_save.gif" width="59" height="18"  vspace="5" border=0 hspace=10></a><a href="#" onclick="window.close(); return false;"><img src="/images/btn_close.gif" width="57" height="18" vspace="6" border=0 ></a><br>* 저장을 끝내시려면 닫기 버튼을 누르세요.</td>
                </tr>
			</table>
<!--  -->
	</td>
    <td background="/images/pop_right_middle_bg.gif">&nbsp;</td>
  </tr>
  <tr>
    <td><img src="/images/pop_left_bottom_bg.gif" width="22" height="25"></td>
    <td background="/images/pop_center_bottom_bg.gif">&nbsp;</td>
    <td><img src="/images/pop_right_bottom_bg.gif" width="23" height="25"></td>
  </tr>
</table>
</form>
<iframe src='about:blank' name='processFrm' width='500' height='600' frameborder="0"></iframe>
</body>
</html>
<div id='debugConsole'></div>