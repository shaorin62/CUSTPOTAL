<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>

<body  oncontextmenu="return false">
<form>
<table width="522" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> 신규 계약 등록 </td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="540" border="0" cellspacing="0" cellpadding="0">
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
				<td colspan="2" bgcolor="#cacaca" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd">계약명</td>
				<td class="tdbd s7"><input name="txttitle" type="text" id="txttitle"maxlength="100"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd">최초계약일</td>
				<td class="tdbd s7"><input name="txtfirstdate" type="text" id="txtfirstdate" maxlength="10"> <img src="/images/calendar.gif" width="39" height="20" border="0" align="absmiddle" onclick="Calendar_D(document.all.txtfirstdate)"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd">시작일</td>
				<td class="tdbd s7"><input name="txtstartdate" type="text" id="txtstartdate"> <img src="/images/calendar.gif" width="39" height="20" border="0" align="absmiddle" onclick="Calendar_D(document.all.txtstartdate)"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd">종료일</td>
				<td class="tdbd s7"><input name="txtenddate" type="text" id="txtenddate"> <img src="/images/calendar.gif" width="39" height="20" border="0" align="absmiddle" onclick="Calendar_D(document.all.txtenddate)"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
<!-- 			<tr>
				<td class="tdhd">총광고료(원)</td>
				<td class="tdbd"><input name="txttotalprice" type="text" id="txttotalprice" value="0" class="number" onfocus="initNum(this);" onblur="initZero(this);"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd">월광고료(원)</td>
				<td class="tdbd"><input name="txtmonthprice" type="text" id="txtmonthprice" value="0" class="number" onfocus="initNum(this);" onblur="initZero(this);"> <input type="checkbox" name="chktrust" onclick="set_trust_text();"> 위수탁인 경우 선택하세요.</td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd">월외주비(원)</td>
				<td class="tdbd"><input type="text" name="txtexpense" id="txtmonthpay"  value="0" class="number" onfocus="initNum(this);" onblur="initZero(this);"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd"><span id="account">내수액(원)</span></td>
				<td class="tdbd"><input name="txtincome" type="text" id="txtincome" value="0" class="number" onfocus="initNum(this);" onblur="initZero(this);"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd"><span id="ratio">내수율(%)</span></td>
				<td class="tdbd"><input name="txtincomeratio" type="text" id="txtincomeratio" value="0" class="number" onfocus="initNum(this);" onblur="initZero(this);"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr> -->
			<tr>
				<td class="tdhd">사업부</td>
				<td class="tdbd s7"><input type="text" name="txtdeptname" id="txtdeptname" readonly size="35">  <input name="txtdeptcode" type="hidden" id="txtdeptcode" size="10" readonly> <img src="/images/btn_find.gif" width="39" height="20" align="absmiddle" class="stylelink" onClick="get_pop_custcode();"> </td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd">광고주</td>
				<td class="tdbd s7"><input type="text" name="txtcustname" id="txtcustname" readonly size="35"> <input name="txtcustcode" type="hidden" id="txtcustcode" size="10" readonly>
                  </td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd">지역특성</td>
				<td class="tdbd s7"><textarea name="txtregionmemo" rows="5"  style="width:410px" id="txtregionmemo"></textarea></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd">매체특성</td>
				<td class="tdbd s7"><textarea name="txtmediummemo" rows="5"  style="width:410px" id="txtmediummemo"></textarea></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd">특이사항</td>
				<td class="tdbd s7"><textarea name="txtcomment" rows="5"  style="width:410px" id="txtcomment"></textarea></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			</table>
			  <table width="756" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="50%" height="50" align="left" valign="bottom"><a href="/hq/outdoor/md_list.asp"><img src="/images/btn_list.gif" width="59" height="20" border="0"></a></td>
                  <td width="50%" align="right" valign="bottom"><img src="/images/btn_save.gif" width="59" height="20" hspace="10" vspace="5" style="cursor:hand" onClick="check_submit();"><img src="/images/btn_init.gif" width="67" height="20" vspace="5" style="cursor:hand" onClick="set_reset();"></td>
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
</body>
</html>

<script type="text/javascript" src="/js/calendar.js"></script>
<script type="text/javascript" src="/js/script.js"></script>
<script language="javascript">
<!--
	function check_submit() {
		var frm = document.forms[0];
		if (frm.txttitle.value == "") {
			alert("계약명을 입력하세요.") ;
			frm.txttitle.focus();
			return false;
		}
		if (frm.txtstartdate.value == "") {
			alert("계약시작일을 입력하세요.") ;
			frm.txtstartdate.focus();
			return false;
		}
		if (frm.txtenddate.value == "") {
			alert("계약종료일을 입력하세요.");
			frm.txtenddate.focus();
			return false;
		}
//		if (frm.txttotalprice.value == "") {
//			alert("총 계약금을 입력하세요.") ;
//			frm.txttotalprice.focus();
//			return false;
//		}
//		if (frm.txtmonthprice.value == "") {
//			alert("월계약금을 입력하세요.");
//			frm.txtmonthprice.focus();
//			return false;
//		}
//		if (frm.txtexpense.value == "") {
//			alert("월외주비를 입력하세요.");
//			frm.txtexpense.focus();
//			return false;
//		}
		if (frm.txtcustcode.value == "") {
			alert("사업부를  입력하세요.");
			frm.txtcustcode.focus();
			return false;
		}

		frm.method = "POST";
		frm.action = "contact_reg_proc.asp";
		frm.submit();
	}

	function get_pop_custcode() {
		var url = "pop_custcode.asp";
		var name = "pop_custcode";
		var opt = "width=500, height=500, resziable=no, scrollbars = yes, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function set_reset() {
		document.forms[0].reset();
	}

	function set_trust_text() {
		var frm = document.forms[0];
		var bln = frm.chktrust.checked ;
		var account = document.getElementById("account")
		var ratio = document.getElementById("ratio")
		if (bln) {
			frm.txtmonthpay.disabled = bln ;
			account.innerHTML = "수수료(원)";
			ratio.innerHTML = "수수료율(%)" ;

		} else {
			frm.txtmonthpay.disabled = bln ;
			account.innerHTML = "내수액(원)";
			ratio.innerHTML = "내수율(%)" ;
		}
		frm.txtmonthpay.value = "0";
		frm.txtincome.value = "0";
		frm.txtincomeratio.value = "0";
	}

	window.onload = function () {

	}
//-->
</script>