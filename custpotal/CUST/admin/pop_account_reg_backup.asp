<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
	dim def_custcode : def_custcode = "A00058"
	dim def_custname : def_custname = "SK마케팅앤컴퍼니"
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
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> 신규 계정 등록 </td>
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
                <td  class="hw" >아이디</td>
                <td class="bw"><input name="txtaccount" type="text" id="txtaccount" maxlength="12"> <img src="/images/btn_overlap.gif" width="57" height="20" border="0" alt="" align="absmiddle" class="stylelink" onclick="pop_account_check()" id="btnAccount"></span><INPUT TYPE="hidden" NAME="checkAccount" value="N"></td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
                <td  class="hw" >이름</td>
                <td class="bw"><input name="txtname" type="text" id="txtname" maxlength="5"> </td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
              <tr>
                <td  class="hw">비밀번호</td>
                <td class="bw"><input name="txtpassword" type="password"  id="txtpassword" value="" maxlength="12"></td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
              <tr>
                <td  class="hw">비밀번호확인</td>
                <td class="bw"><input name="txtrepassword" type="password" id="txtrepassword" value="" maxlength="12"></td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
              <tr>
                <td  class="hw">접속권한</td>
                 <td class="bw">
					<input name="rdoclass" type="radio" value="A" onclick="check_auth()" checked><span style="width:80px;">Admin</span>
					<input name="rdoclass" type="radio" value="N" onclick="check_auth()" ><span style="width:100px;">Admin(Non-SKT)</span><br>
                    <input name="rdoclass" type="radio" value="C"  onclick="check_auth()"><span style="width:80px;">광고주</span>
                    <input name="rdoclass" type="radio" value="G"  onclick="check_auth()"><span style="width:100px;">광고주 관리자</span><br>
                    <input name="rdoclass" type="radio" value="D"  onclick="check_auth()"><span style="width:80px;">팀</span>
                    <input name="rdoclass" type="radio" value="H"  onclick="check_auth()"><span style="width:100px;">팀관리자</span><br>
                    <input name="rdoclass" type="radio" value="F"   onclick="check_auth()"><span style="width:80px;">옥외 모니터링</span>
                    <input name="rdoclass" type="radio" value="O"  onclick="check_auth()"><span style="width:100px;">옥외 관리자</span><br>
                  </td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
              <tr>
                <td  class="hw">계정소속</td>
                <td class="bw"><input name="txtcustcode" type="hidden" id="txtcustcode" value="" readonly>
                  <input name="txtcustname" type="text" class="kor" id="txtcustname" value="" size="30" readonly> <img src="/images/btn_search.gif" width="39" height="20" border="0" alt="" align="absmiddle" class="stylelink" onclick="pop_custcode()" id="btnSearch"> </td>
              </tr>
              <tr>
                <td  height="50" align="left" valign="bottom"><img src="/images/space.gif" width="59" height="20" border="0"></td>
                <td  align="right" valign="bottom"><img src="/images/btn_save.gif" width="59" height="18" hspace="10" vspace="5" style="cursor:hand" onClick="check_submit();"><img src="/images/btn_init.gif" width="67" height="18" vspace="5" style="cursor:hand" onClick="set_reset();"><img src="/images/btn_close.gif" width="57" height="18" vspace="5" style="cursor:hand" onClick="set_close();" hspace="10" ></td>
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

<script language="JavaScript">
<!--
	window.onload = function init() {
		self.focus();
				document.forms[0].txtcustcode.value = "<%=def_custcode%>";
				document.forms[0].txtcustname.value = "<%=def_custname%>";
				document.getElementById("btnSearch").style.display = "none";
	}

	function check_submit() {
		var frm = document.forms[0];
		var flag = true ;
		var chk_rdoclass ;
		for (var i=0; i < frm.rdoclass.length ; i++) {
			if (frm.rdoclass[i].checked) chk_rdoclass = frm.rdoclass[i].value;
		}


		if (chk_rdoclass != "M"){
			if (frm.txtaccount.value == "" ) {
				alert("계정을 입력하셔야 합니다.");
				frm.txtaccount.focus();
				return false;
			}

			if (frm.checkAccount.value == "N") {
				alert("아이디 중복검사를 않하셨습니다.");
				pop_account_check();
				return false;
			}

		}
		if (frm.txtname.value == "") {
			alert("이름을 입력하셔야 합니다.");
			frm.txtname.focus();
			return false;
		}

		if (frm.txtpassword.value == "" || frm.txtpassword.value.length < 8) {
			alert("비밀번호를 영문, 숫자 조합으로 8~12자리로 입력하셔야 합니다.");
			frm.txtpassword.focus();
			return false;
		}
		if (frm.txtpassword.value != frm.txtrepassword.value) {
			alert("비밀번호를 잘못입력하셨습니다.\n\n비밀번호를 정확하게 입력하셔야 합니다.");
			frm.txtrepassword.value = "";
			frm.txtrepassword.focus();
			return false;
		}

		var bln = true;
		var regexp = /^[a-z\d]{8,12}$/i;
		var regexp_str = /[a-z]/i;
		var regexp_num = /[\d]/i;
		if (!(regexp.test(frm.txtpassword.value) && regexp_str.test(frm.txtpassword.value) && regexp_num.test(frm.txtpassword.value))) {
			alert("비밀번호는 영자,숫자 조합 8~12자리 이상으로 작성하세요.");
			frm.txtpassword.select();
			return false ;
		}

		if (frm.txtcustname.value == "")
		{
			alert("계정소속은 필수입력입니다.");
			return false;
		}

		for (var i = 0 ; i < frm.rdoclass.length; i++ ) {
			if (frm.rdoclass[i].checked) flag = false ;
		}

		if (flag) {
			alert("계정의 권한을 선택하셔야 합니다.");
			frm.rdoclass[0].focus();
			return false;
		}
		frm.method = "POST";
		frm.action = "account_reg_proc.asp";
		frm.submit();
	}

	function pop_account_check() {
		var userid = document.forms[0].txtaccount;
		if (userid.value == "") {
			alert("아이디를 먼저 입력하셔야 합니다.");
			userid.focus();
			return false;
		}
		if (userid.value.length < 4) {
			alert("아이디는 4~12자리로입력하셔야 합니다");
			userid.focus();
			return false;
		}
		var str = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
		var firstword = userid.value.toUpperCase().substring(0,1);
		if (str.indexOf(firstword) == -1) {
			alert("아이디는 영문자로만 시작할 수 있습니다.");
			userid.focus();
			return false;
		}

		var url = "pop_account_check.asp?txtaccount="+userid.value ;
		var name = "pop_account_check";
		var opt = "width=540, height=296, resizable=no, top=100, left=660";
		window.open(url, name, opt);
	}


	function set_close() {
		this.close();
	}

	function set_reset() {
		var frm = document.forms[0].reset();
		check_auth();
		return false;
	}

	function pop_custcode() {
		var frm = document.forms[0];
		var chk_rdoclass ;
		for (var i=0; i < frm.rdoclass.length ; i++) {
			if (frm.rdoclass[i].checked) chk_rdoclass = frm.rdoclass[i].value;
		}
		if (chk_rdoclass == undefined) {
			alert("접속권한을 먼저 선택하세요");
			return false;
		}
		var url ;

		switch (chk_rdoclass) {
			case "A":
			case "N":
			case "F":
			case "O":
				break;
			case "C":	// 광고주
			case "G":
				url = "pop_custcode.asp";
				break;
			case "D":	// 팀
			case "H":
				url = "pop_timcode.asp";
				break;
			case "M":	// 매체사
				url = "pop_real_medcode.asp";
				break;
		}

		var name = "pop_custcode";
		var opt = "width=540, height=396, resizable=no, scrollbars=yes, top=100, left=660";
		window.open(url, name, opt);
	}

	function check_auth() {

		var frm = document.forms[0];
		var btnSearch = document.getElementById("btnSearch");
		var txtaccount = document.getElementById("txtaccount");
		var btnAccount = document.getElementById("btnAccount");


		var chk_rdoclass ;
		for (var i=0; i < frm.rdoclass.length ; i++) {
			if (frm.rdoclass[i].checked) chk_rdoclass = frm.rdoclass[i].value;
		}

		btnAccount.style.display = "";
		txtaccount.style.display = "";

		switch (chk_rdoclass) {

			case "A":
			case "N":
				frm.txtcustcode.value = "<%=def_custcode%>";
				frm.txtcustname.value = "<%=def_custname%>";
				btnSearch.style.display = "none";
				break;
			case "C":
			case "G":
				frm.txtcustcode.value = "";
				frm.txtcustname.value = "";
				btnSearch.style.display = "";
				break;
			case "D":
			case "H":
				frm.txtcustcode.value = "";
				frm.txtcustname.value = "";
				btnSearch.style.display = "";
				break;
			case "F":
				frm.txtcustcode.value = null;
				frm.txtcustname.value = "옥외 모니터링 외주업체";
				btnSearch.style.display = "none";
				break;
//			case "M":
//				frm.txtaccount.value = "";
//				frm.txtcustcode.value = "";
//				btnAccount.style.display = "none";
//				txtaccount.style.display = "none";
//				frm.txtcustcode.value = "";
//				frm.txtcustname.value = "";
//				btnSearch.style.display = "";
//				break;
			case "O":
				frm.txtcustcode.value = "<%=def_custcode%>";
				frm.txtcustname.value = "<%=def_custname%>";
				btnSearch.style.display = "none";
				break;
		}
	}
//-->
</script>