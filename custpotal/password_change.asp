<%
	dim userid : userid = request("userid")
	dim password : password = request("password")

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<title>▒SK MARKETING EXCELLENT▒ </title>
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

<body>
<form >
<INPUT TYPE="hidden" NAME="userid" value="<%=userid%>">
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> 비밀번호 변경 </td>
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
	  <table border="0" cellpadding="0" cellspacing="0" align="center">
          <tr>
            <td class="hw">신규 비밀번호</td>
            <td colspan="3" class="bw"><INPUT TYPE="password" NAME="password" style='width:320px;' ></td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">비밀번호 확인</td>
            <td colspan="3" class="bw"><INPUT TYPE="password" NAME="repassword" style='width:320px;'></td>
          </tr>
		  <tr>
		  <tr>
				<td colspan="4"  height="50" valign="bottom" align="right"><img src="/images/btn_save.gif" width="59" height="18" vspace="5" style="cursor:hand" onclick="checkForSubmit()"  hspace="10"> <img src="/images/btn_close.gif" width="57" height="18" vspace="5" style="cursor:hand" onClick="set_close();"   >
	</td>
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
<SCRIPT LANGUAGE="JavaScript">
<!--
	function checkForSubmit() {

		var frm = document.forms[0];
		if (frm.password.value == "") {
			alert("비밀번호를 입력하세요");
			frm.password.focus();
			return false;
		}

		if (frm.password.value.length < 8 ) {
			alert("비밀번호는 8~12자 사이로 입력하세요");
			return false ;
		}
		if (frm.password.value == "<%=password%>") {
			alert("기존과 동일한 비밀번호로 설정할 수 없습니다.");
			return false;
		}
		if (frm.password.value != frm.repassword.value) {
			alert("비밀번호가 일치하지 않습니다.");
			return false ;
		}
		if (frm.password.value == "<%=userid%>"){
			alert("아이디와 동일한 비밀번호는 설정할 수 없습니다..");
			return false ;
		}

		var bln = true;
		var regexp = /^[a-z\d]{8,12}$/i;
		var regexp_str = /[a-z]/i;
		var regexp_num = /[\d]/i;
		if (!(regexp.test(frm.password.value) && regexp_str.test(frm.password.value) && regexp_num.test(frm.password.value))) {
			alert("비밀번호는 영자,숫자의 조합만으로 작성하세요.");
			return false ;
		}

		frm.action = "password_check_proc.asp";
		frm.method = "post";
		frm.submit();
	}

	function set_close() {
		this.close();
	}
//-->
</SCRIPT>