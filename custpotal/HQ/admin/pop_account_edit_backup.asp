<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
	dim def_custcode : def_custcode = "A00058"
	dim def_custname : def_custname = "SK마케팅앤컴퍼니"
	dim userid : userid = request.QueryString("userid")

	Dim objrs, sql
	'dim sql : sql = "select a.userid,a.username, a.password, a.custcode, c.custname, a.class, a.isuse from dbo.wb_account a left outer join dbo.sc_cust_temp c on a.custcode = c.custcode where a.userid = '" & userid & "'"

	sql = "select a.userid, a.username, a.password, a.custcode, case when class= 'C' or class= 'G' then  c.custname else b.custname end  as custname,  a.class, a.isuse "
	sql = sql &  " from wb_account a  inner join sc_cust_dtl b on a.custcode=b.custcode  "
	sql = sql &  " inner join sc_cust_hdr c on b.highcustcode=c.highcustcode  where a.userid = '" & userid & "'"


	Call get_recordset(objrs, sql)

	Dim  username,password, class_, custcode, custname , isuse
	if not objrs.eof Then
		username = objrs("username")
		userid = objrs("userid")
		password = objrs("password")
		class_ = objrs("class")
		custcode = objrs("custcode")
		custname = objrs("custname")
		isuse= objrs("isuse")
		if custcode = "" then custcode = null
	else
		response.write "<script type='text/javascript'> alert('삭제된 계정이거나 잘못된 계정아이디 입니다.'); this.close();</script>"
	end if
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
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> 계정 정보 변경 </td>
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
                <td  class="hw">아이디</td>
                <td class="bw" ><%=userid%><input type="hidden" name="userid" value="<%=userid%>"></td>
              </tr>
			  	<tr>
					<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
				</tr>
			   <tr>
                <td  class="hw">이름</td>
                <td class="bw" ><input  name="username" value="<%=username%>"></td>
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
                <td class="bw" >
					<input name="rdoclass" type="radio" value="A" onclick="check_auth()" <%if class_ = "A" then response.write " checked"%>>
					<span style="width:80px;">Admin</span>
					<input name="rdoclass" type="radio" value="N" onclick="check_auth()"  <%if class_ = "N" then response.write " checked"%>>
					<span style="width:100px;">Admin(Non-SKT)</span><br>
                    <input name="rdoclass" type="radio" value="C"  onclick="check_auth()" <%if class_ = "C" then response.write " checked"%>>
					<span style="width:80px;">광고주</span>
                    <input name="rdoclass" type="radio" value="G"  onclick="check_auth()" <%if class_ = "G" then response.write " checked"%>>
					<span style="width:80px;">광고주 관리자</span><br>
                    <input name="rdoclass" type="radio" value="D"  onclick="check_auth()" <%if class_ = "D" then response.write " checked"%>>
					<span style="width:80px;">팀</span>
                    <input name="rdoclass" type="radio" value="H"  onclick="check_auth()" <%if class_ = "H" then response.write " checked"%>>
					<span style="width:80px;">팀 관리자</span><br>
                    <input name="rdoclass" type="radio" value="F"  onclick="check_auth()" <%if class_ = "F" then response.write " checked"%>>
					<span style="width:80px;">옥외 모니터링</span>
                    <input name="rdoclass" type="radio" value="O"  onclick="check_auth()" <%if class_ = "O" then response.write " checked"%>>
					<span style="width:80px;">옥외 관리자</span><br>
				</td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
              <tr>
                <td  class="hw">계정소속</td>
                <td class="bw"><input name="txtcustcode" type="hidden" id="txtcustcode"  readonly value="<%=custcode%>">
                  <input name="txtcustname" type="text" class="kor" id="txtcustname" size="30" readonly value="<%=custname%>"> <img src="/images/btn_search.gif" width="39" height="20" border="0" alt="" align="absmiddle" class="stylelink" onclick="pop_custcode()" id="btnSearch"> </td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
              <tr>
                <td  height="50" align="left" valign="bottom"><img src="/images/space.gif" width="12" height="20" border="0"></td>
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
		resizeTo(550, 421);
		var frm = document.forms[0];
		var btnSearch = document.getElementById("btnSearch");
		var chk_rdoclass ;
		for (var i=0; i < frm.rdoclass.length ; i++) {
			if (frm.rdoclass[i].checked) chk_rdoclass = frm.rdoclass[i].value;
		}
		switch (chk_rdoclass) {
			case "A":
			case "N":
			case "O":
				btnSearch.style.display = "none";
				break;
			case "C":
			case "D":
			case "M":
			case "G":
			case "H":
				btnSearch.style.display = "";
				break;
			case "F":
				btnSearch.style.display = "none";
				frm.txtcustname.value = "옥외 모니터링 업체";
				break;
		}
	}

	function check_submit() {
		var frm = document.forms[0];
		var flag = true ;

		if (frm.txtpassword.value != "" && frm.txtpassword.value.length < 8) {
			alert("비밀번호를 영문, 숫자형식으로 8~12자리로 입력하셔야 합니다.");
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
		if (frm.txtpassword.value != "") {
			if (!(regexp.test(frm.txtpassword.value) && regexp_str.test(frm.txtpassword.value) && regexp_num.test(frm.txtpassword.value))) {
				alert("비밀번호는 영자,숫자 조합 8~12자리로 작성하세요.");
				frm.txtpassword.select();
				return false ;
			}
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
		frm.action = "account_edit_proc.asp";
		frm.submit();
	}

	function set_close() {
		this.close();
	}

	function set_reset() {
		var frm = document.forms[0].reset();
		return false;
	}

	function pop_custcode() {
		var frm = document.forms[0];
		var chk_rdoclass ;
		for (var i=0; i < frm.rdoclass.length ; i++) {
			if (frm.rdoclass[i].checked) chk_rdoclass = frm.rdoclass[i].value;
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
			case "D":	// 사업부서
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

		var chk_rdoclass ;
		for (var i=0; i < frm.rdoclass.length ; i++) {
			if (frm.rdoclass[i].checked) chk_rdoclass = frm.rdoclass[i].value;
		}



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