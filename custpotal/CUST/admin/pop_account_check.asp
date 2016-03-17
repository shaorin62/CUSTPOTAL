<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
	Dim userid : userid = request("txtaccount")
	dim sql , objrs
	sql = "select userid from dbo.wb_account where userid = '" & userid & "' "
	Call get_recordset(objrs, sql)

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
	background-image: url(/images/pop_bg.gif);
}
-->
</style></head>

<body  oncontextmenu="return false">
<form>
<table width="540" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> 계정 중복 조회 </td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="540" border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
  <tr>
    <td width="22"><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td>
<!--  -->
<TABLE>
  <TR>
	<TD  class="tdbd"  align="center" style="width:476" colspan="2"><% If objrs.eof Then Response.write "<h3>"&userid & "</h3> 는 사용가능한 아이디입니다.<br> 사용하시겠습니까?<br> <span ><img src='/images/btn_confirm.gif' width='57' height='18' vspace='6' style='cursor:hand' onclick='checkForAccount()' hspace='10' ></span>" Else Response.write  "<h3>"&userid & "</h3> 는 현재 사용 중인 아이디입니다.<br>아이디를 다시 검색하세요.<P>" End if%>
	</TD>
  </TR>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
  <TR>
	<TD class="tdbd" align="center" style="width:476" colspan="2"><INPUT TYPE="text" NAME="txtaccount" onkeyup="checkForKey();"> <img src="/images/btn_search.gif" width="39" height="20" align="absmiddle" onclick="checkForSubmit();" class="stylelink"></TD>
  </TR>
              <tr>
                <td  height="50" align="left" valign="bottom">&nbsp;</td>
                <td  align="right" valign="bottom"><img src="/images/btn_close.gif" width="57" height="18" vspace="6" style="cursor:hand" onClick="set_close();" hspace="10" ></td>
              </tr>
  </TABLE>
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
<SCRIPT LANGUAGE="JavaScript" SRC="/js/admin.js"></SCRIPT>
<SCRIPT LANGUAGE="JavaScript">
<!--
	function checkForAccount() {
		var winFrm = window.opener.document.forms[0]
		winFrm.checkAccount.value = "Y";
		winFrm.txtaccount.value = "<%=userid%>";
		this.close();
	}

	function checkForSubmit() {
		var frm = document.forms[0];
		if (frm.txtaccount.value == "" || frm.txtaccount.value.indexOf(" ") != -1) {
			alert("아이디에는 공백을 사용할 수 없습니다.");
			frm.txtaccount.value = "";
			frm.txtaccount.focus();
			return false;
		}
		frm.action = "pop_account_check.asp";
		frm.method = "post";
		frm.submit();
	}

	function set_close() {
		this.close();
	}
//-->
</SCRIPT>