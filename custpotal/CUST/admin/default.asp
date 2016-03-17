<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
	dim objrs, sql
%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  oncontextmenu="return false">

<form target="scriptFrame">
<!--#include virtual="/cust/top.asp" -->
<input type="hidden" name="actionurl" value="account.asp">
<input type="hidden" name="tcustcode">
  <table id="Table_01" width="1240" height="600" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/cust/left_admin_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_admin.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td height="600" align="left" valign="top"><table width="1030" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="19">&nbsp;</td>
          </tr>
          <tr>
            <td height="17"><TABLE  width="100%">
            <TR>
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle" id="subtitle"> 계정관리 </span></TD>
				<TD width="50%" align="right"><span class="navigator" id="navi">관리모드 &gt; 계정관리 </span></TD>
            </TR>
            </TABLE>
			</td>
          </tr>
          <tr>
            <td height="15">&nbsp;</td>
          </tr>
          <tr>
            <td ><table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
                  <td width="50%" align="left" background="/images/bg_search.gif"><span id="searchsection"><input type="text" name="txtsearchstring"> <img src="/images/btn_search.gif" width="39" height="20" align="top" class="styleLink" onClick="checkForSearch(document.forms[0].txtsearchstring.value)"></span></td>
                  <td width="50%" align="right" background="/images/bg_search.gif"><img src="/images/btn_acc_reg.gif" width="78" height="18" alt="" border="0" class="account" onclick="pop_reg();" id="btnReg" style="cursor:hand;"></td>
                  <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="15" >&nbsp;</td>
          </tr>
          <tr>
            <td ><iframe src="account.asp" width="1032" height="625" frameborder="2" border="2" name="scriptFrame" id="scriptFrame"></iframe></td>
          </tr>
      </table></td>
    </tr>
  </table>
<!--#include virtual="/bottom.asp" -->
</body>
</html>
<script language="JavaScript">
<!--
	function checkForView(uid) {
		var url = "pop_account_view.asp?userid=" + uid;
		var name = "pop_account_view";
		var opt = "width=540, height=296, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}
	function pop_reg() {
		var p = document.getElementById("btnReg") ;
		var custcode = document.forms[0].tcustcode.value.replace("null","") ;
		if (p.getAttribute("class") == "account" || p.getAttribute("class") == null) {
			var url = "pop_account_reg.asp?tcustcode="+custcode;
			var name = "pop_reg";
			var opt = "width=540, height=366, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		} else {
			var url = "pop_menu_reg.asp?tcustcode="+custcode;
			var name ="pop_menu_reg" ;
			var opt = "width=540, height=205, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		}
		window.open(url, name, opt);
	}

	function checkForSearch(str) {
		var frm = document.forms[0];
		if (str !="") {
			if (str.indexOf("--") != -1) {
				alert("사용할 수 없는 문자를 입력하셨습니다.");
				frm.txtsearchstring.value = "";
				frm.txtsearchstring.focus();
				return false;
			}
		}
		frm.action = frm.actionurl.value;
		frm.method = "post";
		frm.submit();
	}
//-->
</script>