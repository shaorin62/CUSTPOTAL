<!--#include virtual="/inc/getdbcon.asp" -->

<%
	Dim userid : userid = request("txtaccount")
	dim sql : sql = "DBO.WEB_ACCOUNT"

	dim objrs : set objrs = server.createobject("adodb.recordset")
	objrs.activeconnection = application("connectionstring")
	objrs.cursorlocation = aduseclient
	objrs.cursortype = adopenstatic
	objrs.locktype = adlockreadonly
	objrs.source = sql
	objrs.open

	objrs.find = "USERID = '" & userid & "'"
%>
<HTML>
 <HEAD>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
	<link href="../style.css" rel="stylesheet" type="text/css">
 </HEAD>

 <BODY  oncontextmenu="return false">
<FORM>
	  <TABLE>
  <TR>
	<TD><% If objrs.eof Then Response.write userid & "는 사용가능한 아이디입니다.<br> 사용하시겠습니까?<P> <span onclick='checkForAccount()'>확인</span>" Else Response.write userid & "는 현재사용중인 아이디입니다.<br>아이디를 다시 검색하세요.<P>" End if%></TD>
  </TR>
  <TR>
	<TD><INPUT TYPE="text" NAME="txtaccount" onkeyup="checkForKey();"> <img src="/images/btn_search.gif" width="44" height="22" align="absmiddle" onclick="checkForSubmit();" class="stylelink"></TD>
  </TR>
  </TABLE>
  </FORM>
 </BODY>
</HTML>

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
		frm.action = "checkAccount.asp";
		frm.method = "post";
		frm.submit();
	}
//-->
</SCRIPT>