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
<title>�Ƣ� SK M&C | Media Management System �Ƣ� </title>
	<link href="../style.css" rel="stylesheet" type="text/css">
 </HEAD>

 <BODY  oncontextmenu="return false">
<FORM>
	  <TABLE>
  <TR>
	<TD><% If objrs.eof Then Response.write userid & "�� ��밡���� ���̵��Դϴ�.<br> ����Ͻðڽ��ϱ�?<P> <span onclick='checkForAccount()'>Ȯ��</span>" Else Response.write userid & "�� ���������� ���̵��Դϴ�.<br>���̵� �ٽ� �˻��ϼ���.<P>" End if%></TD>
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
			alert("���̵𿡴� ������ ����� �� �����ϴ�.");
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