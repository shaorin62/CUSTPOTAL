<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
	dim userid : userid = request.QueryString("userid")

	Dim objrs, sql

	sql = "select userid, username, password, class, isuse "
	sql = sql &  " from wb_account  "
	sql = sql &  " where userid = '" & userid & "'"
	
	Call get_recordset(objrs, sql)

	Dim  username,password, class_ , isuse
	if not objrs.eof Then
		username = objrs("username")
		userid = objrs("userid")
		password = objrs("password")
		class_ = objrs("class")
		isuse= objrs("isuse")
	else
		response.write "<script type='text/javascript'> alert('������ �����̰ų� �߸��� �������̵� �Դϴ�.'); this.close();</script>"
	end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<title>�Ƣ� SK M&C | Media Management System �Ƣ� </title>
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

<body  background="/images/pop_bg.gif" oncontextmenu="return false">
<form>
<table width="540" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> ���� ���� ���� </td>
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
<!--  --><table border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
              <tr>
                <td  class="hw">���̵�</td>
                <td class="bw bbd" ><%=userid%><input type="hidden" name="userid" value="<%=userid%>"></td>
              </tr>
			  	<tr>
					<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
				</tr>
			   <tr>
                <td  class="hw">�̸�</td>
                <td class="bw bbd" ><input  name="username" value="<%=username%>"></td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
              <tr>
                <td  class="hw">��й�ȣ</td>
                <td class="bw bbd"><input name="txtpassword" type="password"  id="txtpassword" value="" maxlength="12"></td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
              <tr>
                <td  class="hw">��й�ȣȮ��</td>
                <td class="bw bbd"><input name="txtrepassword" type="password" id="txtrepassword" value="" maxlength="12"></td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
              <tr>
                <td  class="hw">���ӱ���</td>
                <td class="bw bbd" >
					<input name="rdoclass" type="radio" value="A" onclick="check_auth()" <%if class_ = "A" then response.write " checked"%>>
					<span style="width:50px;">Admin</span>
					<input name="rdoclass" type="radio" value="G" onclick="check_auth()"  <%if class_ = "G" then response.write " checked"%>>
					<span style="width:50px;">MP</span>
                    <input name="rdoclass" type="radio" value="C"  onclick="check_auth()" <%if class_ = "C" then response.write " checked"%>>
					<span style="width:50px;">������</span>
				</td>
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

		var frm = document.forms[0];
		var btnSearch = document.getElementById("btnSearch");
		var chk_rdoclass ;
		for (var i=0; i < frm.rdoclass.length ; i++) {
			if (frm.rdoclass[i].checked) chk_rdoclass = frm.rdoclass[i].value;
		}
		
	}

	function check_submit() {
		var frm = document.forms[0];
		var flag = true ;

		if (frm.txtpassword.value != "" && frm.txtpassword.value.length < 8) {
			alert("��й�ȣ�� ����, ������������ 8~12�ڸ��� �Է��ϼž� �մϴ�.");
			frm.txtpassword.focus();
			return false;
		}
		if (frm.txtpassword.value != frm.txtrepassword.value) {
			alert("��й�ȣ�� �߸��Է��ϼ̽��ϴ�.\n\n��й�ȣ�� ��Ȯ�ϰ� �Է��ϼž� �մϴ�.");
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
				alert("��й�ȣ�� ����,���� ���� 8~12�ڸ��� �ۼ��ϼ���.");
				frm.txtpassword.select();
				return false ;
			}
		}


		for (var i = 0 ; i < frm.rdoclass.length; i++ ) {
			if (frm.rdoclass[i].checked) flag = false ;
		}

		if (flag) {
			alert("������ ������ �����ϼž� �մϴ�.");
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


	function check_auth() {

		var frm = document.forms[0];
		var btnSearch = document.getElementById("btnSearch");

		var chk_rdoclass ;
		for (var i=0; i < frm.rdoclass.length ; i++) {
			if (frm.rdoclass[i].checked) chk_rdoclass = frm.rdoclass[i].value;
		}

	}
//-->
</script>