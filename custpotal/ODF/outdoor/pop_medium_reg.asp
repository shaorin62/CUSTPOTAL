<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%

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

<body  oncontextmenu="return false">
<form  enctype="multipart/form-data">
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> ��ü ��� </td>
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
				<td class="hw">��ü��</td>
				<td class="bw"><input name="txttitle" type="text"  style="width:340px" maxlength="100" class="kor"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">��ü��</td>
				<td class="bw"> <% call get_medium_custcode(null, null)%>
				</td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">��ü�з�</td>
				<td class="bw"><span id="category">��ȸ�� ������ �� ��ü�з��� �����ϼ���.</span> <img src="/images/btn_find.gif" width="39" height="20" align="absmiddle" class="stylelink" onClick="pop_medium_category();"> <input name="txtcategoryidx" type="hidden" id="txtcategoryidx"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">��������</td>
				<td class="bw"><input name="rdounit" type="radio" value="����" checked onclick="check_valid_unit();"> ���� <input name="rdounit" type="radio" value="��" onclick="check_valid_unit();"> �� <input name="rdounit" type="radio" value="��" onclick="check_valid_unit();"> �� <input name="rdounit" type="radio" value="��Ÿ" onclick="check_valid_unit();"> �����Է� <input type="text" name="txtunit"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">��ġ����</td>
				<td class="bw"><%call get_region_code(null, null) %> </td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">��ġ����</td>
				<td class="bw"><input name="txtlocate" type="text"  style="width:340px" maxlength="100" class="kor"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">�൵����</td>
				<td class="bw"><input name="txtmap" type="file" id="txtmap"  style="width:340px"> </td>
			</tr>
			<tr>
				<td ></td>
				<td >* �൵�� jpg, gif ����, ������ 500*350 ���� ����ϼ���.</td>
			</tr>
             <tr>
              <td width="50%" height="50" align="left" valign="bottom"><img src="/images/space.gif" width="59" height="20" border="0"></td>
               <td width="50%" align="right" valign="bottom"><img src="/images/btn_save.gif" width="59" height="18" hspace="10" vspace="5" style="cursor:hand" onClick="check_submit();"><img src="/images/btn_init.gif" width="67" height="18" vspace="5" style="cursor:hand" onClick="set_reset();"><img src="/images/btn_close.gif" width="57" height="18" vspace="6" style="cursor:hand" onClick="set_close();" hspace="10" ></td>
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
<script language="javascript">
<!--
	window.onload = function () {
		var frm = document.forms[0];
		frm.txtunit.disabled = true;
	}
	function check_submit() {
		var frm = document.forms[0];
		if (frm.txttitle.value == "") {
			alert("��ü���� �ʼ��Է� �����Դϴ�.");
			frm.txttitle.focus();
			return false;
		}
		if (frm.selcustcode.value == "") {
			alert("��ü��� �ʼ��Է� ���մϴ�.");
			frm.selcustcode.focus();
			return false;
		}
		if (frm.txtcategoryidx.value == "") {
			alert("��ü�з��� �ʼ��Է� �����Դϴ�.");
			pop_medium_category();
			return false;
		}

		frm.method = "POST";
		frm.action = "medium_reg_proc.asp";
		frm.submit();
	}

	function pop_medium_custcode() {
		var url = "pop_medium_custcode.asp";
		var name = "pop_medium_custcode";
		var opt = "width=500, height=500, resziable=no, scrollbars = yes, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function pop_medium_category() {
		var url = "pop_medium_category.asp";
		var name = "pop_medium_category";
		var opt = "width=540, height=525, resziable=no, scrollbars = yes, status=yes, top=100, left=660";
		window.open(url, name, opt);
	}

	function check_valid_unit() {
		var frm = document.forms[0];
		var bln = frm.rdounit[3].checked ;
		frm.txtunit.disabled = !bln;
		if (bln) {
			frm.txtunit.focus();
			frm.txtunit.value = "";
		}
	}

	function set_reset() {
		document.forms[0].reset();
	}

	function set_close() {
		this.close();
	}
//-->
</script>