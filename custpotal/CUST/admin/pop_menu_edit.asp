<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
	dim midx : midx = request("midx")

	dim objrs, sql
	sql = " select title, isfile, iscomment, isemail, isuse from dbo.wb_menu_mst where midx = " & midx

	call get_recordset(objrs, sql)

	dim title, custcode, custcode2, file, comment, email, custname, custname2, isuse, lvl, subtitle
	if not objrs.eof then
		title = objrs("title")
		file = objrs("isfile")
		comment = objrs("iscomment")
		email = objrs("isemail")
		isuse = objrs("isuse")
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
</style>
</head>

<body  oncontextmenu="return false">
<form>
<input type="hidden" name="midx" value="<%=midx%>">
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> <%=title%> </td>
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
			<tr height="31">
				<td class="hw">�޴���</td>
				<td class="bw bbd"><INPUT TYPE="text" NAME="txttitle" style="width:350px" size="50" value="<%=title%>"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr height="31">
				<td class="hw">�޴����</td>
				<td class="bw bbd"><INPUT TYPE="checkbox" NAME="chkfile" <%if file then response.write " checked"%>> ����÷��<INPUT TYPE="checkbox" NAME="chkemail" <%if email then response.write " checked"%>> ���Ϲ߼� <INPUT TYPE="checkbox" NAME="chkcomment" <%if comment then response.write " checked"%>> ����ۼ� </td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
            <tr>
                  <td  height="50" valign="bottom"><img src="/images/space.gif" width="57" height="20" border="0"></td>
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
	function set_close() {
		this.close();
	}

	function set_reset() {
		document.forms[0].reset();
	}

	function check_submit() {
		var frm = document.forms[0];
		if (frm.txttitle.value == "") {
			alert("�޴����� �ʼ��Է� �׸��Դϴ�.");
			frm.txttitle.focus();
			return false ;
		}
		frm.action = "menu_edit_proc.asp";
		frm.method = "post";
		frm.submit();
	}

	window.onload = function () {
		self.focus();

	}
//-->
</script>