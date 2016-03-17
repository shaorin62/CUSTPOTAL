<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim mdidx : mdidx = request("mdidx")
	dim title : title = request("title")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<title><%=title%> 면별 정보 등록 : SK MARKETING & COMPANY </title>
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
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> <%=title%> 면별 정보 등록 </td>
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
	<td class="hw">면</td>
	<td class="bw"><%call get_side_code(null)%></td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#E7E7DE" height="1"></td>
  </tr>
  <tr>
	<td class="hw">규격</td>
	<td class="bw"><input type="text" name="txtstandard"> 가로 * 세로, "(인치), 동영상(상영시간) </td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#E7E7DE" height="1"></td>
  </tr>
  <tr>
	<td class="hw">재질</td>
	<td class="bw"><%call get_quality_code(null)%></td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#E7E7DE" height="1"></td>
  </tr>
  <tr>
	<td class="hw">월단가(원)</td>
	<td class="bw"><input type="text" name="txtunitprice" class="number" value="0"></td>
  </tr>
  <tr>
    <td colspan="2"  height="50" valign="bottom" align="right"> <img src="/images/btn_save.gif" width="59" height="18" vspace="5" onclick="check_submit();" style="cursor:hand"><img src="/images/btn_init.gif" width="67" height="18" vspace="5" style="cursor:hand" onClick="set_reset();" hspace="10" ><img src="/images/btn_close.gif" width="57" height="18" vspace="5" style="cursor:hand" onClick="set_close();" >
	</td>
  </tr>
  </table>
  <input type="hidden" name="mdidx" value="<%=mdidx%>">
  <input type="hidden" name="txttitle" value="<%=title%>">
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
	function check_submit() {
		var frm = document.forms[0];
		if (frm.txtstandard.value == "") {
			alert("규격정보는 필수입력입니다.");
			frm.txtstandard.focus();
			return false;
		}
		frm.action = "pop_side_reg_proc.asp";
		frm.method = "post";
		frm.submit();
	}

	function set_reset() {
		document.forms[0].reset();
	}

	function set_close() {
		this.close();
	}

	window.onload = function() {
		self.focus();
	}
//-->
</script>
