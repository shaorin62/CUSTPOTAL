<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<% dim mdidx : mdidx = request("mdidx") %>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
</head>
<body border="0" cellpadding="0" cellspacing="0"   oncontextmenu="return false">
<form>
<table>
  <tr>
	<td colspan="2" bgcolor="#cacaca" height="1"></td>
  </tr>
  <tr>
	<td class="tdhd">면</td>
	<td class="tdbd s"><%call get_side_code(null)%></td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#E7E7DE" height="1"></td>
  </tr>
  <tr>
	<td class="tdhd">규격(M)</td>
	<td class="tdbd s"><input type="text" name="txtstandard"> (가로 * 세로)</td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#E7E7DE" height="1"></td>
  </tr>
  <tr>
	<td class="tdhd">재질</td>
	<td class="tdbd s"><%call get_quality_code(null)%></td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#E7E7DE" height="1"></td>
  </tr>
  <tr>
	<td class="tdhd">월단가(원)</td>
	<td class="tdbd s"><input type="text" name="txtunitprice" class="number" value="0"></td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#E7E7DE" height="1"></td>
  </tr>
  <tr>
    <td colspan="2"  height="50" valign="bottom" align="right"> <img src="/images/btn_save.gif" width="59" height="20" hspace="10" vspace="5" style="cursor:hand" onClick="check_submit();"><img src="/images/btn_init.gif" width="67" height="20" vspace="5" style="cursor:hand" onClick="set_reset();">
	</td>
  </tr>
  </table>
  <input type="hidden" name="mdidx" value="<%=mdidx%>">
  </form>
 </body>
</html>
<script language="JavaScript">
<!--
	function check_submit() {
		var frm = document.forms[0];
		if (frm.txtstandard.value == "") {
			alert("규격정보를 입력하세요");
			frm.txtstandard.focus();
			return false;
		}
		frm.action = "side_reg_proc.asp";
		frm.method = "post";
		frm.submit();
	}

	function set_reset() {
		document.forms[0].reset();
	}

	window.onload = function () {
		self.focus();
	}
//-->
</script>
