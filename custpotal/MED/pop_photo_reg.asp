<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim idx : idx = request("idx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<title>SK MARKETING & COMPANY </title>
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
<form enctype="multipart/form-data" >
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="cyear" value="<%=cyear%>">
<input type="hidden" name="cmonth" value="<%=cmonth%>">
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> 매체 사진 등록  </td>
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
	  <table border="0" cellpadding="0" cellspacing="0" align="center">
          <tr>
            <td class="hw">사진설명</td>
            <td  class="bw"><input type="text" name="txttitle" style="width:350" maxlength="100"></td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">사진첨부</td>
            <td  class="bw"><input type="file" name="txtfile" style="width:350"></td>
          </tr>
          <tr>
            <td class="hw">comment</td>
            <td  class="bw"><TEXTAREA NAME="txtcomment" ROWS="2"  style="width:350"></TEXTAREA></td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">사진첨부</td>
            <td  class="bw"><input type="file" name="txtfile" style="width:350"></td>
          </tr>
          <tr>
            <td class="hw">comment</td>
            <td  class="bw"><TEXTAREA NAME="txtcomment" ROWS="2"  style="width:350"></TEXTAREA></td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">사진첨부</td>
            <td  class="bw"><input type="file" name="txtfile" style="width:350"></td>
          </tr>
          <tr>
            <td class="hw">comment</td>
            <td  class="bw"><TEXTAREA NAME="txtcomment" ROWS="2"  style="width:350"></TEXTAREA></td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">사진첨부</td>
            <td  class="bw"><input type="file" name="txtfile" style="width:350"></td>
          </tr>
          <tr>
            <td class="hw">comment</td>
            <td  class="bw"><TEXTAREA NAME="txtcomment" ROWS="2"  style="width:350"></TEXTAREA></td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
		  <tr>
				<td colspan="4"  height="50" valign="bottom" align="right"><img src="/images/btn_save.gif" width="59" height="18"  vspace="5" style="cursor:hand" onClick="check_submit();"><img src="/images/btn_init.gif" width="67" height="18" vspace="5" style="cursor:hand" onClick="set_reset();" hspace="10"><img src="/images/btn_close.gif" width="57" height="18" vspace="5" style="cursor:hand" onClick="set_close();" ></td>
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
	function check_submit() {
		var frm = document.forms[0];
		var bln = true;

		for (var x = 0 ; x < frm.txtfile.length; x++) {
			if (frm.txtfile(x).value != "") bln = false;
		}

		if (bln) {
			alert("한개 이상의 사진을 등록하셔야 합니다.");
			return false;
		}
		frm.action = "photo_reg_proc.asp";
		frm.method="post";
		frm.submit();
	}

	function set_reset() {
		document.forms[0].reset();
	}

	function set_close() {
		this.close();
	}

	window.onload = function () {
		self.focus();
	}
//-->
</script>