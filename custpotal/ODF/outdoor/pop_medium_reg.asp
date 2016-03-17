<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%

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
}
-->
</style></head>

<body  oncontextmenu="return false">
<form  enctype="multipart/form-data">
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> 매체 등록 </td>
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
				<td class="hw">매체명</td>
				<td class="bw"><input name="txttitle" type="text"  style="width:340px" maxlength="100" class="kor"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">매체사</td>
				<td class="bw"> <% call get_medium_custcode(null, null)%>
				</td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">매체분류</td>
				<td class="bw"><span id="category">조회를 누르신 후 매체분류를 선택하세요.</span> <img src="/images/btn_find.gif" width="39" height="20" align="absmiddle" class="stylelink" onClick="pop_medium_category();"> <input name="txtcategoryidx" type="hidden" id="txtcategoryidx"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">수량단위</td>
				<td class="bw"><input name="rdounit" type="radio" value="구좌" checked onclick="check_valid_unit();"> 구좌 <input name="rdounit" type="radio" value="기" onclick="check_valid_unit();"> 기 <input name="rdounit" type="radio" value="면" onclick="check_valid_unit();"> 면 <input name="rdounit" type="radio" value="기타" onclick="check_valid_unit();"> 직접입력 <input type="text" name="txtunit"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">설치지역</td>
				<td class="bw"><%call get_region_code(null, null) %> </td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">위치정보</td>
				<td class="bw"><input name="txtlocate" type="text"  style="width:340px" maxlength="100" class="kor"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hw">약도파일</td>
				<td class="bw"><input name="txtmap" type="file" id="txtmap"  style="width:340px"> </td>
			</tr>
			<tr>
				<td ></td>
				<td >* 약도는 jpg, gif 형식, 사이즈 500*350 으로 등록하세요.</td>
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
			alert("매체명은 필수입력 사항입니다.");
			frm.txttitle.focus();
			return false;
		}
		if (frm.selcustcode.value == "") {
			alert("매체사는 필수입력 사합니다.");
			frm.selcustcode.focus();
			return false;
		}
		if (frm.txtcategoryidx.value == "") {
			alert("매체분류는 필수입력 사항입니다.");
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