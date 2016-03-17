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
<form >
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> 매체 분류 삭제 </td>
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
            <td class="hw">대분류</td>
            <td colspan="3" class="bw"><span id="ggroup"><% call get_category_grand(null, null, "get_category_grand()") %></div></td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">중분류</td>
            <td colspan="3" class="bw"><span id="mgroup"><select name="selmcategory" style='width:320px;' ><option value="">중분류를 선택하세요</option></select></span></td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">소분류</td>
            <td colspan="3" class="bw"><span id="sgroup"><select name="selscategory" style='width:320px;' ><option value="">소분류를 선택하세요</option></select></span></td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">세분류</td>
            <td colspan="3" class="bw"><span id="dgroup"><select name="seldcategory" style='width:320px;' ><option value="">세분류를 선택하세요</option></select></span></td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">메뉴명</td>
            <td colspan="3" class="bw"><INPUT TYPE="text" NAME="txtcategoryname" style="width:320px;" readonly></td>
          </tr>
		  <tr>
				<td colspan="4"  height="50" valign="bottom" align="right"> *삭제할 메뉴명을 선택하세요. <img src="/images/btn_delete.gif" width="59" height="18" " vspace="5" style="cursor:hand" onClick="check_submit();" hspace="10" > <img src="/images/btn_close.gif" width="57" height="18" vspace="5" style="cursor:hand" onClick="set_close();"  >
	</td>
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
<iframe id="ifrm" src="about:blank;" width="0" height="0" frameborder="0"></iframe>
</body>
</html>
	<script language="JavaScript" src="/js/script.js"></script>
	<script language="JavaScript">
	<!--
		function check_submit() {
			var frm = document.forms[0];

			if (frm.txtcategoryname.value == "") {
				alert("삭제할 메뉴명을 선택하세요");
				frm.txtcategoryname.focus();
				return false;
			}

			if ((confirm("선택한 메뉴명을 삭제하시겠습니까?")) {
				frm.action = "pop_category_delete_proc.asp";
				frm.method = "post";
				frm.submit();
			}

		}

		function set_reset() {
			document.forms[0].reset();
		}

		function set_close() {
			this.close();
		}

		window.onload=function () {
			self.focus();
		}

		function get_category_grand() {
			var frm = document.forms[0];
			if (frm.selgcategory.selectedIndex == 0) {
				document.getElementById("mgroup").innerHTML = "<select name='selmcategory' style='width:320px'><option value=''>중분류를 선택하세요</option></select>"
				document.getElementById("sgroup").innerHTML = "<select name='selscategory' style='width:320px'><option value=''>중분류를 선택하세요</option></select>"
				document.getElementById("dgroup").innerHTML = "<select name='seldcategory' style='width:320px'><option value=''>중분류를 선택하세요</option></select>"
				frm.txtcategoryname.value = "";
			} else {
				ifrm.location.href="/inc/frm_category_edit_code.asp?ggroupidx="+frm.selgcategory.options[frm.selgcategory.selectedIndex].value;
				frm.txtcategoryname.value = frm.selgcategory.options[frm.selgcategory.selectedIndex].text ;
			}
		}

		function get_category_middle() {
			var frm = document.forms[0];
			if (frm.selmcategory.selectedIndex == 0) {
				document.getElementById("sgroup").innerHTML = "<select name='selscategory' style='width:320px'><option value=''>중분류를 선택하세요</option></select>"
				document.getElementById("dgroup").innerHTML = "<select name='seldcategory' style='width:320px'><option value=''>중분류를 선택하세요</option></select>"
				frm.txtcategoryname.value = "";
			} else {
				ifrm.location.href="/inc/frm_category_edit_code.asp?ggroupidx="+frm.selgcategory.options[frm.selgcategory.selectedIndex].value+"&mgroupidx="+frm.selmcategory.options[frm.selmcategory.selectedIndex].value;
				frm.txtcategoryname.value = frm.selmcategory.options[frm.selmcategory.selectedIndex].text ;
			}
		}

		function get_category_small() {
			var frm = document.forms[0];
			if (frm.selscategory.selectedIndex == 0) {
				document.getElementById("dgroup").innerHTML = "<select name='seldcategory' style='width:320px'><option value=''>중분류를 선택하세요</option></select>"
				frm.txtcategoryname.value = "";
			} else {
				ifrm.location.href="/inc/frm_category_edit_code.asp?ggroupidx="+frm.selgcategory.options[frm.selgcategory.selectedIndex].value+"&mgroupidx="+frm.selmcategory.options[frm.selmcategory.selectedIndex].value+"&sgroupidx="+frm.selscategory.options[frm.selscategory.selectedIndex].value;
				frm.txtcategoryname.value = frm.selscategory.options[frm.selscategory.selectedIndex].text ;
			}
		}

		function get_category_detail() {
			var frm = document.forms[0];
			if (frm.seldcategory.selectedIndex == 0) {
				frm.txtcategoryname.value = "";
			} else {
				ifrm.location.href="/inc/frm_category_edit_code.asp?ggroupidx="+frm.selgcategory.options[frm.selgcategory.selectedIndex].value+"&mgroupidx="+frm.selmcategory.options[frm.selmcategory.selectedIndex].value+"&sgroupidx="+frm.selscategory.options[frm.selscategory.selectedIndex].value+"&dgroupidx="+frm.seldcategory.options[frm.seldcategory.selectedIndex].value;
				frm.txtcategoryname.value = frm.seldcategory.options[frm.seldcategory.selectedIndex].text ;
			}
		}
	//-->
	</script>
