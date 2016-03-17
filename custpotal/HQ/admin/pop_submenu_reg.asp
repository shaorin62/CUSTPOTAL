<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
	dim midx : midx = request("midx")
	dim FLAG : FLAG = request("tflag")
	dim gubun : gubun = request("gubun")
	dim mc : mc = request("mc")
	if FLAG = "" then FLAG = Null

	dim objrs, sql
	sql = " select title, custcode, lvl from dbo.wb_menu_mst where midx = " & midx

	call get_recordset(objrs, sql)

	dim title , custcode, lvl
	if not objrs.eof then
		title = objrs("title")
		custcode = objrs("custcode")
		lvl = objrs("lvl")
	end if
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
</style>
</head>

<body  oncontextmenu="return false">
<form>
<input type="hidden" name="midx" value="<%=midx%>">
<input type="hidden" name="custcode" value="<%=custcode%>">
<input type="hidden" name="lvl" value="<%=lvl%>">
<input type="hidden" name="FLAG" value="<%=FLAG%>">
<input type="hidden" name="gubun" value="<%=gubun%>">
<input type="hidden" name="mc" value="<%=mc%>">
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> <%=title%> 하위메뉴 등록</td>
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
				<td class="hw">메뉴명</td>
				<td class="bw bbd"><INPUT TYPE="text" NAME="txttitle" style="width:350px" size="50"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr height="31">
				<td class="hw">메뉴기능</td>
				<td class="bw bbd"><INPUT TYPE="checkbox" NAME="chkfile"> 파일첨부<INPUT TYPE="checkbox" NAME="chkemail"> 메일발송 <INPUT TYPE="checkbox" NAME="chkcomment"> 댓글작성 <INPUT TYPE="checkbox" NAME="chkMP"> 내부(MP)용</td>
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
			alert("메뉴명은 필수입력 항목입니다.");
			frm.txttitle.focus();
			return false ;
		}
		frm.action = "submenu_reg_proc.asp";
		frm.method = "post";
		frm.submit();
	}

	function go_page(url) {
		var frm = document.forms[0];
		frm.action = url;
		frm.method = "post";
		frm.submit();
	}
	window.onload = function () {
		self.focus();
	}
//-->
</script>