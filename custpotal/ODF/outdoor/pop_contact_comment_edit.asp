<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim contidx : contidx = request("contidx")
	dim objrs, sql
	sql = "select title, regionmemo, mediummemo, comment from dbo.wb_contact_mst  where contidx = " & contidx
	call get_recordset(objrs, sql)

	dim title : title = objrs("title")
	dim regionmemo : regionmemo = objrs("regionmemo")								'지역특성 변수
	dim mediummemo : mediummemo = objrs("mediummemo")						'매체특성
	dim comment : comment = objrs("comment")											'특이사항
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
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> <%=title%> 특성 관리 </td>
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
	  <input type="hidden" name="contidx" value="<%=contidx%>">
	  <table border="0" cellpadding="0" cellspacing="0" align="center">
          <tr>
            <td class="tdhd s4">매체특성</td>
            <td colspan="3" class="tdbd s8"><textarea name="txtmediummemo" rows="5"  style="width:412;"><%if not isnull(mediummemo) then response.write mediummemo%></textarea></td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">지역특성</td>
            <td colspan="3" class="tdbd s8"><textarea name="txtregionmemo" rows="5" style="width:412;"><%if not isnull(regionmemo) then response.write regionmemo%></textarea></td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="tdhd s4">특이사항</td>
            <td colspan="3" class="tdbd s8"><textarea name="txtcomment" rows="5"  style="width:412;"><%if not isnull(comment) then response.write comment%></textarea></td>
          </tr>
		  <tr>
				<td colspan="4"  height="50" valign="bottom" align="right"> <img src="/images/btn_save.gif" width="59" height="18" " vspace="5" style="cursor:hand" onClick="check_submit();"  hspace="10"><img src="/images/btn_init.gif" width="67" height="18" vspace="5" style="cursor:hand" onClick="set_reset();"><img src="/images/btn_close.gif" width="57" height="18" vspace="6" style="cursor:hand" onClick="set_close();" hspace="10" >
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
</body>
</html>
	<script language="JavaScript" src="/js/script.js"></script>
	<script language="JavaScript">
	<!--
		function check_submit() {
			var frm = document.forms[0];
			frm.action = "pop_contact_comment_edit_proc.asp";
			frm.method = "post";
			frm.submit();

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
	//-->
	</script>
