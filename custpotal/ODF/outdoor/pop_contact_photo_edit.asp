<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim contidx : contidx = request("contidx")
	dim sidx : sidx = request("sidx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")

	dim objrs, sql
	sql = "select title from dbo.wb_contact_mst  where contidx = " & contidx
	call get_recordset(objrs, sql)

	dim title : title = objrs("title")

	sql = "select photo_1, photo_2, photo_3, photo_4 from dbo.wb_contact_md_dtl  where contidx = " & contidx &" and sidx = "&sidx&" and cyear = "&cyear&" and cmonth = "&cmonth
	call get_recordset(objrs, sql)

	dim photo_1 : photo_1 = objrs("photo_1")
	dim photo_2 : photo_2 = objrs("photo_2")
	dim photo_3 : photo_3 = objrs("photo_3")
	dim photo_4 : photo_4 = objrs("photo_4")

	objrs.close
	set objrs = nothing
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
<form enctype="multipart/form-data">
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> <%=title%>  </td>
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
	  <input type="hidden" name="sidx" value="<%=sidx%>">
	  <input type="hidden" name="cyear" value="<%=cyear%>">
	  <input type="hidden" name="cmonth" value="<%=cmonth%>">
	  <table border="0" cellpadding="0" cellspacing="0" align="center">
          <tr>
            <td class="hw">매체사진1</td>
            <td  class="bw"><input type="file" name="txtphoto_1" style="width:340;"><input type="hidden" name="txtphoto1" value="<%=photo_1%>"></td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">매체사진2</td>
            <td  class="bw"><input type="file" name="txtphoto_2" style="width:340;"><input type="hidden" name="txtphoto2" value="<%=photo_2%>"></td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">매체사진3</td>
            <td  class="bw"><input type="file" name="txtphoto_3" style="width:340;"><input type="hidden" name="txtphoto3" value="<%=photo_3%>"></td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">매체사진4</td>
            <td  class="bw"><input type="file" name="txtphoto_4" style="width:340;"><input type="hidden" name="txtphoto4" value="<%=photo_4%>"></td>
          </tr>
          <tr>
          <tr>
            <td > </td>
            <td height="20">*  오른쪽부터 순차적으로 등록됩니다.<BR> * 사진 용량은 150K이하, 사이즈는 가로 500px 이하로 등록하세요.</td>
          </tr>
		  <tr>
				<td colspan="4"  height="50" valign="bottom" align="right"> <img src="/images/btn_save.gif" width="59" height="18"  vspace="5" style="cursor:hand" onClick="check_submit();"  hspace="10"><img src="/images/btn_init.gif" width="67" height="18" vspace="5" style="cursor:hand" onClick="set_reset();"><img src="/images/btn_close.gif" width="57" height="18" vspace="6" style="cursor:hand" onClick="set_close();" hspace="10" >
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
	<script language="JavaScript" src="/js/calendar.js"></script>
	<script language="JavaScript" src="/js/script.js"></script>
	<script language="JavaScript">
	<!--
		function check_submit() {
			var frm = document.forms[0];
			frm.action = "pop_contact_photo_edit_proc.asp";
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
