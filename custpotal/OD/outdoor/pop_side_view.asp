<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim sidx : sidx = request("sidx")
	dim title : title = request("title")
	dim objrs, sql
	sql = "select sidx, mdidx, side, standard, quality, unitprice from dbo.WB_MEDIUM_DTL where sidx = " & sidx
	call get_recordset(objrs, sql)

	dim side, standard, quality, unitprice
	if not objrs.eof then
		side = objrs("side")
		standard = objrs("standard")
		quality = objrs("quality")
		unitprice = objrs("unitprice")
	else
		if objrs.eof then response.write "<script> alert('삭제 또는 잘못된 면 정보 입니다.');  this.close(); </script>"
	end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<title>▒▒ SK M&C | Media Management System ▒▒  </title>
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
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> <%=title%> 면별 정보 </td>
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
	<td class="bw"><%=side%></td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#E7E7DE" height="1"></td>
  </tr>
  <tr>
	<td class="hw">규격(M)</td>
	<td class="bw"><%=replace(standard, chr(34), """")%>  </td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#E7E7DE" height="1"></td>
  </tr>
  <tr>
	<td class="hw">재질</td>
	<td class="bw"><%=quality%></td>
  </tr>
  <tr>
    <td colspan="2" bgcolor="#E7E7DE" height="1"></td>
  </tr>
  <tr>
	<td class="hw">월단가(원)</td>
	<td class="bw"><%=unitprice%></td>
  </tr>
  <tr>
    <td colspan="2"  height="50" valign="bottom" align="right"> <img src="/images/btn_edit.gif" width="59" height="18" vspace="5" border="0" class="stylelink" onClick="go_side_edit();"><img src="/images/btn_delete.gif" width="59" height="18" vspace="5" border="0" class="stylelink" onClick="go_side_delete();" hspace="10" ><img src="/images/btn_close.gif" width="57" height="18" vspace="5" style="cursor:hand" onClick="set_close();" >
  </tr>
  </table>
  <input type="hidden" name="sidx" value="<%=sidx%>">
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
	function go_side_edit() {
		location.href = "pop_side_edit.asp?sidx=<%=sidx%>&title=<%=title%>";
	}

	function go_side_delete() {
		if (confirm("면 정보를 삭제하시겠습니까?"))
			location.href = "pop_side_delete_proc.asp?sidx=<%=sidx%>";
		return false;
	}

	function set_close() {
		this.close();
	}

	window.onload = function () {
		self.focus();
	}
//-->
</script>
