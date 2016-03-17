<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim idx : idx = request("idx")
	dim contidx
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	dim photo : photo = request("photo")

	dim objrs, sql
	sql = "select m.contidx, title, photo_1, photo_2, photo_3, photo_4 from dbo.wb_contact_mst m inner join dbo.wb_contact_md m2 on m.contidx = m2.contidx inner join dbo.wb_contact_md_dtl d on m2.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on d.idx = a.idx where a.idx = " & idx & " and a.cyear = '" & cyear & "' and a.cmonth = '" & cmonth & "' "
	call get_recordset(objrs, sql)

	dim title, photo_1, photo_2, photo_3, photo_4

	if not objrs.eof then
		contidx = objrs("contidx")
		title = objrs("title")
		photo_1 = objrs("photo_1")
		photo_2 = objrs("photo_2")
		photo_3 = objrs("photo_3")
		photo_4 = objrs("photo_4")
		if photo = photo_1 then photo = photo_1
		if photo = photo_2 then photo = photo_2
		if photo = photo_3 then photo = photo_3
		if photo = photo_4 then photo = photo_4
	end if

	objrs.close

	set objrs = nothing
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

<body background="/images/pop_bg.gif" >
<form>
<table width="640" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> <%=title%>&nbsp;<%=cyear%>.<%=cmonth%> </td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="640" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="22"><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td bgcolor="#FFFFFF">
<!--  -->
<img src="/pds/media/<%=photo%>" width="600" border="1" onclick="set_close();" class="stylelink" alt="이미지를 클릭하시면 창이 닫힙니다.">
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

	function set_close() {
		this.close();
	}

	window.onload = function () {
		self.focus();
	}
//-->
</script>
