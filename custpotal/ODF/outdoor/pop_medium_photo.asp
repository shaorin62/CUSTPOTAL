<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim sidx : sidx = request("sidx")
	dim contidx : contidx = request("contidx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	dim photo : photo = request("photo")

	dim objrs, sql
	sql = "select title, photo_1, photo_2, photo_3, photo_4 from dbo.wb_contact_md m inner join dbo.wb_contact_md_dtl d on m.contidx = d.contidx where m.contidx = " & contidx & " and m.sidx = " & sidx & " and cyear = '" & cyear & "' and cmonth = '" & cmonth & "'"

	call get_recordset(objrs, sql)

	dim title, photo_1, photo_2, photo_3, photo_4
	if not objrs.eof then
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
<title>�Ƣ� SK M&C | Media Management System �Ƣ� </title>
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
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> <%=title%>&nbsp;<%=cyear%>.<%=cmonth%>. </td>
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
<img src="/pds/media/<%=photo%>" width="600" border="1" onclick="set_close();" class="stylelink" alt="�̹����� Ŭ���Ͻø� â�� �����ϴ�.">
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
<script type="text/javascript" src="/js/calendar.js"></script>
<script language="javascript">
<!--
	function go_page(url) {
		var frm = document.forms[0];
		frm.action = "pop_job_reg.asp";
		frm.method = "post";
		frm.submit();
	}

	function check_submit() {
		var frm = document.forms[0];
		if (frm.selcustcode.value == "") {
			alert("�����ָ� �����ϼ���");
			frm.selcustcode.focus();
			return false;
		}

		if (frm.selseqno.value == "") {
			alert("�귣�带 �����ϼ���");
			frm.selseqno.focus();
			return false;
		}

		if (frm.txtthema.value == "") {
			alert("������� �Է��ϼ���");
			frm.txtthema.focus();
			return false;
		}
		frm.action = "job_reg_proc.asp";
		frm.method = "post";
		frm.submit();
	}

	function set_reset() {
		document.forms[0].reset();
	}

	function set_close() {
		this.close();
	}

	window.onload = function () {
	}
//-->
</script>
