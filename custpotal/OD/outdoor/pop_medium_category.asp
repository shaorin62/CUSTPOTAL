<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim objrs, sql
	sql = "select ggroupname, mgroupname, sgroupname, dgroupname, mdidx, mdname from  dbo.vw_medium_category order by ggroupidx,mgroupname, sgroupname, dgroupname "
	call get_recordset(objrs, sql)

	dim ggroupname, mgroupname, sgroupname, dgroupname, mdidx, mdname
	if not objrs.eof then
		set ggroupname = objrs("ggroupname")
		set mgroupname = objrs("mgroupname")
		set sgroupname = objrs("sgroupname")
		set dgroupname = objrs("dgroupname")
		set mdidx = objrs("mdidx")
		set mdname = objrs("mdname")
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
<table width="522" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> 매체별 분류표 </td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="522" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="22"><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td>
<!--  -->
<TABLE border="0" cellpadding="0" cellspacing="1" bgcolor="#ECECEC">
	<TR>
		<TD class="thd" width="77">대분류</TD>
		<TD class="thd" width="100">중분류</TD>
		<TD class="thd" width="150">소분류</TD>
		<TD class="thd" width="150">세분류</TD>
	</TR>
	<% Do Until objrs.eof %>
	<TR class="stylelink" bgcolor="#FFFFFF"onclick="put_medium_cateogorycode('<%=mdidx%>','<%=mdname%>')">
		<TD height="29">&nbsp;<%=ggroupname%></TD>
		<TD >&nbsp;<%=mgroupname%></TD>
		<TD >&nbsp;<%=sgroupname%></TD>
		<TD >&nbsp;<%=dgroupname%></TD>
	</TR>
	<%
		objrs.movenext
		Loop

		objrs.close
		Set objrs = Nothing
	%>
	</TABLE>
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
	function put_medium_cateogorycode(dcode, dname) {
		var frm = window.opener.document.forms[0];
		frm.txtcategoryidx.value = dcode;
		frm.txtcategoryname.value = dname ;
		this.close();
	}
	window.onload = function () {
		self.focus();
	}
</script>