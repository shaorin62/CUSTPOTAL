<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim idx : idx = request("idx")
	dim photoIdx : photoIdx = request("photoIdx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
'
'	dim item
'	for each item in request.form
'		response.write item & " : " & request.form(item) & "<br>"
'	next

	dim objrs, sql
	sql = "select comment, filename, note from dbo.wb_contact_photo_mst m inner join  dbo.wb_contact_photo_dtl d on m.idx = d.mstidx where d.idx = " & photoIdx
	call get_recordset(objrs, sql)

	dim filename, note, comment

	if not objrs.eof then
		comment = objrs("comment")
		filename = objrs("filename")
		note = objrs("note")
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

<body background="/images/pop_bg.gif"  oncontextmenu="return false">
<form>
<input type="hidden" name="idx" value="<%=idx%>">
<input type="hidden" name="photoIdx" value="<%=photoIdx%>">
<input type="hidden" name="cyear" value="<%=cyear%>">
<input type="hidden" name="cmonth" value="<%=cmonth%>">
<table width="640" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top">&nbsp; <%=comment%></td>
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
<img src="/pds/media/<%=filename%>" width="600" border="1" onclick="set_close();" class="stylelink" alt="�̹����� Ŭ���Ͻø� â�� �����ϴ�.">
<!--  -->
	</td>
    <td background="/images/pop_right_middle_bg.gif">&nbsp;</td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
	<td height="31" bgcolor="#FFFFFF"><TEXTAREA NAME="txtnote" ROWS="2" style="width:600px;"><%=note%></TEXTAREA></td>
    <td background="/images/pop_right_middle_bg.gif">&nbsp;</td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
	<td height="31"  bgcolor="#FFFFFF" align="right"><IMG SRC="../images/btn_save.gif" WIDTH="57" HEIGHT="18" BORDER="0" ALT="" style="cursor:hand" onclick="edit_photo()"><IMG SRC="../images/btn_close.gif" WIDTH="57" HEIGHT="18" BORDER="0" ALT="" hspace="5" onclick="set_close()"></td>
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
	function edit_photo() {
		var frm = document.forms[0];
			frm.action = "edit_medium_photo_edit_proc.asp";
			frm.method = "post";
			frm.submit();

	}

	window.onload = function () {
		self.focus();
	}
//-->
</script>
