<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")

	dim objrs, sql
	sql = "select title, photo_1, photo_2, photo_3, photo_4 from dbo.wb_contact_md m inner join dbo.wb_contact_md_dtl d on m.contidx = d.contidx and m.sidx = d.sidx where m.contidx ="&contidx&" and m.sidx="&sidx&" and d.cyear="&cyear&" and d.cmonth="&cmonth
	call get_recordset(objrs, sql)


%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<title>SK MARKETING & COMPANY </title>
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
<table width="522" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> 계약 매체 사진 관리  </td>
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
	  <table border="0" cellpadding="0" cellspacing="0" align="center">
          <tr>
            <td class="hw">매체명</td>
            <td colspan="3" class="bw"><%=objrs("title")%></td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">등록년월</td>
            <td colspan="3" class="bw"><%=cyear%>.<%=cmonth%></td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
		  <tr>
			<td colspan="4"  align="center" >
				<table border="0" cellpadding="0" cellspacing="0" align="center">
				  <tr>
					<td style="padding-top:5px;padding-bottom:5px;" width="476"><img src="<%if not isnull(objrs("photo_1")) then response.write "/pds/media/"&objrs("photo_1") else response.write "/images/noimage.gif"%>" width="476" border="0" align="top"> </td>
				  </tr>
				  <tr>
					<td  height="30" align="right" style="padding-bottom:5px;"> <%if isnull(objrs("photo_1")) then response.write "<span onclick='pop_photo_reg(1)' class='stylelink'>등록</span> | " else response.write "<span onclick='pop_photo_edit(1)' class='stylelinke'>수정</span> | "%><span class="stylelink" onclick="set_photo_delete(1)">삭제</span></td>
				  </tr>
				  <tr>
					<td  bgcolor="#E7E7DE" height="1" style="padding-bottom:5px;"></td>
				  </tr>
				</table>
			</td>
		  </tr>
		  <tr>
			<td colspan="4" align="center" >
				<table border="0" cellpadding="0" cellspacing="0" align="center">
				  <tr>
					<td style="padding-top:5px;padding-bottom:5px;" width="476"><img src="<%if not isnull(objrs("photo_2")) then response.write "/pds/media/"&objrs("photo_2") else response.write "/images/noimage.gif"%>" width="476" border="0" align="top"> </td>
				  </tr>
				  <tr>
					<td  height="30" align="right" style="padding-bottom:5px;"> <%if isnull(objrs("photo_2")) then response.write "<span onclick='pop_photo_reg(2)' class='stylelink'>등록</span> | " else response.write "<span onclick='pop_photo_edit(2)' class='stylelinke'>수정</span> | "%><span class="stylelink" onclick="set_photo_delete(2)">삭제</span></td>
				  </tr>
				  <tr>
					<td  bgcolor="#E7E7DE" height="1" style="padding-bottom:5px;"></td>
				  </tr>
				</table>
			</td>
		  </tr>
		  <tr>
			<td colspan="4"  align="center" >
				<table border="0" cellpadding="0" cellspacing="0" align="center">
				  <tr>
					<td style="padding-top:5px;padding-bottom:5px;" width="476"><img src="<%if not isnull(objrs("photo_3")) then response.write "/pds/media/"&objrs("photo_3") else response.write "/images/noimage.gif"%>" width="476" border="0" align="top"> </td>
				  </tr>
				  <tr>
					<td  height="30" align="right" style="padding-bottom:5px;"><%if isnull(objrs("photo_3")) then response.write "<span onclick='pop_photo_reg(3)' class='stylelink'>등록</span> | " else response.write "<span onclick='pop_photo_edit(3)' class='stylelinke'>수정</span> | "%><span class="stylelink" onclick="set_photo_delete(3)">삭제</span></td>
				  </tr>
				  <tr>
					<td  bgcolor="#E7E7DE" height="1" style="padding-bottom:5px;"></td>
				  </tr>
				</table>
			</td>
		  </tr>
		  <tr>
			<td colspan="4"  align="center" >
				<table border="0" cellpadding="0" cellspacing="0" align="center">
				  <tr>
					<td style="padding-top:5px;padding-bottom:5px;" width="476"><img src="<%if not isnull(objrs("photo_4")) then response.write "/pds/media/"&objrs("photo_4") else response.write "/images/noimage.gif"%>" width="476" border="0" align="top"> </td>
				  </tr>
				  <tr>
					<td  height="30" align="right" style="padding-bottom:5px;"> <%if isnull(objrs("photo_4")) then response.write "<span onclick='pop_photo_reg(4)' class='stylelink'>등록</span> | " else response.write "<span onclick='pop_photo_edit(4)' class='stylelinke'>수정</span> | "%><span class="stylelink" onclick="set_photo_delete(4)">삭제</span></td>
				  </tr>
				  <tr>
					<td  bgcolor="#E7E7DE" height="1" style="padding-bottom:5px;"></td>
				  </tr>
				</table>
			</td>
		  </tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
		  <tr>
				<td colspan="4"  height="50" valign="bottom" align="right"><img src="/images/btn_close.gif" width="57" height="18" vspace="6" style="cursor:hand" onClick="set_close();" hspace="10" >
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
<script language="JavaScript">
<!--
	function set_close() {
		this.close();
	}

	function check_submit() {
		var frm = document.forms[0];
		frm.action = "pop_monitor_edit_proc.asp";
		frm.method = "post";
		frm.submit();
	}

	function set_monitor_delete(idx) {
		var frm = document.forms[0];
		if (confirm("삭제하시겠습니까?")) {
			location.href = "pop_photo_delete.asp?num="+idx;
		}
	}

	function pop_photo_reg(idx) {
		var url = "pop_photo_reg.asp?idx="+idx;
		var name = "pop_photo_reg";
		var opt = "width=540, height=300, resizable=no, scrollbars=no, status=yes, left=100, top=660";
		window.open(url, name, opt);
	}

	function pop_photo_edit(idx) {
		var url = "pop_photo_edit.asp?num="+idx;
		var name = "pop_photo_edit";
		var opt = "width=540, height=300, resizable=no, scrollbars=no, status=yes, left=100, top=660";
		window.open(url, name, opt);
	}


	window.onload = function () {
		self.focus();
	}
//-->
</script>