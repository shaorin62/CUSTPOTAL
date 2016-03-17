<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim pidx : pidx =  request("pidx")
	dim objrs, sql
	sql = "select m.contidx, m.title, m2.acceptdate, m2.status, m2.acceptweek, m2.nextacceptdate, m2.comment from dbo.wb_contact_mst m inner join  dbo.wb_contact_monitor_mst m2 on m.contidx = m2.contidx where m2.pidx = " & pidx
'	response.write sql
'	response.end
	call get_recordset(objrs, sql)

	dim contidx, title, acceptdate, status, acceptweek, nextacceptdate, comment
	if not objrs.eof then
		contidx = objrs("contidx")
		title = objrs("title")
		acceptdate = objrs("acceptdate")
		status = objrs("status")
		acceptweek = objrs("acceptweek")
		nextacceptdate = objrs("nextacceptdate")
		comment = objrs("comment")
	end if
	objrs.close
	sql = "select didx, filename from dbo.wb_contact_monitor_dtl where pidx = " & pidx
	call get_recordset(objrs, sql)

	dim didx, filename
	if not objrs.eof then
		set didx = objrs("didx")
		set filename = objrs("filename")
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
<form  enctype="multipart/form-data">
<input type="hidden" name="pidx" value="<%=pidx%>">
<input type="hidden" name="txtuserid" value="<%=request.cookies("userid")%>">
<table width="522" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> 옥외 모니터링 현황 <<%=title%>>  </td>
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
            <td class="hw">등록일자</td>
            <td colspan="3" class="bw"><input name="txtacceptdate" type="text" id="txtacceptdate"  value="<%=acceptdate%>"> <img src="/images/calendar.gif" width="39" height="20" border="0" align="absmiddle" onclick="Calendar_D(document.all.txtacceptdate)"  class="styleliink"></td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">검수주차</td>
            <td colspan="3" class="bw">
			<select name="selweek" >
				<option value="1" <%if acceptweek = "1" then response.write " selected "%>> 1주차
				<option value="2" <%if acceptweek = "2" then response.write " selected "%>> 2주차
				<option value="3" <%if acceptweek = "3" then response.write " selected "%>> 3주차
				<option value="4" <%if acceptweek = "4" then response.write " selected "%>> 4주차
				<option value="5" <%if acceptweek = "5" then response.write " selected "%>> 5주차
            </select></td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">검수상태</td>
            <td colspan="3" class="bw"><input type="radio" value="양호" name="rdostatus"  <%if status = "양호" then response.write " checked "%>> 양호 <input type="radio" value="불량" name="rdostatus" <%if status = "불량" then response.write " checked "%>> 불량 </td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">검수예정일</td>
            <td colspan="3" class="bw"><input name="txtnextacceptdate" type="text" id="txtnextacceptdate" value="<%=nextacceptdate%>"> <img src="/images/calendar.gif" width="39" height="20" border="0" align="absmiddle" onclick="Calendar_D(document.all.txtnextacceptdate)"  class="styleliink"></td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
		<% do until objrs.eof %>
		  <tr>
			<td colspan="4" >
				<table border="0" cellpadding="0" cellspacing="0" align="center">
				  <tr>
					<td style="padding-top:5px;padding-bottom:5px;" width="466"><input type="checkbox" name="didx" value="<%=didx%>" onclick="set_monitor_delete(<%=didx%>);"><img src="/pds/monitor/<%=filename%>" width="440" border="0" align="absmiddle"> </td>
				  </tr>
				  <tr>
					<td  bgcolor="#E7E7DE" height="1" style="padding-bottom:5px;"></td>
				  </tr>
				</table>
			</td>
		  </tr>
		<%
			objrs.movenext
			loop
			objrs.close
			set objrs = nothing
		%>
          <tr>
            <td class="hw">사진첨부</td>
            <td colspan="3" class="bw"><input type="file" name="txtfile" style="width:372;"></td>
          </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">특이사항</td>
            <td colspan="3" class="bw"><textarea name="txtcomment" rows="5"  style="width:372;padding-top:3px;"><%=comment%></textarea></td>
          </tr>
		  <tr>
				<td colspan="4"  height="50" valign="bottom" align="right"> <img src="/images/btn_save.gif" width="59" height="18" vspace="5" style="cursor:hand" onClick="check_submit();"  ><img src="/images/btn_close.gif" width="57" height="18" vspace="6" style="cursor:hand" onClick="set_close();" hspace="10" >
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
		if (confirm("선택한 모니터링 사진을 삭제하시겠습니까?")) {
			location.href = "pop_monitor_delete.asp?didx="+idx+"&pidx=<%=pidx%>";
		} else  {
			for (var i = 0 ; i < frm.didx.length; i++) {
				frm.didx[i].checked = false;
			}
		}
	}

	window.onload = function () {
		self.focus();
	}
//-->
</script>