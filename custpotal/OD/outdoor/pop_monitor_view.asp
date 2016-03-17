<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim pidx : pidx =  request("pidx")
	dim objrs, sql
	sql = "select m.title, m2.acceptdate, m2.status, m2.acceptweek, m2.nextacceptdate, m2.comment from dbo.wb_contact_mst m inner join  dbo.wb_contact_monitor_mst m2 on m.contidx = m2.contidx where m2.pidx = " & pidx
'	response.write sql
'	response.end
	call get_recordset(objrs, sql)

	dim title, acceptdate, status, acceptweek, nextacceptdate, comment
	if not objrs.eof then
		title = objrs("title")
		acceptdate = objrs("acceptdate")
		status = objrs("status")
		acceptweek = objrs("acceptweek")
		nextacceptdate = objrs("nextacceptdate")
		comment = objrs("comment")
	end if
	objrs.close
	sql = "select filename from dbo.wb_contact_monitor_dtl where pidx = " & pidx
	call get_recordset(objrs, sql)

	dim filename
	if not objrs.eof then
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
<form>
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
	<table border="0" cellpadding="0" cellspacing="0">
		<tr height="31">
			<td class="thd" width="110">검수일자</td>
			<td class="tbd w140" ><%=acceptdate%></td>
			<td class="thd" width="110">검수주차</td>
			<td class="tbd w140" ><%=acceptweek%> 주차</td>
		</tr>
		<tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		</tr>
		<tr height="31">
			<td class="thd" width="110">등록상태</td>
			<td class="tbd w140" ><%=status%></td>
			<td class="thd" width="110">다음 검수예정일</td>
			<td class="tbd w140" ><%=nextacceptdate%></td>
		</tr>
		<tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		</tr>
		<% do until objrs.eof %>
		<tr>
			<td class="tdbd" colspan="4" style="padding-left:0px; padding-top:5px; padding-bottom:5px;" align="center"><img src="/pds/monitor/<%=filename%>" width="460" border="0"></td>
		</tr>
		<tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		</tr>
		<%
			objrs.movenext
			loop
			objrs.close
			set objrs = nothing
		%>
		<tr height="31">
			<td class="thd" width="100">특이사항</td>
			<td class="tdbd s7" colspan="3"><%if not isnull(comment) then response.write replace(comment, chr(13)&chr(10), "<br>")%>&nbsp;</td>
		</tr>
		<tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		</tr>
        <tr>
          <td width="50%" height="50" align="left" valign="bottom" colspan="2"></td>
              <td width="50%" align="right" valign="bottom" colspan="2"><img src="/images/btn_edit.gif" width="57" height="18" hspace="10" vspace="5" style="cursor:hand" onClick="pop_monitor_edit();"  ><img src="/images/btn_close.gif" width="57" height="18" vspace="5" style="cursor:hand" onClick="set_close();"  ></td>
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

	function pop_monitor_edit() {
		location.href = "pop_monitor_edit.asp?pidx=<%=pidx%>";
	}
	window.onload = function () {
		self.focus();
	}
//-->
</script>