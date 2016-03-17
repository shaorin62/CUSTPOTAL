<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim jobidx : jobidx =  request("jobidx")
	dim objrs, sql
	sql = "select s.seqno, s.thema, j.custcode, j.clientsubcode from dbo.wb_jobcust s inner join dbo.sc_jobcust j on s.seqno = j.seqno where jobidx = " & jobidx
'	response.write sql
'	response.end
	call get_recordset(objrs, sql)

	dim seqno : seqno = objrs("seqno")
	dim custcode : custcode = objrs("custcode")
	dim clientsubcode : clientsubcode = objrs("clientsubcode")
	dim thema : thema = objrs("thema")

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

<body  oncontextmenu="return false">
<form>
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> 소재 수정 </td>
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
				<td class="hw">광고주</td>
				<td class="bw"><%call get_custcode_mst(custcode, "R", "pop_job_edit.asp")%></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr height="31">
				<td class="hw">사업부서</td>
				<td class="bw"> <% call get_custcode_custcode2(custcode, clientsubcode, "R")%>
				</td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr height="31">
				<td class="hw">브랜드</td>
				<td class="bw"> <span id="jobcust"><% call get_jobcust(clientsubcode, seqno, "R", null)%></span>
				</td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr height="31">
				<td class="hw">소재명</td>
				<td class="bw"><input name="txtthema" type="text" id="txtthema"  style="width:330px" maxlength="100" value="<%=thema%>" class="kor"></td>
			</tr>
            <tr>
              <td  height="50" align="left" valign="bottom"><img src="/images/space.gif" width="59" height="20" border="0"></td>
              <td  align="right" valign="bottom"><img src="/images/btn_save.gif" width="59" height="18" hspace="10" vspace="5" style="cursor:hand" onClick="check_submit();"><img src="/images/btn_init.gif" width="67" height="18" vspace="5" style="cursor:hand" onClick="set_reset();"><img src="/images/btn_close.gif" width="57" height="18" vspace="5" style="cursor:hand" onClick="set_close();" hspace="10" ></td>
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
<input type="hidden" name="jobidx" value="<%=jobidx%>">
</form>
<iframe id="ifrm" name="ifrm" width="0" height="0" frameborder="0" src="about:blank"></iframe>
</body>
</html>
<script language="javascript">
<!--
	function check_submit() {
		var frm = document.forms[0];
		if (frm.txtthema.value == "" ) {
			alert("변경할 소재명을 입력하세요");
			frm.txtthema.focus();
			return false ;
		}

		frm.action = "job_edit_proc.asp";
		frm.method = "post";
		frm.submit() ;

	}

	function set_reset() {
		document.forms[0].reset();
	}

	function set_close() {
		this.close();
	}

	window.onload = function () {
		self.focus();
	}
//-->
</script>