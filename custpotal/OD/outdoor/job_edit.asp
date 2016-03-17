<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim jobidx : jobidx =  request("jobidx")
	dim gotopage : gotopage = request("gotopage")
	dim searchstring : searchstring = request("searchstring")
	dim objrs, sql
	sql = "select s.seqno, s.thema, j.custcode from dbo.wb_jobcust s inner join dbo.sc_jobcust j on s.seqno = j.seqno where jobidx = " & jobidx
'	response.write sql
'	response.end
	call get_recordset(objrs, sql)

	dim seqno : seqno = objrs("seqno")
	dim custcode : custcode = objrs("custcode")
	dim thema : thema = objrs("thema")

	objrs.close
	set objrs = nothing
%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  oncontextmenu="return false">
<form>
<!--#include virtual="/od/top.asp" -->
  <table id="Table_01" width="1240" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/od/left_outdoor_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_outdoor.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td height="400" align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="24">&nbsp;</td>
          </tr>
          <tr>
            <td height="19" valign="top" class="navigator" > 옥외관리 &gt; 소재관리 &gt; 소재정보변경</td>
          </tr>
          <tr>
            <td height="17"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle">소재정보변경</span></td>
          </tr>
          <tr>
            <td height="27">&nbsp;</td>
          </tr>
          <tr>
            <td valign="top" class="bdpdd">
			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td colspan="2" bgcolor="#cacaca" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd">광고주</td>
				<td class="tdbd"><%call get_custcode_mst(custcode, null, "job_reg.asp")%></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr height="31">
				<td class="tdhd">브랜드</td>
				<td class="tdbd"> <% call get_jobcust(custcode, seqno, null, null)%>
				</td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr height="31">
				<td class="tdhd">소재명</td>
				<td class="tdbd"><input name="txtthema" type="text" id="txtthema" size="50" maxlength="100" value="<%=thema%>"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			</table>
			  <table width="756" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="50%" height="50" align="left" valign="bottom"><a href="/od/outdoor/job_list.asp"><img src="/images/btn_list.gif" width="59" height="20" border="0"></a></td>
                  <td width="50%" align="right" valign="bottom"><img src="/images/btn_save.gif" width="59" height="20" hspace="10" vspace="5" style="cursor:hand" onClick="check_submit();"><img src="/images/btn_init.gif" width="67" height="20" vspace="5" style="cursor:hand" onClick="set_reset();"></td>
                </tr>
              </table></td>
          </tr>
      </table>
	  </td>
    </tr>
  </table>
<!--#include virtual="bottom.asp" -->
<input type="hidden" name="jobidx" value="<%=jobidx%>">
<input type="hidden" name="gotopage" value="<%=gotopage%>">
<input type="hidden" name="searchstring" value="<%=searchstring%>">
</form>
</body>
</html>
<script type="text/javascript" src="/js/calendar.js"></script>
<script language="javascript">
<!--
	function go_page(url) {
		var frm = document.forms[0];
		frm.action = "job_reg.asp";
		frm.method = "post";
		frm.submit();
	}

	function check_submit() {
		var frm = document.forms[0];
		if (frm.selcustcode.value == "") {
			alert("광고주를 선택하세요");
			frm.selcustcode.focus();
			return false;
		}

		if (frm.selseqno.value == "") {
			alert("브랜드를 선택하세요");
			frm.selseqno.focus();
			return false;
		}

		if (frm.txtthema.value == "") {
			alert("소재명을 입력하세요");
			frm.txtthema.focus();
			return false;
		}
		frm.action = "job_edit_proc.asp";
		frm.method = "post";
		frm.submit();
	}

	function set_reset() {
		document.forms[0].reset();
	}

	window.onload = function () {
	}
//-->
</script>