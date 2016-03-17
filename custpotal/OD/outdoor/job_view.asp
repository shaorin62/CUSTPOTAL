<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim gotopage : gotopage = request("gotopage")
	dim searchstring : searchstring = request("searchstring")
	dim jobidx : jobidx =  request("jobidx")
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
      <td height="600" align="left" valign="top"><table width="1002" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="24">&nbsp;</td>
          </tr>
          <tr>
            <td height="19" valign="top" class="navigator" > 옥외관리 &gt; 소재관리 &gt; 소재정보</td>
          </tr>
          <tr>
            <td height="17"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"> 소재정보 </span></td>
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
				<td class="hw">광고주</td>
				<td class="bw bd"><%call get_custcode_mst(custcode, "R", "job_reg.asp")%></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr height="31">
				<td class="hw">브랜드</td>
				<td class="bw bd"> <% call get_jobcust(custcode, seqno, "R", null)%>
				</td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr height="31">
				<td class="hw">소재명</td>
				<td class="bw bd"><%=thema%></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			</table>
			  <table width="1002" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="50%" height="50" align="left" valign="bottom"><a href="/od/outdoor/job_list.asp"><img src="/images/btn_list.gif" width="59" height="18" border="0"></a></td>
                  <td width="50%" align="right" valign="bottom"><img src="/images/btn_edit.gif" width="59" height="18" hspace="10" vspace="5" style="cursor:hand" onClick="go_jobcust_edit();"><img src="/images/btn_delete.gif" width="59" height="18" vspace="5" style="cursor:hand" onClick="go_jobcust_delete();"></td>
                </tr>
              </table></td>
          </tr>
      </table>
	  </td>
    </tr>
  </table>
<!--#include virtual="bottom.asp" -->
</form>
</body>
</html>
<script language="javascript">
<!--
	function go_jobcust_edit() {
		if (confirm("소재정보를 수정하시겠습니까?"))
		var url = "pop_job_edit.asp?jobidx=<%=jobidx%>";
		var name = "pop_job_edit";
		var opt = "width=540, height=233, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function go_jobcust_delete() {
		if (confirm("소재정보가 시스템에서 영구삭제됩니다.\n\n소재정보를 삭제하시겠습니까?"))
			location.href = "job_delete_proc.asp?jobidx=<%=jobidx%>&gotopage=<%=gotopage%>&searchstring=<%=searchstring%>";
	}

	window.onload = function () {
	}
//-->
</script>