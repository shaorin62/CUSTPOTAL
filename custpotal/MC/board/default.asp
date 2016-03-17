<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
	dim midx, mtitle , objrs, sql
%>

<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<script type='text/javascript' src="/js/report.js"></script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  oncontextmenu="return false">

<form target="scriptFrame">
<!--#include virtual="/mc/top.asp" -->
<input type="hidden" name="midx" id="midx">
  <table id="Table_01" width="1240" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/mc/left_report_menu.asp" --></td>
      <td height="65" valign="top"><img src="/images/default_03.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td   valign="top"  height="100">
		<table width="1030" border="0" cellspacing="0" cellpadding="0" valign="top">
          <tr>
            <td height="19" valign="top">&nbsp;</td>
          </tr>
          <tr>
            <td height="17" valign="top">
				<TABLE  width="100%">
				<TR>
					<TD  valign="top"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle" id="subtitle"> 공지사항(공통) </span></TD>
					<TD  align="right" valign="top">  <span class="navigator"  id="navigate">매체별 리포트 &gt; 공지사항(공통) </span></TD>
				</TR>
				</TABLE>
			</td>
          </tr>
          <tr>
            <td height="15" valign="top">&nbsp;</td>
          </tr>
          <tr>
            <td valign="top" >
				<table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
                  <td width="80%" align="left" background="/images/bg_search.gif"> <span id="custcode2"></span> <input type="text" name="txtsearchstring" size="30"> <img src="/images/btn_search.gif" width="39" height="20"  class="styleLink" onClick="checkForSearch()" align="absmiddle"></td>
                  <td width="20%" align="right" background="/images/bg_search.gif"></td>
                  <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table>
			</td>
          </tr>
          <tr>
            <td valign="top"><iframe id="scriptFrame" name="scriptFrame" width="1030" height="570" frameborder="0" src=""></iframe></td>
          </tr>
      </table></td>
    </tr>
	<tr>
	<td colspan="2"><!--#include virtual="/bottom.asp" --></td>
	</tr>
  </table>
</body>
</html>
<script language="JavaScript">
<!--
	function pop_report_reg() {
		var midx = document.forms[0].midx.value ;
		var url = "pop_report_reg.asp?midx="+midx ;
		var name = "" ;
		var opt = "width=658, height=680, resizable=no, scrollbars=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function checkForSearch() {
		var frm = document.forms[0];
		var searchstring = frm.txtsearchstring.value ;
		if (searchstring.indexOf("--") != -1) {
			alert("사용할 수 없는 문자를 입력하셨습니다.");
			frm.txtsearchstring.value = "";
			frm.txtsearchstring.focus();
			return false;
		}
		frm.action = "list.asp";
		frm.method = "post";
		frm.submit();

frm.txtsearchstring.value = "" ;
	}
//-->
</script>