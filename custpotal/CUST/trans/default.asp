<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
	dim cyear, cyear2, cmonth, cmonth2
	dim yearmon, yearmon2


	if cyear = "" then cyear = year(date)
	if cmonth = "" then cmonth = month(date)
	if cyear2 = "" then cyear2 = year(date)
	if cmonth2 = "" then cmonth2 = month(date)
%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<SCRIPT LANGUAGE="JavaScript" src="/js/trans.js"></SCRIPT>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  oncontextmenu="return false">
<form target="scriptFrame">
<input type="hidden" name="actionurl" id="actionurl"  value="public_01.asp">
<input type="hidden" name="tcustcode" id="tcustcode">
<!--#include virtual="/cust/top.asp" -->
  <table width="1240" height="700" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="Table_01">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/cust/left_trans_menu.asp"--></td>
      <td height="65" valign="top"><img src="/images/middle_navigater_trans.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td height="600" align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="19">&nbsp;</td>
          </tr>
          <tr>
            <td height="17"><TABLE  width="100%">
            <TR>
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle" id="subtitle">월별 매체별 광고비 </span><span id="subname" class="subtitle"></span></TD>
				<TD width="50%" align="right"><span class="navigator"> 광고비 집행 &gt; <span  id="navigator">월별 매체별 광고비 </TD>
            </TR>
            </TABLE>
			</td>
          </tr>
          <tr>
            <td height="15">&nbsp;</td>
          </tr>
          <tr>
            <td >			<table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
                  <td align="left" background="/images/bg_search.gif">
				            <img src="/images/icon_dot_search.gif" width="4" height="3" align="absmiddle"> 검색년월
				            <span id="cyear"><%call get_year(Int(cyear))%></span><span id="cmonth"><%call get_month(Int(cmonth))%></span><span id="e7">~</span><span id="cyear2"><%call get_year2(Int(cyear2))%></span><span id="cmonth2"><%call get_month2(Int(cmonth2))%></span> <span id="custcode2"></span><img src="/images/btn_search.gif" width="39" height="20" align="top" class="styleLink" onclick="go_search();" hspace="5"> </td>
                  <td align="right" background="/images/bg_search.gif"><img src="/images/btn_excel.gif" width="85" height="22" align="absmiddle" vspace="5" class="stylelink" onclick="get_excel_sheet();"></td>
                  <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="26" align="right"> &nbsp; </td>
          </tr>
          <tr>
            <td >
				 <iframe src="public_01.asp?cyear=2000&cmonth=01&cyear2=2000&cmonth2=01" width="1032" height="544" frameborder="2" border="2" name="scriptFrame" id="scriptFrame"></iframe>
			</td>
          </tr>
      </table></td>
    </tr>
	<tr>
	<td colspan="2"><!--#include virtual="/bottom.asp" --></td>
	</tr>
  </table>

</form>
</body>
</html>
<SCRIPT LANGUAGE="JavaScript">
<!--
	function get_excel_sheet() {
		var frm = document.forms[0];
		var url = frm.actionurl.value ;

		var cmonth = frm.cmonth.options[frm.cmonth.selectedIndex].value;
		var cmonth2 = frm.cmonth2.options[frm.cmonth2.selectedIndex].value;
		if (cmonth< 10) {cmonth = "0" + cmonth.toString();}
		if (cmonth2< 10) {cmonth2 = "0" + cmonth2.toString();}
		var yearmon = frm.cyear.options[frm.cyear.selectedIndex].value + cmonth ;
		var yearmon2 = frm.cyear2.options[frm.cyear2.selectedIndex].value + cmonth2 ;
		var custcode2 = frm.tcustcode2.options[frm.tcustcode2.selectedIndex].value ;
		var custcode = frm.tcustcode.value ;
		location.href= "xls_"+url+"?yearmon="+yearmon+"&yearmon2="+yearmon2+"&custcode="+custcode+"&custcode2="+custcode2;
	}

	function go_search() {
		var frm = document.forms[0];
		var url = frm.actionurl.value ;
		frm.action = url ;
		frm.method = "post";
		frm.submit();

		var custname2 = frm.tcustcode2.options[frm.tcustcode2.selectedIndex].text ;
		var subname = document.getElementById("subname") ;
		if (frm.tcustcode2.selectedIndex == 0) {
			subname.innerText = "" ;
		} else {
			subname.innerText = "" ;
			subname.innerText = "(" + custname2 +")" ;
		}
	}

	function get_excel_sheet() {
		var frm = document.forms[0];
		var url = frm.actionurl.value ;
		var cmonth = frm.cmonth.options[frm.cmonth.selectedIndex].value;
		var cmonth2 = frm.cmonth2.options[frm.cmonth2.selectedIndex].value;
		if (cmonth< 10) {cmonth = "0" + cmonth.toString();}
		if (cmonth2< 10) {cmonth2 = "0" + cmonth2.toString();}
		var yearmon = frm.cyear.options[frm.cyear.selectedIndex].value + cmonth ;
		var yearmon2 = frm.cyear2.options[frm.cyear2.selectedIndex].value + cmonth2 ;
		var custcode2 = frm.tcustcode2.options[frm.tcustcode2.selectedIndex].value ;
		var custcode = frm.tcustcode.value ;
		location.href= "xls_"+url+"?yearmon="+yearmon+"&yearmon2="+yearmon2+"&custcode="+custcode +"&custcode2="+custcode2;
	}



//-->
</SCRIPT>