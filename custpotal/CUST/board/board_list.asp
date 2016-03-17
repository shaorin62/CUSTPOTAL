<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
	dim objrs, sql
	dim custcode : custcode = request("selcustcode")
	if custcode = "" then custcode = request.cookies("custcode")
	dim custcode2 : custcode2 = request("selcustcode2")
	if custcode2 = "" then custcode2 = custcode
	if custcode2 = "" then custcode2 = request.cookies("custcode2")

	dim midx : midx = request("midx")
	dim mtitle : mtitle = request("mtitle")
	if midx = "" then
		sql = "select midx, title from dbo.wb_menu_mst where custcode is null"
		call get_recordset(objrs, sql)
		midx = objrs(0).value
		mtitle = objrs(1).value
		objrs.close
	end if


	dim gotopage : gotopage = request.QueryString("gotopage")
	if gotopage = "" then gotopage = 1
	dim searchstring : searchstring = request("txtsearchstring")

	dim pagesize : pagesize = 10
	dim objrs2
%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  oncontextmenu="return false">

<form >
<!--#include virtual="/cust/top.asp" -->
  <table id="Table_01" width="1240" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/cust/left_report_menu.asp" --></td>
      <td height="65"><img src="/images/default_03.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td height="600" align="left" valign="top"><table width="1030" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="19">&nbsp;</td>
          </tr>
          <tr>
            <td height="17"><TABLE  width="100%">
            <TR>
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"> <%=mtitle%> </span></TD>
				<TD width="50%" align="right"><span class="navigator" >매체별 리포트 &gt; 리포트 목록 &gt; <%=mtitle%></span></TD>
            </TR>
            </TABLE>
			</td>
          </tr>
          <tr>
            <td height="15">&nbsp;</td>
          </tr>
          <tr>
            <td ><table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
                  <td width="80%" align="left" background="/images/bg_search.gif"> <%'call get_custcode_mst(custcode, null, "contact_list.asp")%><span id="custcode2"><%'call get_blank_select("사업부를 선택하세요", 207)%></span> <input type="text" name="txtsearchstring" size="30"> <img src="/images/btn_search.gif" width="39" height="20"  class="styleLink" onClick="checkForSearch()" align="absmiddle"></td>
                  <td width="20%" align="right" background="/images/bg_search.gif"><img src="/images/btn_report_reg.gif" width="88" height="18" border="0" onclick="pop_report_reg();" class="stylelink"></td>
                  <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="26" >&nbsp;</td>
          </tr>
          <tr>
            <td ><!--  -->c
<!--  --></td>
          </tr>
      </table></td>
    </tr>
  </table>
<!--#include virtual="bottom.asp" -->
</body>
<iframe id="ifrm" name="ifrm" width="0" height="0" frameborder="0" src="about:blank"></iframe>
</html>
<script language="JavaScript">
<!--
	function pop_report_reg() {
		var url = "pop_report_reg.asp?midx=<%=midx%>&custcode=<%=custcode%>&custcode2=<%=custcode2%>" ;
		var name = "pop_report_reg" ;
		var opt = "width=658, height=500, resizable=no, scrollbars=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function pop_report_view(ridx) {
		var url = "pop_report_view.asp?ridx="+ridx+"&midx=<%=midx%>";
		var name = "pop_report_view" ;
		var opt = "width=658, height=500, resizable=no, scrollbars=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function checkForView(idx) {
		location.href="view.asp?idx="+idx+"&gotopage=<%=gotopage%>";
	}


	function go_page(url) {
//		var frm = document.getElementById("ifrm") ;
//		var custcode = document.forms[0].selcustcode.options[document.forms[0].selcustcode.selectedIndex].value ;
//		var custcode2 = "<%=custcode2%>" ;
//		frm.src="/inc/frm_code.asp?custcode="+custcode +"&custcode2="+custcode2;
	}


	function checkForDownload(name) {
		location.href="download.asp?filename="+name;
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

	}

	window.onload = function () {
		go_page("");
	}
//-->
</script>