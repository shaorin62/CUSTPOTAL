<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim objrs, sql
%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<script language="javascript" src="/js/ajax.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!--

//==================================공통페이지부분=============================================
		function getdata(page) {
//			 광고주 콤보 박스 가져오기
			page = (!!!page) ? "menu.asp" : page ;
//			var findID = 	 (document.getElementById("txtfindID")) ? document.getElementById("txtfindID").value : "" ;
			var params = "";
			sendRequest(page, params, _getdata, "GET");
			document.getElementById("process").innerHTML = "<img src='http://ymail.net/skin/default/images/loding_clock.gif' style='margin-top:50px;'>" ;
		}

		function _getdata() {
			var process = document.getElementById("process");
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
						process.innerHTML = xmlreq.responseText ;
				} else {
						process.innerHTML = "<img src='http://ymail.net/skin/default/images/loding_clock.gif' style='margin-top:50px;'>" ;
				}
			}
		}

		function left_getdata(page) {
//			 광고주 콤보 박스 가져오기
			page = (!!!page) ? "menu.asp" : page ;
//			var cyear = 	 (document.getElementById("cyear")) ? document.getElementById("cyear").value : "" ;
//			var cmonth = 	 (document.getElementById("cmonth")) ? document.getElementById("cmonth").value : "" ;
//			var cyear2 = 	 (document.getElementById("cyear2")) ? document.getElementById("cyear2").value : "" ;
//			var cmonth2 = (document.getElementById("cmonth2")) ? document.getElementById("cmonth2").value : "" ;
//			var custcode2 = (document.getElementById("custcode2")) ? document.getElementById("custcode2").value : "" ;
			var params = "";
			sendRequest(page, params, _left_getdata, "GET");
			document.getElementById("process").innerHTML = "<img src='http://ymail.net/skin/default/images/loding_clock.gif' style='margin-top:50px;'>" ;

			var tcustcode = document.getElementById("tcustcode");
			tcustcode.innerText = "" ;
			var tflag = document.getElementById("tflag");
			tflag.innerText = "" ;

		}

		function _left_getdata() {
			var process = document.getElementById("process");
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
						process.innerHTML = xmlreq.responseText ;
				} else {
						process.innerHTML = "<img src='http://ymail.net/skin/default/images/loding_clock.gif' style='margin-top:50px;'>" ;
				}
			}
		}


		function left_menu_getdata(page,custcode,FLAG) {
//			 광고주 콤보 박스 가져오기

			page = (!!!page) ? "menu.asp" : page ;
//			var cyear = 	 (document.getElementById("cyear")) ? document.getElementById("cyear").value : "" ;
//			var cmonth = 	 (document.getElementById("cmonth")) ? document.getElementById("cmonth").value : "" ;
//			var cyear2 = 	 (document.getElementById("cyear2")) ? document.getElementById("cyear2").value : "" ;
//			var cmonth2 = (document.getElementById("cmonth2")) ? document.getElementById("cmonth2").value : "" ;

			var params = "custcode=" + custcode + "&FLAG=" + FLAG ;
			sendRequest(page, params, _left_menu_getdata, "GET");

			document.getElementById("process").innerHTML = "<img src='http://ymail.net/skin/default/images/loding_clock.gif' style='margin-top:50px;'>" ;

			var tcustcode = document.getElementById("tcustcode");
			tcustcode.innerText = custcode ;

			var tflag = document.getElementById("tflag");
			tflag.innerText = FLAG ;


		}

		function _left_menu_getdata() {
			var process = document.getElementById("process");
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
						process.innerHTML = xmlreq.responseText ;
				} else {
						process.innerHTML = "<img src='http://ymail.net/skin/default/images/loding_clock.gif' style='margin-top:50px;'>" ;
				}
			}

		}


//==================================여기까지공통페이지부분=============================================



window.onload = getdata;
-->
</SCRIPT>

</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<form target="scriptFrame">
<!--#include virtual="/cust/top.asp" -->
<input type="hidden" name="actionurl" value="menu.asp">
<input type="text" name="tcustcode">
<input type="text" name="tflag">
  <table id="Table_01" width="1240" height="600" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/cust/left_admin_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_admin.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td height="600" align="left" valign="top"><table width="1030" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="19">&nbsp;</td>
          </tr>
          <tr>
            <td height="17">
			 <TABLE  width="100%">
            <TR>
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle" id="subtitle"> <%=request.cookies("custname")%> </span></TD>
				<TD width="50%" align="right"><span class="navigator" id="navi">관리모드 &gt; 메뉴관리 &gt; <%=request.cookies("custname")%> </span></TD>
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
                  <td width="50%" align="left" background="/images/bg_search.gif"><span id="searchsection">
				  <!--<input type="text" name="txtsearchstring"> <img src="/images/btn_search.gif" width="39" height="20" align="top" class="styleLink" onClick="checkForSearch(document.forms[0].txtsearchstring.value)">-->
				  </span></td>
                  <td width="50%" align="right" background="/images/bg_search.gif"><img src="/images/btn_menu_reg.gif" width="78" height="18" alt="" border="0" class="account" onclick="pop_reg();" id="btnReg" style="cursor:hand;"></td>
                  <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="15" >&nbsp;</td>
          </tr>
          <tr>
            <td ><div id="process" style="text-align:center;"></div></td>
          </tr>
      </table></td>
    </tr>
  </table>
<!--#include virtual="/bottom.asp" -->
</body>
</html>

