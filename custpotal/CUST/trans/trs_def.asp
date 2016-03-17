<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<script language="javascript" src="/js/ajax.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!--

		function getdata(page) {
//			 광고주 콤보 박스 가져오기
			page = (!!!page) ? "trs_01.asp" : page ;
			var cyear = 	 (document.getElementById("cyear")) ? document.getElementById("cyear").value : "" ;
			var cmonth = 	 (document.getElementById("cmonth")) ? document.getElementById("cmonth").value : "" ;
			var cyear2 = 	 (document.getElementById("cyear2")) ? document.getElementById("cyear2").value : "" ;
			var cmonth2 = (document.getElementById("cmonth2")) ? document.getElementById("cmonth2").value : "" ;
			var custcode2 = (document.getElementById("custcode2")) ? document.getElementById("custcode2").value : "" ;
			var params = "cyear="+cyear+"&cmonth="+cmonth+"&cyear2="+cyear2+"&cmonth2="+cmonth2+"&custcode2=" + custcode2 +"&initpage=1";
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
			page = (!!!page) ? "trs_01.asp" : page ;
			var cyear = 	 (document.getElementById("cyear")) ? document.getElementById("cyear").value : "" ;
			var cmonth = 	 (document.getElementById("cmonth")) ? document.getElementById("cmonth").value : "" ;
			var cyear2 = 	 (document.getElementById("cyear2")) ? document.getElementById("cyear2").value : "" ;
			var cmonth2 = (document.getElementById("cmonth2")) ? document.getElementById("cmonth2").value : "" ;
			var custcode2 = (document.getElementById("custcode2")) ? document.getElementById("custcode2").value : "" ;
			var params = "cyear="+cyear+"&cmonth="+cmonth+"&cyear2="+cyear2+"&cmonth2="+cmonth2+"&custcode2="+custcode2 +"&initpage=0";
			sendRequest(page, params, _getdata, "GET");
			document.getElementById("process").innerHTML = "<img src='http://ymail.net/skin/default/images/loding_clock.gif' style='margin-top:50px;'>" ;

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

		function get_excel_sheet(page) {
			// 엑셀전환
			page = (!!!page) ? "not" : page ;
			if (page == "not"){
				alert("등록되지않은 레포트 입니다.")
				return;
			}
			var cyear = 	 (document.getElementById("cyear")) ? document.getElementById("cyear").value : "" ;
			var cmonth = 	 (document.getElementById("cmonth")) ? document.getElementById("cmonth").value : "" ;
			var cyear2 = 	 (document.getElementById("cyear2")) ? document.getElementById("cyear2").value : "" ;
			var cmonth2 = (document.getElementById("cmonth2")) ? document.getElementById("cmonth2").value : "" ;
			var custcode2 = (document.getElementById("custcode2")) ? document.getElementById("custcode2").value : "" ;

			location.href = "xls_"+page+"?cyear="+cyear+"&cmonth="+cmonth+"&cyear2="+cyear2+"&cmonth2="+cmonth2+"&custcode2=" + custcode2 +"&initpage=1";
		}

	window.onload = getdata;


//-->
</SCRIPT>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  >
<form target="scriptFrame">
<!--#include virtual="/cust/top.asp" -->

  <table width="1240" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="Table_01">
    <tr>
      <td valign="top"><!--#include virtual="/cust/trans_menu.asp"--></td>
      <td align="left" valign="top"><img src="/images/middle_navigater_trans.gif" width="1030" height="65" alt="">
	  <div id="process" style="text-align:center;"></div>
	  </td>
    </tr>
	<tr>
	<td colspan="2"><!--#include virtual="/bottom.asp" --></td>
	</tr>
  </table>
</form>
</body>
</html>
