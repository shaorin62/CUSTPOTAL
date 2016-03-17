<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
	dim midx, mtitle , objrs, sql
	Dim cookiemidx

%>

<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<script language="javascript" src="/js/ajax.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!--

//==================================공통페이지부분=============================================
		function getdata(page) {
//			 광고주 콤보 박스 가져오기
			page = (!!!page) ? "list.asp" : page ;
			var midx = 	 (document.getElementById("midx")) ? document.getElementById("midx").value : "" ;
			var custcode = 	 (document.getElementById("tcustcode")) ? document.getElementById("tcustcode").value : "" ;
			var FLAG = 	 (document.getElementById("tflag")) ? document.getElementById("tflag").value : "" ;
			var searchstring = 	 (document.getElementById("txtsearchstring")) ? document.getElementById("txtsearchstring").value : "" ;

			var params = "midx=" + midx + "&custcode=" + custcode + "&FLAG=" + FLAG + "&searchstring=" + searchstring ;

			document.getElementById("process").innerHTML = "<img src='http://ymail.net/skin/default/images/loding_clock.gif' style='margin-top:50px;'>" ;

			sendRequest(page, params, _getdata, "GET");
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


		function left_menu_getdata(page,midx,custcode,FLAG) {

			page = (!!!page) ? "list.asp" : page ;

			var params = "midx=" + midx + "&custcode=" + custcode + "&FLAG=" + FLAG ;
			document.getElementById("process").innerHTML = "<img src='http://ymail.net/skin/default/images/loding_clock.gif' style='margin-top:50px;'>" ;

			var midx = document.getElementById("midx");
			//var mstrmidx = document.getElementById("mstrmidx");
			//midx.innerText = midx ;
			var tcustcode = document.getElementById("tcustcode");
			if ( custcode == "midx" ) {
				custcode = ""
			}
			tcustcode.innerText = custcode ;
			var tflag = document.getElementById("tflag");
			tflag.innerText = FLAG ;

			sendRequest(page, params, _left_menu_getdata, "GET");
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




//==================================list.asp=============================================


function pop_report_reg() {
		var midx = document.forms[0].midx.value ;
		var tcustcode = document.forms[0].tcustcode.value ;
		var tflag = document.forms[0].tflag.value ;
		var url = "pop_report_reg.asp?midx="+midx+"&custcode="+tcustcode+"&flag="+tflag  ;
		var name = "" ;
		var opt = "width=658, height=680, resizable=no, scrollbars=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function pop_report_view(ridx) {
		var midx = document.forms[0].midx.value ;
		var url = "pop_report_view.asp?ridx="+ridx+"&midx=" +midx;
		var name = "" ;
		var opt = "width=658, height=680, resizable=no, scrollbars=yes, top=100, left=100";
		window.open(url, name, opt);
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

		getdata("list.asp");
		frm.txtsearchstring.value = "" ;
	}

//	window.onload = function () {
//
//		getdata();
//
//	}

//==================================여기까지list.asp=============================================



-->
</SCRIPT>

</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  >

<form target="scriptFrame" accept-charset="utf-8" >
<!--#include virtual="/mc/top.asp" -->
<input type="hidden" name="actionurl" value="list.asp">
<input type="hidden" name="midx" id="midx">
<input type="hidden" name="tcustcode" >
<input type="hidden" name="tflag" value= "midx" >
  <table id="Table_01" width="1240" height="600"  border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/mc/left_report_menu.asp"--></td>
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
					<TD  valign="top"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle" id="subtitle"></span></TD>
					<TD  align="right" valign="top">  <span class="navigator"  id="navi"></span></TD>
				</TR>
				</TABLE>
			</td>
          </tr>
          <tr>
            <td height="15" valign="top" colspan="2">&nbsp;</td>
          </tr>
          <tr>
            <td valign="top" colspan='2'>
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
            <td height="15" >&nbsp;</td>
          </tr>
          <tr>
            <td valign="top"><div id="process" style="text-align:center;"></div></td>
          </tr>
      </table></td>
    </tr>
	<tr>
	</tr>
  </table>
  <!--#include virtual="/bottom.asp" -->
</body>
</html>

