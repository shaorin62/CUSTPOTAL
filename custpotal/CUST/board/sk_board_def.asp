<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<%
	dim midx, mtitle , objrs, sql
	Dim pcustcode : pcustcode = ""
	Dim cyear : cyear =  Year(date)
	Dim cmonth : cmonth = Month(date)
	If Len(cmonth) = 1 Then cmonth = "0"&cmonth
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
			var custcode2 = 	 (document.getElementById("cmbcustcode")) ? document.getElementById("cmbcustcode").value : "" ;
			var FLAG = 	 (document.getElementById("tflag")) ? document.getElementById("tflag").value : "" ;
			var searchstring = 	 (document.getElementById("txtsearchstring")) ? document.getElementById("txtsearchstring").value : "" ;
			var highcategory = 	 (document.getElementById("cmbhighcategory")) ? document.getElementById("cmbhighcategory").value : "" ;
			var category		= 	 (document.getElementById("cmbcategory")) ? document.getElementById("cmbcategory").value : "" ;

//			var params = "midx=" + midx + "&custcode=" + custcode + "&FLAG=" + FLAG + "&searchstring=" + escape(searchstring);

			var params = "midx=" + midx + "&custcode=" + custcode2 + "&FLAG=" + FLAG + "&searchstring=" + escape(searchstring)  + "&highcategory=" + highcategory  + "&category=" + category ;

			document.getElementById("process").innerHTML = "<img src='http://ymail.net/skin/default/images/loding_clock.gif' style='margin-top:50px;'>" ;

			var tsearchstring = document.getElementById("txtsearchstring");
			tsearchstring.value = searchstring ;
			
			if (highcategory != "")
			{
				var thighcategory = document.getElementById("cmbhighcategory");
				thighcategory.value = highcategory ;
			}

			if (category != "")
			{
				var tcategory = document.getElementById("cmbcategory");
				tcategory.value = category ;
			}
			if (custcode2 != "")
			{
				var tcustcode = document.getElementById("cmbcustcode");
				tcustcode.value = custcode ;
			}

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

		function  get_pageSRC(page,midx,custcode,FLAG,title, gotopage, timgubun) {
			var srcHTML
			srcHTML = " "
			document.getElementById("midx").value = midx ;
			document.getElementById("ttitle").value = title ;
			document.getElementById("srcTD").innerHTML = srcHTML

			if (midx == "16")
			{
				srcHTML = "  <table width='1030' height='35' border='0' cellpadding='0' cellspacing='0'> "
				srcHTML = srcHTML + "  <tr> "
				srcHTML = srcHTML + "   <td width='13'><img src='/images/bg_search_left.gif' width='13' height='35'></td>"
				srcHTML = srcHTML + "   <td width='80%' align='left' background='/images/bg_search.gif'>"
				srcHTML = srcHTML + "   <span id='highcategory'>대분류검색</span> <span id='category'>중분류 검색</span>"
				srcHTML = srcHTML + "   <input type='text' name='txtsearchstring' size='20'>"
				srcHTML = srcHTML + "   <img src='/images/btn_search.gif' width='39' height='20'  class='styleLink' onClick='checkForSearch()' align='absmiddle'>"
				srcHTML = srcHTML + "   </td>"
				srcHTML = srcHTML + "   <td width='20%' align='right' background='/images/bg_search.gif'>"
				srcHTML = srcHTML + "   <img src='/images/btn_report_reg.gif' width='88' height='18' border='0' onclick='pop_report_regooh();' class='stylelink'>"
				srcHTML = srcHTML + "  </td> "
				srcHTML = srcHTML + "   <td width='13'><img src='/images/bg_search_right.gif' width='13' height='35'></td>"
				srcHTML = srcHTML + "   </tr>"
				srcHTML = srcHTML + "   </table>"
				
				document.getElementById("srcTD").innerHTML = srcHTML ;
				document.getElementById("tgotopage").value = gotopage;
				document.getElementById("attr02").value = "";
				gethighcategorycombo();

			}else if (timgubun == "inter"){
				srcHTML = "  <table width='1030' height='35' border='0' cellpadding='0' cellspacing='0'> "
				srcHTML = srcHTML + "  <tr> "
				srcHTML = srcHTML + "   <td width='13'><img src='/images/bg_search_left.gif' width='13' height='35'></td>"
				srcHTML = srcHTML + "   <td width='80%' align='left' background='/images/bg_search.gif'>"
				srcHTML = srcHTML + "    검색년월  <%call getyear3(cyear)%><%call getmonth3(cmonth)%>"
				srcHTML = srcHTML + "   <span id='custcode'></span>"
				srcHTML = srcHTML + "   <input type='text' name='txtsearchstring' size='30'>"
				srcHTML = srcHTML + "   <img src='/images/btn_search.gif' width='39' height='20'  class='styleLink' onClick='checkForSearch()' align='absmiddle'>"
				srcHTML = srcHTML + "   </td>"
				srcHTML = srcHTML + "   <td width='20%' align='right' background='/images/bg_search.gif'>"
				srcHTML = srcHTML + "   <img src='/images/btn_report_reg.gif' width='88' height='18' border='0' onclick='pop_report_regooh();' class='stylelink'>"
				srcHTML = srcHTML + "  </td> "
				srcHTML = srcHTML + "   <td width='13'><img src='/images/bg_search_right.gif' width='13' height='35'></td>"
				srcHTML = srcHTML + "   </tr>"
				srcHTML = srcHTML + "   </table>"

				document.getElementById("srcTD").innerHTML = srcHTML ;		
				document.getElementById("tgotopage").value = gotopage;
				document.getElementById("attr02").value = timgubun;
				getcustcombo_report()

//				document.getElementById("txtsearchstring").value = searchstring ;
				
			} else {
				srcHTML = "  <table width='1030' height='35' border='0' cellpadding='0' cellspacing='0'> "
				srcHTML = srcHTML + "  <tr> "
				srcHTML = srcHTML + "   <td width='13'><img src='/images/bg_search_left.gif' width='13' height='35'></td>"
				srcHTML = srcHTML + "   <td width='80%' align='left' background='/images/bg_search.gif'>"
				srcHTML = srcHTML + "   <span id='custcode2'></span>"
				srcHTML = srcHTML + "   <input type='text' name='txtsearchstring' size='30'>"
				srcHTML = srcHTML + "   <img src='/images/btn_search.gif' width='39' height='20'  class='styleLink' onClick='checkForSearch()' align='absmiddle'>"
				srcHTML = srcHTML + "   </td>"
				srcHTML = srcHTML + "   <td width='20%' align='right' background='/images/bg_search.gif'>"
				srcHTML = srcHTML + "   <img src='/images/btn_report_reg.gif' width='88' height='18' border='0' onclick='pop_report_regooh();' class='stylelink'>"
				srcHTML = srcHTML + "  </td> "
				srcHTML = srcHTML + "   <td width='13'><img src='/images/bg_search_right.gif' width='13' height='35'></td>"
				srcHTML = srcHTML + "   </tr>"
				srcHTML = srcHTML + "   </table>"

				document.getElementById("srcTD").innerHTML = srcHTML ;
				
//				document.getElementById("txtsearchstring").value = searchstring ;
				document.getElementById("tgotopage").value ="";
				document.getElementById("attr02").value = "";
				left_menu_getdata(gotopage);
			}


		}

		function left_menu_getdata(gotopage) {
			var page			= "list.asp" ;
			var midx			= 	 (document.getElementById("midx")) ? document.getElementById("midx").value : "" ;
			var custcode	= 	 (document.getElementById("tcustcode")) ? document.getElementById("tcustcode").value : "" ;
			var title			= 	 (document.getElementById("ttitle")) ? document.getElementById("ttitle").value : "" ;
			var FLAG		= 	 (document.getElementById("tflag")) ? document.getElementById("tflag").value : "" ;
			var searchstring	= 	 (document.getElementById("txtsearchstring")) ? document.getElementById("txtsearchstring").value : "" ;
			var highcategory = 	 (document.getElementById("cmbhighcategory")) ? document.getElementById("cmbhighcategory").value : "" ;
			var category		= 	 (document.getElementById("cmbcategory")) ? document.getElementById("cmbcategory").value : "" ;
			var ccustcode		= 	 (document.getElementById("cmbcustcode")) ? document.getElementById("cmbcustcode").value : "" ;
			var cyear		= 	 (document.getElementById("cyear")) ? document.getElementById("cyear").value : "" ;
			var cmonth		= 	 (document.getElementById("cmonth")) ? document.getElementById("cmonth").value : "" ;
			var attr02		= 	 (document.getElementById("attr02")) ? document.getElementById("attr02").value : "" ;

			var params = "midx=" + midx + "&custcode=" + custcode + "&FLAG=" + FLAG + "&searchstring=" + escape(searchstring)  + "&highcategory=" + highcategory  + "&category=" + category  + "&gotopage=" + gotopage + "&ccustcode=" + ccustcode + "&cyear=" + cyear + "&cmonth=" + cmonth + "&attr02="+attr02;

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
			var ttitle = document.getElementById("ttitle");
			ttitle.innerText = title ;
	
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

		var highcategory = 	 (document.getElementById("cmbhighcategory")) ? document.getElementById("cmbhighcategory").value : "" ;
		var category		= 	 (document.getElementById("cmbcategory")) ? document.getElementById("cmbcategory").value : "" ;

		var url = "pop_report_reg.asp?midx="+midx+"&custcode="+tcustcode+"&flag="+tflag +"&highcategory="+highcategory +"&category="+category  ;
		var name = "" ;
		var opt = "width=658, height=680, resizable=no, scrollbars=yes, top=100, left=100";
		window.open(url, name, opt);
	}

function pop_report_regooh() {
		var midx = document.forms[0].midx.value ;
		var tcustcode = document.forms[0].tcustcode.value ;
		var tflag = document.forms[0].tflag.value ;

		var highcategory = 	 (document.getElementById("cmbhighcategory")) ? document.getElementById("cmbhighcategory").value : "" ;
		var category		= 	 (document.getElementById("cmbcategory")) ? document.getElementById("cmbcategory").value : "" ;

		var url = "pop_report_reg.asp?midx="+midx+"&custcode="+tcustcode+"&flag="+tflag +"&highcategory="+highcategory +"&category="+category  ;
		var name = "" ;
		var opt = "width=658, height=680, resizable=no, scrollbars=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function pop_report_view(ridx) {
		var midx = document.forms[0].midx.value ;

		var url = "pop_report_view.asp?ridx="+ridx+"&midx="+midx;
		var name = "" ;
		var opt = "width=658, height=680, resizable=no, scrollbars=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function pop_report_cntupdate(ridx) {
		var midx = document.forms[0].midx.value ;

		var url = "report_cntupdate_proc.asp?ridx=" + ridx+"&midx="+midx;
		var name = "" ;
		var opt = "width=0, height=0, resizable=yes, scrollbars=no, top=0, left=0";
		window.open(url, name, opt);

//			location.href="report_cntupdate_proc.asp?ridx=" + ridx+"&midx="+midx;
	}

	function pop_report_downcntupdate(ridx) {
		var midx = document.forms[0].midx.value ;

		var url = "report_downcntupdate_proc.asp?ridx=" + ridx+"&midx="+midx;
		var name = "" ;
		var opt = "width=0, height=0, resizable=yes, scrollbars=no, top=0, left=0";
		window.open(url, name, opt);

//			location.href="report_cntupdate_proc.asp?ridx=" + ridx+"&midx="+midx;
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

		//getdata("list.asp");
		left_menu_getdata(1);
		//frm.txtsearchstring.value = "" ;
	}



	function gethighcategorycombo() {
			// 광고주 콤보 박스 가져오기
			var highcategory =  (document.getElementById("chighcategory")) ? document.getElementById("chighcategory").value : "" ;
			var params = "highcategory="+highcategory;

			sendRequest("/inc/getreporthighcategory.asp", params, _gethighcategorycombo, "GET");
		}

		function _gethighcategorycombo() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
						var highcategory = document.getElementById("highcategory");
						highcategory.innerHTML = xmlreq.responseText ;
						getcategorycombo();
				}
			}
		}

		function getcategorycombo() {
			// 운영팀 콤보 박스 가져오기
			var highcategory = document.getElementById("cmbhighcategory").value;
			var category = (document.getElementById("ccategory")) ? document.getElementById("ccategory").value : "" ;

			var params = "highcategory="+highcategory+"&category="+category;

			sendRequest("/inc/getreportcategory.asp", params, _getcategorycombo, "GET");
		}

		function _getcategorycombo() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
						var category = document.getElementById("category");
						category.innerHTML = xmlreq.responseText ;
						left_menu_getdata(document.getElementById("tgotopage").value);
				}
			}
		}


		function getcustcombo_report() {
			// 광고주 콤보 박스 가져오기
			var scope = null;
			var custcode = null;
			var params = "scope="+scope+"&custcode="+custcode;
			sendRequest("/inc/getcustcombo_report.asp", params, _getcustcombo_report, "GET");
		}

		function _getcustcombo_report() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
						var custcode = document.getElementById("custcode");
						custcode.innerHTML = xmlreq.responseText ;
						left_menu_getdata(document.getElementById("tgotopage").value);
				}
			}
		}


	document.onkeydown = function() {
		if (event.keyCode == "13") return false;
	}

//	window.onload = function () {
//
//		getdata();
//
//	}

//==================================여기까지list.asp=============================================

//page,midx,custcode,FLAG,title

-->
</SCRIPT>

</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  >

<form target="scriptFrame" runat="server" defaultbutton="mybutton">

<!--#include virtual="/cust/top.asp" -->
<input type="hidden" name="actionurl" value="list.asp" >
<input type="hidden" name="midx" id="midx">
<input type="hidden" name="tcustcode" >
<input type="hidden" name="tflag" value= "midx" >
<input type="hidden" name="ttitle"  >
<input type="hidden" name="thighcategory"  >
<input type="hidden" name="tcategory"  >
<input type="hidden" name="chighcategory"  value="<%=request.cookies("cookiehighcategory")%>">
<input type="hidden" name="ccategory"   value="<%=request.cookies("cookiecategory")%>">
<input type="hidden" name="ccustcode"   value="">
<input type="hidden" name="tgotopage"   value="">
<input type="hidden" name="attr02"   value="">
  <table id="Table_01" width="1240" height="600"  border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/cust/left_report_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_report.gif" width="1030" height="65" alt=""></td>
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
            <td valign="top" colspan='2' id="srcTD">
			</td>
          </tr>
		  <tr>
            <td height="15" >&nbsp;</td>
          </tr>
          <tr>
            <td valign="top" ><div id="process"  style="text-align:center;"></div></td>
          </tr>
      </table></td>
    </tr>
	<tr>
	</tr>
  </table>
  <!--#include virtual="/bottom.asp" -->
</body>
</html>

