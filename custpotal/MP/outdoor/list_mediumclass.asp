<!--#include virtual="/mp/outdoor/inc/Function.asp" -->
<%

%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/mp/outdoor/style.css" rel="stylesheet" type="text/css">
<script type='text/javascript' src='/js/ajax.js'></script>
<script type='text/javascript' src='/js/script.js'></script>
<script type="text/javascript">
<!--
	// combo item  cmbhighclass, cmbmiddleclass, cmblowclass, cmbdetailclass
	function gethighclass(crud) {
	// 대분류 코드
		if (typeof crud == "object") crud = 'r';
		if (crud == 'd') {
			if (isdelete()) {
				alert('하위 분류가 존재 하므로 삭제하실 수 없습니다.\n\n하위 분류를 먼저 삭제하십시오');
				return false;
			} else {
				if (!confirm("선택한 분류항목을 삭제하시겠습니까?")) {return false;}
			}
		}
		var highclasscode = document.getElementById("hdnhighclasscode").value
		var highclassname = document.getElementById('txthighclassname').value;
		var params = "crud="+crud+"&highclasscode="+highclasscode+"&highclassname="+encodeURIComponent(highclassname) ;
		_sendRequest("/mp/outdoor/inc/gethighclass.asp", params, _gethighclass, "GET");
		_sendRequest("/mp/outdoor/inc/getmiddleclass.asp", null, _getmiddleclass, "GET");
		_sendRequest("/mp/outdoor/inc/getlowclass.asp", null, _getlowclass, "GET");
		_sendRequest("/mp/outdoor/inc/getdetailclass.asp", null, _getdetailclass, "GET");
	}

	function _gethighclass() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var highclass = document.getElementById("highclass");
				highclass.innerHTML = xmlreq.responseText ;
				document.getElementById("cmbhighclass").attachEvent("onchange", setvalue);
				document.getElementById("cmbhighclass").attachEvent("onchange", getmiddleclass);
				document.getElementById("txthighclassname").value = "";
				document.getElementById("hdnhighclasscode").value = "";
			}
		}
	}

	function getmiddleclass(crud) {
	// 중분류 코드
		if (typeof crud == "object") crud = 'r';
		if (crud == 'd') {
			if (isdelete()) {
				alert('하위 분류가 존재 하므로 삭제하실 수 없습니다.\n\n하위 분류를 먼저 삭제하십시오');
				return false;
			} else {
				if (!confirm("선택한 분류항목을 삭제하시겠습니까?")) {return false;}
			}
		}
		var highclasscode = document.getElementById("cmbhighclass").value;
		var middleclasscode = document.getElementById("hdnmiddleclasscode").value;
		var middleclassname = document.getElementById('txtmiddleclassname').value;
		var params = "crud="+crud+"&highclasscode="+highclasscode+"&middleclasscode="+middleclasscode+"&middleclassname="+encodeURIComponent(middleclassname) ;
		_sendRequest("/mp/outdoor/inc/getmiddleclass.asp", params, _getmiddleclass, "GET");
		_sendRequest("/mp/outdoor/inc/getlowclass.asp", null, _getlowclass, "GET");
		_sendRequest("/mp/outdoor/inc/getdetailclass.asp", null, _getdetailclass, "GET");

	}

	function _getmiddleclass() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var middleclass = document.getElementById("middleclass");
				middleclass.innerHTML = xmlreq.responseText ;
				document.getElementById("cmbmiddleclass").attachEvent("onchange", setvalue);
				document.getElementById("cmbmiddleclass").attachEvent("onchange", getlowclass);
				document.getElementById("txtmiddleclassname").value = "";
				document.getElementById("hdnmiddleclasscode").value = "";
			}
		}
	}

	function getlowclass(crud) {
	// 소분류 코드
		if (typeof crud == "object") crud = 'r';
		if (crud == 'd') {
			if (isdelete()) {
				alert('하위 분류가 존재 하므로 삭제하실 수 없습니다.\n\n하위 분류를 먼저 삭제하십시오');
				return false;
			} else {
				if (!confirm("선택한 분류항목을 삭제하시겠습니까?")) {return false;}
			}
		}
		var middleclasscode = document.getElementById("cmbmiddleclass").value;
		var lowclasscode = document.getElementById('hdnlowclasscode').value;
		var lowclassname = document.getElementById('txtlowclassname').value;
		var params = "crud="+crud+"&middleclasscode="+middleclasscode+"&lowclasscode="+lowclasscode+"&lowclassname="+encodeURIComponent(lowclassname) ;
		_sendRequest("/mp/outdoor/inc/getlowclass.asp", params, _getlowclass, "GET");
		_sendRequest("/mp/outdoor/inc/getdetailclass.asp", null, _getdetailclass, "GET");
	}

	function _getlowclass() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var lowclass = document.getElementById("lowclass");
				lowclass.innerHTML = xmlreq.responseText ;
				document.getElementById("cmblowclass").attachEvent("onchange", setvalue);
				document.getElementById("cmblowclass").attachEvent("onchange", getdetailclass);
				document.getElementById("txtlowclassname").value = "";
				document.getElementById("hdnlowclasscode").value = "";
			}
		}
	}

	function getdetailclass(crud) {
	// 세분류 코드
		if (typeof crud == "object") crud = 'r';
		if (crud == 'd') {
				if (!confirm("선택한 분류항목을 삭제하시겠습니까?")) {return false;}
		}
		var lowclasscode = document.getElementById('cmblowclass').value;
		var detailclasscode = document.getElementById('hdndetailclasscode').value;
		var detailclassname = document.getElementById('txtdetailclassname').value;
		var params = "crud="+crud+"&lowclasscode="+lowclasscode+"&detailclasscode="+detailclasscode+"&detailclassname="+encodeURIComponent(detailclassname) ;
		sendRequest("/mp/outdoor/inc/getdetailclass.asp", params, _getdetailclass, "GET");

	}

	function _getdetailclass() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var detailclass = document.getElementById("detailclass");
				detailclass.innerHTML = xmlreq.responseText ;
				document.getElementById("cmbdetailclass").attachEvent("onchange", setvalue);
				document.getElementById("txtdetailclassname").value = "";
				document.getElementById("hdndetailclasscode").value = "";
			}
		}
	}

	function setvalue() {
		var clickElement = event.srcElement;
		var targetElement = document.getElementsByTagName("input");
		for (var i = 0 ; i < targetElement.length ; i++) {
			if (targetElement[i].type == "text") {
				if (targetElement[i].getAttribute("className") == clickElement.getAttribute("className")) targetElement[i].setAttribute("value", clickElement.options[clickElement.selectedIndex].text);
			}
			if (targetElement[i].type == "hidden") {
				if (targetElement[i].getAttribute("className") == clickElement.getAttribute("className")) targetElement[i].setAttribute("value", clickElement.value);
			}
		}
	}

	function _debug() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var debugConsole = document.getElementById("debugConsole");
					debugConsole.innerHTML = xmlreq.responseText ;
			}
		}
	}

	function isdelete() {
		var className  =event.srcElement.getAttribute("className");
		var aryclass = ['highclass', 'middleclass', 'lowclass','detailclass',''];
		for (var i = 0 ; i < aryclass.length ; i++) {
			if (className == aryclass[i]) {
				var tag = document.getElementsByTagName("select");
				for (var j= 0; j < tag.length ; j++) {
					if (tag[j].getAttribute("className") == aryclass[i+1]) { return tag[j].options.length; }
				}
			}
		}
	}

	window.onload = function () {
		_sendRequest("/mp/outdoor/inc/gethighclass.asp", null, _gethighclass, "GET");
		_sendRequest("/mp/outdoor/inc/getmiddleclass.asp", null, _getmiddleclass, "GET");
		_sendRequest("/mp/outdoor/inc/getlowclass.asp", null, _getlowclass, "GET");
		_sendRequest("/mp/outdoor/inc/getdetailclass.asp", null, _getdetailclass, "GET");
	}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<form name="form1" method="post" action="">
<INPUT TYPE="hidden" NAME="menunum" value="<%=request("menunum")%>">
<!--#include virtual="/mp/top.asp" -->
  <table id="Table_01" width="1240"  border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/mp/left_outdoor_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_outdoor.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td align="left" valign="top" height="600" >
	  <table width="1002" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="19">&nbsp;</td>
          </tr>
          <tr>
            <td height="17"><TABLE  width="100%">
            <TR>
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="boldtitle"> 분류관리 </span></TD>
				<TD width="50%" align="right"><span class="navigator" > 옥외광고 &gt; 매체관리 &gt; 분류관리</span></TD>
            </TR>
            </TABLE>
			</td>
          </tr>
          <tr>
            <td height="15">&nbsp;</td>
          </tr>
          <tr>
            <td>
			<table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
                  <td align="left" background="/images/bg_search.gif"><!-- search section  -->&nbsp; </td>
                  <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="30" valign='middle'><img src='/images/m_add.gif' width='14' height='15' alt="추가" align='absmiddle'> 추가  <img src='/images/m_edit.gif' width='16' height='15' alt="수정"> 수정 <img src='/images/m_delete.gif' width='15' height='15' alt="제"> 삭제 </td>
          </tr>
          <tr>
            <td >
			<table border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td width='257' height='30'><img src='/images/m_arw.gif' width='5' height='8' hspace='2'> 대분류 </td>
					<td width='257'><img src='/images/m_arw.gif' width='5' height='8' hspace='2'> 중분류 </td>
					<td width='257'><img src='/images/m_arw.gif' width='5' height='8' hspace='2'> 소분류 </td>
					<td width='257'><img src='/images/m_arw.gif' width='5' height='8' hspace='2'> 세분류 </td>
				</tr>
				<tr>
					<td><div id="highclass"></div></td>
					<td><div id="middleclass"></div></td>
					<td><div id="lowclass"></div></td>
					<td><div id="detailclass"></div></td>
				</tr>
				<tr>
					<td height='40' align='bottom'><input type="text" id="txthighclassname"  name="txthighclassname" class="highclass" style='width:200px;'/> <a href="#" onclick="gethighclass('c'); return false;"><img src='/images/m_add.gif' width='14' height='15' alt="대분류 추가"></a><a href="#" onclick="gethighclass('u'); return false;"><img src='/images/m_edit.gif' width='16' height='15' hspace=2 alt="대분류 수정"></a><a href="#" onclick="gethighclass('d'); return false;"><img src='/images/m_delete.gif' width='15' height='15' alt="대분류 삭제" class ="highclass"></a><input type="hidden" class ="highclass" id="hdnhighclasscode" name="hdnhighclasscode" /></td>
					<td align='bottom'> <input type="text" id="txtmiddleclassname"  name="txtmiddleclassname" class="middleclass" style='width:200px;'/> <a href="#" onclick="getmiddleclass('c'); return false;"><img src='/images/m_add.gif' width='14' height='15' alt="중분류 추가"></a><a href="#" onclick="getmiddleclass('u'); return false;"><img src='/images/m_edit.gif' width='16' height='15' hspace=2 alt="중분류 수정"></a><a href="#" onclick="getmiddleclass('d'); return false;"><img src='/images/m_delete.gif' width='15' height='15' alt="중분류 삭제" class="middleclass" ></a><input type="hidden" id="hdnmiddlecode" name="hdnmiddleclasscode" class="middleclass" /></td>
					<td align='bottom'> <input type="text" id="txtlowclassname"  name="txtlowclassname" style='width:200px;' class="lowclass"/> <a href="#" onclick="getlowclass('c'); return false;"><img src='/images/m_add.gif' width='14' height='15' alt="소분류 추가"></a><a href="#" onclick="getlowclass('u'); return false;"><img src='/images/m_edit.gif' width='16' height='15' hspace=2 alt="소분류 수정"></a><a href="#" onclick="getlowclass('d'); return false;"><img src='/images/m_delete.gif' width='15' height='15' alt="소분류 삭제"  class="lowclass"></a><input type="hidden" id="hdnlowclasscode" name="hdnlowclasscode" class="lowclass" /></td>
					<td align='bottom'> <input type="text" id="txtdetailclassname" name="txtdetailclassname" style='width:200px;' class="detailclass"/><a href="#" onclick="getdetailclass('c'); return false;"> <img src='/images/m_add.gif' width='14' height='15' alt="세분류 추가"></a> <a href="#" onclick="getdetailclass('u'); return false;"><img src='/images/m_edit.gif' width='16' height='15' hspace=2 alt="세분류 수정"></a><a href="#" onclick="getdetailclass('d'); return false;"><img src='/images/m_delete.gif' width='15' height='15' alt="세분류 삭제" class="detailclass"><input type="hidden" id="hdndetailclasscode" name="hdndetailclasscode" class="detailclass"/></a></td>
				</tr>
			</table>

<div id='debugConsole'></div>
			</td>
          </tr>
      </table>
	  </td>
    </tr>
  </table>
<!--#include virtual="bottom.asp" -->
</form>
</body>
</html>