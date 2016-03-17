<!--#include virtual="/MP/outdoor/inc/Function.asp" -->
<%
	Dim pcontidx : pcontidx = request("contidx")
	Dim pcyear : pcyear = request("cyear")
	Dim pcmonth : pcmonth = request("cmonth")
'	response.write pcontidx

	' 광고 계약 기초 정보
'	Dim sql : sql = "select c.title, c.comment, c.mediummemo, c.regionmemo, m.map, m.mdidx, t.highcustcode, c.startdate, c.enddate, c.custcode  "
'	sql = sql & " from wb_contact_mst c "
'	sql = sql & " left outer join wb_contact_md m on c.contidx = m.contidx "
'	sql = sql & "  left outer  join sc_cust_dtl t on c.custcode = t.custcode "
'	sql = sql & "  left outer  join vw_contact_exe_monthly e on c.contidx = e.contidx and e.cyear = '" & pcyear & "' and e.cmonth = '" & pcmonth & "' "
'	sql = sql & " where c.contidx = " & pcontidx

	Dim sql : sql = "select c.title, c.comment, c.mediummemo, c.regionmemo, m.map, m.mdidx, t.highcustcode, c.startdate, c.enddate, c.custcode, DBO.WB_CATEGORYIDX_FUN(m.categoryidx) categoryidx  "
	sql = sql & " from wb_contact_mst c "
	sql = sql & " left outer join wb_contact_md m on c.contidx = m.contidx "
	sql = sql & "  left outer  join sc_cust_dtl t on c.custcode = t.custcode "
	sql = sql & "  left outer  join vw_contact_exe_monthly e on c.contidx = e.contidx and e.cyear = '" & pcyear & "' and e.cmonth = '" & pcmonth & "' "
	sql = sql & " where c.contidx = " & pcontidx

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType =adCmdText
	Dim rs : Set rs = cmd.execute
	Set cmd = Nothing
	If Not rs.eof Then
		Dim title : title = rs("title")
		Dim comment : comment = rs("comment")
		Dim mediummemo : mediummemo = rs("mediummemo")
		Dim regionmemo : regionmemo = rs("regionmemo")
		Dim map : map = rs("map")
		Dim mdidx : mdidx = rs("mdidx")
		Dim highcustcode : highcustcode = rs("highcustcode")
		Dim startdate : startdate = rs("startdate")
		Dim enddate : enddate = rs("enddate")
		Dim custcode : custcode = rs("custcode")
		Dim categoryidx : categoryidx = rs("categoryidx")
		If Not IsNull(comment) Then comment = Replace(comment, Chr(13)&Chr(10), "<br>")
		If Not IsNull(mediummemo) Then  mediummemo= Replace(mediummemo, Chr(13)&Chr(10), "<br>")
		If Not IsNull(regionmemo) Then  regionmemo= Replace(regionmemo, Chr(13)&Chr(10), "<br>")
	End If
	Dim custname : custname = getcustname(custcode)
	Dim teamname : teamname = getteamname(custcode)
%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/MP/outdoor/style.css" rel="stylesheet" type="text/css">
<script type='text/javascript' src='/js/ajax.js'></script>
<script type='text/javascript' src='/js/script.js'></script>
<script type="text/javascript">
<!--
	var themeChildWin ; // 소재 관리 창
	var accountChildWin ; // 광고 비용 관리 창
	var photoChildWin; // 사진 관리 창
	var mapChildWin; // 약도 관리 창

	function getyear(){
		var contidx = "<%=pcontidx%>";
		var cyear = "<%=pcyear%>";
		var params = "contidx="+contidx+"&cyear="+cyear;
		sendRequest("/MP/outdoor/inc/getyear.asp", params, _getyear, "GET");
	}

	function _getyear(){
		var divyear = document.getElementById("yearsection");
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				divyear.innerHTML = xmlreq.responseText ;
				getmonth();
			}
		}
	}

	function getmonth() {
		var contidx = "<%=pcontidx%>";
		var cyear = document.getElementById("cyear").value ;
		var params = "contidx="+contidx+"&cyear="+cyear+"&cmonth=<%=pcmonth%>";
		sendRequest("/MP/outdoor/inc/getmonth.asp", params, _getmonth, "GET");

	}

	function _getmonth() {
		var divmonth = document.getElementById("monthsection");
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				divmonth.innerHTML = xmlreq.responseText ;
			}
		}
	}

	function getcontact() {
		//계약 일반 정보
		var contidx = "<%=pcontidx%>";
		var cyear = document.getElementById("cyear").value;
		var cmonth = document.getElementById("cmonth").value;
		var params = "contidx="+contidx+"&cyear="+cyear+"&cmonth="+cmonth;
		sendRequest("/MP/outdoor/inc/getcontactsummary_b.asp", params, _getcontact, "GET");
	}

	function _getcontact() {
		var summaryview = document.getElementById("summaryview");
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				summaryview.innerHTML = xmlreq.responseText ;
				getcontactdetail();
			} else {
				summaryview.innerHTML = "dataview : " + xmlreq.responseText ;
			}
		} else  {
				summaryview.innerHTML = "<center><img src='/images/load.gif' align='center'></center>";
		}
	}

	function getcontactdetail () {
		// 계약 매체 상세 정보 리스트 가져오기
		var side ="";
		var inputElement = document.getElementsByTagName("input");
		for (var i = 0 ; i < inputElement.length; i++ ) {
			if (inputElement[i].getAttribute("type") == "checkbox") {
				if (inputElement[i].checked) {side = inputElement[i].value;}
			}
		}
		var contidx = "<%=pcontidx%>";
		var cyear = document.getElementById("cyear").value;
		var cmonth = document.getElementById("cmonth").value;
		var params = "contidx="+contidx+"&cyear="+cyear+"&cmonth="+cmonth+"&side="+side;
		sendRequest("/MP/outdoor/inc/getcontactdetail_b.asp", params, _getcontactdetail, "GET");
	}

	function _getcontactdetail() {
		// 계약 매체 상제 정보 콜백 함수
		var dataview = document.getElementById("dataview");
		var flag = true;
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				dataview.innerHTML = xmlreq.responseText ;
				var inputElement = document.getElementsByTagName("input") ;
				for (var i = 0 ; i < inputElement.length ; i++) {
					if (inputElement[i].getAttribute("type") == "checkbox") {
						if (inputElement[i].checked) flag = false;
					}
				}
				if (flag) {
					if (document.getElementById("side")) {
						document.getElementById("side").checked = true;
					}
				}
				getcontactphoto();
			} else {
				dataview.innerHTML = "dataview : " + xmlreq.responseText ;
			}
		} else  {
				dataview.innerHTML = "<center><img src='/images/load.gif' align='center'></center>";
		}
	}

	function getcontactphoto() {
		// 계약 매체별 관리 사진 가져오기
		var cyear = document.getElementById("cyear").value;
		var cmonth = document.getElementById("cmonth").value;
		var side = document.getElementsByTagName("input");
		for (var i = 0 ; i < side.length ; i++) {
			if (side[i].getAttribute("type") == "checkbox") {
				if (side[i].checked) {side = side[i].value; }
			}
		}
		var params = "cyear="+cyear+"&cmonth="+cmonth+"&mdidx=<%=mdidx%>&side="+side;
		sendRequest("/MP/outdoor/inc/getcontactphoto.asp", params, _getcontactphoto, "GET");
	}

	function _getcontactphoto() {
		// 계약 매체별 관리 사진 콜백 함수
		var photoview = document.getElementById("photoview");
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				photoview.innerHTML = xmlreq.responseText ;
			} else {
				photoview.innerHTML = "photoview : " + xmlreq.status ;
			}
		} else {
				photoview.innerHTML = "<center><img src='/images/load.gif' align='center'></center>";
		}
	}

	function getmap() {
		// 매체 약도 관리
		var mdidx = "<%=mdidx%>";
		var url = "/MP/outdoor/popup/view_map.asp?mdidx="+mdidx ;
		var name = "winMap";
		var left = screen.width / 2 - 550 / 2;
		var top = 10;
		var opt = "width=550, height=550, resizable=no, scrollbars=no, status=yes, left="+left+"&top="+top;
		mapChildWin = window.open(url, name, opt);

	}

	function getaccount(mdidx, side) {
		//광고 비용 관리
		var url = "/MP/outdoor/popup/view_account.asp?mdidx="+mdidx+"&side="+side+"&flag=B" ;
		var name = "winAccount";
		var left = screen.width / 2 - 550 / 2;
		var top = 10;
		var opt = "width=550, height=550, resizable=no, scrollbars=no, status=yes, left="+left+"&top="+top;
		accountChildWin = window.open(url, name, opt);
	}

	function getphoto(mdidx, side) {
		// 광고 매체 사진 관리
		var lastdate = document.getElementById("lastdate").value ; // 사진 게시 마지막 날을 자동 계산하기 위한 컨트롤 값
		var url = "/MP/outdoor/popup/view_photo.asp?mdidx="+mdidx+"&side="+side+"&lastdate="+lastdate+"&cyear=<%=pcyear%>&cmonth=<%=pcmonth%>" ;
		var name = "winPhoto";
		var left = screen.width / 2 - 720 / 2;
		var top = 10;
		var opt = "width=720, height=550, resizable=no, scrollbars=no, status=yes, left="+left+"&top="+top;
		photoChildWin = window.open(url, name, opt);
	}

	function gettheme(mdidx, side) {
		// 광고 소재 관리
		var lastdate = document.getElementById("lastdate").value ; // 소재의 마지막 날을 자동 계산하기 위한 컨트롤 값
		var custcode = "<%=highcustcode%>"
		var url = "/MP/outdoor/popup/view_theme.asp?mdidx="+mdidx+"&side="+side+"&highcustcode=<%=highcustcode%>&lastdate="+lastdate+"&cyear=<%=pcyear%>&cmonth=<%=pcmonth%>" ;
		var name = "wintheme";
		var left = screen.width / 2 - 550 / 2;
		var top = 10;
		var opt = "width=550, height=550, resizable=no, scrollbars=no, status=yes, left="+left+"&top="+top;
		themeChildWin = window.open(url, name, opt);
	}

	function setitem(arg) {
		// 계약 매체 및 면 정보 세팅으로 관리 사진 초기화 및 선택 사진 가져오기
		var frm = document.forms[0];
		for (var i = 0; i < frm.side.length ; i++) {
			frm.side[i].checked = false;
		}
		arg.checked = true;
	}

	// 년,월 별 계약현황 조회
	function getsearch() {
		var frm = document.forms[0];
		frm.action= "view_b_contact.asp";
		frm.method = "post";
		frm.submit();
	}

	function getprint() {
		var url = "/MP/outdoor/print/prt_b_contact2.asp?contidx=<%=pcontidx%>&cyear=<%=pcyear%>&cmonth=<%=pcmonth%>";
		var name = "Print";
		var opt = "width=1300, resiable=yes, toolbars=yes";
		window.open(url, name, opt);
	}

	function getsavefile() {
		var filename = "<%=pcyear&pcmonth%>"+"_"+"<%=custname%>"+"_"+"<%=teamname%>"+"_"+"<%=title%>.html";
//		alert(filename);
		process.location.href = "/MP/outdoor/print/convert_html.asp?filename="+filename+"&contidx=<%=pcontidx%>&cyear=<%=pcyear%>&cmonth=<%=pcmonth%>&flag=B";

	}

	function _debug() {
		var debug = document.getElementById("debugConsole");
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				debug.innerHTML = xmlreq.responseText;
				debug.innerHTML += xmlreq.status ;
			}
		}
	}

	function preview() {
		var clickElement = event.srcElement ;
		var url = "/MP/outdoor/inc/viewPhoto.asp?src="+clickElement.src;
		var name = "preview";
		var left = screen.width / 2 - 600 / 2;
		var top = screen.height / 2 - 450 / 2;
		var opt = "width=600; height=450; resizable=no, left="+left+", top="+top
		window.open (url, name, opt);
	}

	// 정산후 매체 정보 수정
	function setmedium() {
		var mdidx = document.getElementById("mdidx").value;
		var cyear = document.getElementById("cyear").value;
		var cmonth = document.getElementById("cmonth").value;
		var side = "";
		if (event.srcElement.getAttribute("className")) { side = event.srcElement.getAttribute("className");}
		var url = "/MP/outdoor/popup/view_b_medium3.asp?contidx=<%=pcontidx%>&cyear="+cyear+"&cmonth="+cmonth+"&side="+side+"&mdidx="+mdidx;
		var name = "medium";
		var left = screen.width / 2 - 540 / 2;
		var top = screen.height / 2 - (210 / 2);
		var opt = "width=540, height=206, resizable=no, scrollbars=no, status=yes,left="+left+"&top="+top;
		window.open(url, name, opt);

	}

	// 매체 정보 등록 팝업
	function getmedium(crud) {
		var mdidx = document.getElementById("mdidx").value;
		if (!mdidx) {alert("매체를 먼저 등록하신 후 면정보를 추가하세요"); getmedium2(); return false;}
		var side = "";
		if (event.srcElement.getAttribute("className")) { side = event.srcElement.getAttribute("className");}
		var cyear = document.getElementById("cyear").value;
		var cmonth = document.getElementById("cmonth").value;
		// 매체 정보가 없다면 view_b_medium2 (매체정보 등록) 페이지로 이동한다
		var url = "/MP/outdoor/popup/view_b_medium.asp?contidx=<%=pcontidx%>&cyear="+cyear+"&cmonth="+cmonth+"&side="+side+"&crud="+crud+"&mdidx="+mdidx;
		var name = "medium";
		var left = screen.width / 2 - 540 / 2;
		var top = screen.height / 2 - 596 / 2;
		var opt = "width=540, height=390, resizable=no, scrollbars=no, status=yes,left="+left+"&top="+top;
		window.open(url, name, opt);
	}

	function getmedium2() {
		var name = "medium";
		var mdidx = document.getElementById("mdidx").value;
		var url = "/MP/outdoor/popup/view_b_medium2.asp?contidx=<%=pcontidx%>&mdidx="+mdidx;
		var left = screen.width / 2 - 540 / 2;
		var top = screen.height / 2 - 596 / 2;
		var opt = "width=540, height=390, resizable=no, scrollbars=no, status=yes,left="+left+"&top="+top;
		window.open(url, name, opt);
	}

	function getreportphoto() {
		var url = "/MP/outdoor/popup/view_report_photo.asp?contidx=<%=pcontidx%>&cyear=<%=pcyear%>&cmonth=<%=pcmonth%>";
		var name = "Photo";
		var left = screen.width / 2 - 720 / 2;
		var top = 10;
		var opt = "width=720, height=550, resizable=no, scrollbars=no, status=yes, left="+left+"&top="+top;
		window.open(url, name, opt);
	}

	function getvalidation() {
		var contidx = "<%=pcontidx%>"
		var categoryidx = "<%=categoryidx%>"

		if (categoryidx == 9)
		{
			var url = "/MP/outdoor/validation_led.asp?contidx="+contidx ;
		}else if (categoryidx == 10)
		{
			var url = "/MP/outdoor/validation_neon.asp?contidx="+contidx ;
		}else if (categoryidx == 11)
		{
			var url = "/MP/outdoor/validation_board.asp?contidx="+contidx ;
		}else if (categoryidx == 135)
		{
			var url = "/MP/outdoor/validation_etc.asp?contidx="+contidx ;
		}

		var name = "validation";
		var left = screen.width / 2 - 840 / 2;
		var top = 5;
		var opt = "width=840, height=800, resizable=yes, scrollbars=yes, status=yes, left="+left+"&top="+top;
		window.open(url, name, opt);
	}

	function viewMap() {
		var clickElement = event.srcElement ;
//		alert(clickElement.src);
		var url = "/MP/outdoor/inc/viewPhoto.asp?src="+clickElement.src;
		var name = "map";
		var left = screen.width / 2 - 600 / 2;
		var top = screen.height / 2 - 450 / 2;
		var opt = "width=600; height=450; resizable=no, left="+left+", top="+top
		window.open (url, name, opt);
	}

	window.onload = function () {
		self.focus();
		window.attachEvent("onload", getcontact);
		_sendRequest("/MP/outdoor/inc/getyear.asp", "contidx=<%=pcontidx%>&cyear=<%=pcyear%>", _getyear, "GET");
		var cyear = document.getElementById("cyear").value;
		_sendRequest("/MP/outdoor/inc/getmonth.asp", "contidx=<%=pcontidx%>&cyear=<%=pcyear%>&cmonth=<%=pcmonth%>", _getmonth, "GET");
		document.getElementById("cyear").attachEvent("onchange", getmonth);
	}

	window.onunload = function() {
		if (themeChildWin) themeChildWin.close();
		if (accountChildWin) accountChildWin.close();
		if (photoChildWin) photoChildWin.close();
	}

//-->
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  >
<form>
<input type="hidden" id="contidx" name='contidx' value="<%=pcontidx%>" />
<input type="hidden" id="mdidx" name='mdidx' value="<%=mdidx%>" />
<!-- 계약 헤더 이미지 -->
<table width="1240"  class="title" align="center">
	<tr>
		<td><img src="/images/pop_top.gif" width="1240" height="60" align="absmiddle"></td>
	</tr>
</table>
<!-- // 계약 헤더 이미지 -->

<!-- 계약 타이틀 테이블 -->
<table width="1024"   align="center" >
	<tr>
		<td class="title"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><%=title%> </td>
	</tr>
</table>
<!-- // 계약 타이틀 테이블 -->

<!-- 검색 조건 및 계약 관리 툴 -->
<table width="1024" height="35"   align="center" border=0 cellspacing=0 cellpadding=0 style="margin-top:15px;">
	<tr >
		<td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
		<td width="499" class="search"><img src="/images/icon_dot_search.gif" width="4" height="3" align="absmiddle" hspace='5'><span id="yearsection"></span><span id="monthsection"></span>&nbsp;<a href="#" onclick="getsearch();"><img src="/images/btn_search.gif" width="39" height="20" align="absmiddle" alt="광고 년월별 계약현황"></a> <%Call getReportFile(rs("mdidx"), pcyear, pcmonth)%> <a href="#" onclick="getsavefile(); return false;"><img src="/images/icon_pdf.gif" width="16" height="16"  alt="Pdf 관리" align='absmiddle'  hspace='5'></a> <a href="#"  onclick='getprint(); return false;'><img src="/images/m_print.gif" width="16" height="16"  alt="출력 관리" align='absmiddle'></a> </td>
		<td width="499" class="search" align="right"><!--<% If IsNull(rs("categoryidx")) Then response.write "" Else response.write "<a href='#' onclick='getvalidation(); return false;'><img src='/images/btn_md_validation.gif' width='78' height='18'  alt='매체 평가'></a>" %>  <a href="#"  onclick='getreportphoto(); return false;'><img src="/images/btn_report_photo.gif" width="78" height="18"  alt="보고서용 사진 관리"></a> <a href="#"  onclick='getmedium2(); return false;'><img src="/images/btn_medium_mng.gif" width="78" height="18"  alt="매체 관리"></a> --><a href="#" onclick="self.close(); return false;"><img src="/images/btn_close.gif" width="57" height="18" alt="창 닫기"></a></td>
		<td width="13"><img src="/images/bg_search_right.gif" width="13" height="35" alt="좌측 라운드 처리"></td>
	</tr>
</table>
<!-- // 검색 조건 및 계약 관리 툴 -->

<!-- 계약 기초 정보 -->
<div id="summaryview"><center><img src='/images/load.gif' align='center'></center></div>
<!-- // 계약 기초 정보 -->

<!-- 매체 계약 상세 정보  -->
	<div id="dataview"><center><img src='/images/load.gif' align='center'></center></div>
<!-- 매체 계약 상세 정보  -->

<!-- 매체별 관리 사진 -->
<div id='photoview'><center><img src='/images/load.gif' align='center'></center></div>
<!-- //매체별 관리 사진 -->

<!-- 계약 특성 및 약도 -->
<table width="1024" align="center" style="margin-top:10px;">
	<tr>
	  <th class="title" width='100'>매체특성</td>
	  <td width='684'  class="context"><%=mediummemo %></td>
	  <td rowspan='3' width='230'  class="context" id='map'><%If IsNull(rs("map")) Then response.write "<img src='/images/noimage.gif' width='230' height='190'>" Else response.write "<a href='#' onclick='viewMap(); return false;'><img src='/pds/media/"&rs("map")&"' width='230' height='190' onclick=></a>"%></td>
	</tr>
	<tr>
	  <th class="title">지역특성</td>
	  <td  class="context"><%=regionmemo%></td>
	</tr>
	<tr>
	  <th class="title">특이사항</td>
	  <td  class="context"><%=comment%></td>
	</tr>
</table>
<!-- // 계약 특성 및 약도 -->


<!-- 카피라이터 이미지 -->
<table width="1024" align="center">
  <tr>
    <td ><img src="/images/pop_bottom.gif" width="1240" height="71" align="absmiddle"></td>
  </tr>
</table>
<!-- // 카피라이터 이미지 -->
</form>
</body>
</html>

<div id="debugConsole"></div>
<iframe src="about:blank" width=0 height=0 frameborder=0 name='process'></iframe>