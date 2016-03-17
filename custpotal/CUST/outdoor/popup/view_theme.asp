<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<%
'	Call getquerystringparameter
	Dim pmdidx : pmdidx = request("mdidx")
	Dim pside : pside = request("side")
	If pside = "" Then pside = "F"
	Dim pcustcode : pcustcode = request("custcode")
	dim phighcustcode : phighcustcode = request("highcustcode")
	Dim pcyear : pcyear = request("cyear")
	Dim pcmonth : pcmonth = request("cmonth")
	If pcyear = "" Then pcyear = Year(date)
	If pcmonth = "" Then pcmonth = setmonth(Month(date))
	Dim sql : sql = "select a.startdate, a.enddate, d.highseqname, c.subname, b.thmname, a.thmno, a.no, a.seq, a.cyear, a.cmonth from wb_subseq_exe a inner join wb_subseq_dtl b on a.thmno = b.thmno inner join wb_subseq_mst c on b.subno = c.subno inner join sc_subseq_hdr d on c.seqno = d.highseqno where mdidx = " & pmdidx & " and side = '" & pside &"' order by seq desc"
	'response.write sql

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adcmdText
	Dim rs : Set rs = cmd.execute
	Set cmd = Nothing

	Dim startdate
	Dim enddate
	Dim seqname
	Dim subname
	Dim thmname
	Dim thmno
	Dim subno
	Dim seqno
	Dim no
	Dim seq
	Dim cyear
	Dim cmonth

	If Not rs.eof Then
		Set startdate = rs(0)
		Set enddate = rs(1)
		Set seqname = rs(2)
		Set subname = rs(3)
		Set thmname = rs(4)
		Set thmno = rs(5)
		Set no = rs(6)
		Set seq = rs(7)
		Set cyear = rs(8)
		Set cmonth = rs(9)
	End If
%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<link href="/cust/outdoor/style.css" rel="stylesheet" type="text/css">
	<title>▒▒ SK M&C | Media Management System ▒▒  </title>
	<script type='text/javascript' src='/js/ajax.js'></script>
	<script type='text/javascript' src='/js/script.js'></script>
    <script type="text/javascript" src="/js/calendar.js"></script>
	<script type="text/javascript">
	<!--

		function get() {
			var params = "";
			sendRequest(url, params, _get, "GET");
		}

		function _get() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
				}
			}
		}



		// theme Layer
		function getbrandcode() {
		// 광고주를 선택 했을때 실행
			var highcustcode = "<%=phighcustcode%>";
			var seqno = document.getElementById("seqno").value ;
			var params = "highcustcode="+highcustcode+"&seqno="+seqno ;
			_sendRequest("/cust/outdoor/inc/getbrandcode.asp", params, _getbrandcode, "GET");
			_sendRequest("/cust/outdoor/inc/getsubbrandcode.asp",  null, _getsubbrandcode, "GET");
			_sendRequest("/cust/outdoor/inc/getthemecode.asp", null, _getthemecode, "GET");
		}

		function _getbrandcode() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
					var brandview = document.getElementById("brandview");
					if (brandview) {
						brandview.innerHTML = xmlreq.responseText ;
						document.getElementById("cmbseqno").attachEvent("onchange", getsubbrandcode);
						document.getElementById("cmbseqno").style.width = "155px";
						document.getElementById("cmbseqno").style.height = "340px";
					}
				}
			}
		}

		function getsubbrandcode() {
			// 브랜드를 선택 했을때 실행
			var seqno = document.getElementById("cmbseqno").value;
			var subno = document.getElementById("subno").value ;
			var params = "seqno="+seqno+"&subno="+subno ;
			_sendRequest("/cust/outdoor/inc/getsubbrandcode.asp", params, _getsubbrandcode, "GET");
			_sendRequest("/cust/outdoor/inc/getthemecode.asp", null, _getthemecode, "GET");
		}

		function  _getsubbrandcode() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
					var subbrandview = document.getElementById("subbrandview");
					if (subbrandview) {
						subbrandview.innerHTML = xmlreq.responseText ;
						document.getElementById("cmbsubno").attachEvent("onchange", getthemecode);
						document.getElementById("cmbsubno").style.width = "155px";
						document.getElementById("cmbsubno").style.height = "340px";
					}
				}
			}
		}

		function getthemecode() {
			//tj 브랜드를 선택 했을때 실행
			var subno = document.getElementById("cmbsubno").value;
			var thmno = document.getElementById("thmno").value ;
			var params = "subno="+subno+"&thmno="+thmno ;
			sendRequest("/cust/outdoor/inc/getthemecode.asp", params, _getthemecode, "GET");
		}

		function _getthemecode() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
					var themeview = document.getElementById("themeview");
					if (themeview) {
						themeview.innerHTML = xmlreq.responseText ;
						document.getElementById("cmbthmno").style.width = "175px";
						document.getElementById("cmbthmno").style.height = "340px";
					}
				}
			}
		}

		function getaddtheme() {
			document.getElementById("txtno").value = "1" ;
			document.getElementById("seqno").value = "" ;
			document.getElementById("subno").value = "" ;
			document.getElementById("thmno").value = "" ;
			document.getElementById("crud").value = 'c' ;
			document.getElementById("themeLayer").style.display = "block";
			getbrandcode();
			getsubbrandcode();
			getthemecode();
		}

		function getEdittheme(cyear, cmonth, thmno, no, seq) {
			document.getElementById("txtno").value = no;
			document.getElementById("seqno").value = thmno.substring(0,8) ;
			document.getElementById("subno").value = thmno.substring(0,10) ;
			document.getElementById("thmno").value = thmno ;
			document.getElementById("seq").value = seq;
			document.getElementById("crud").value = 'u' ;
			document.getElementById("themeLayer").style.display = "block";
			getbrandcode();
			getsubbrandcode();
			getthemecode();
		}

		function getDeletetheme(no, seq) {
			if (confirm("선택하신 소재를 삭제하시겠습니까?")){
				document.getElementById("txtno").value = no;
			  document.getElementById("seq").value = seq;
				document.getElementById("crud").value = 'd' ;
				submitChange();
			}
		}

		function setclose() {
			document.getElementById("themeLayer").style.display = "none";
		}
	// theme Layer

		function setCRUDtheme() {
			if (document.getElementById("cmbthmno").value == "") {
				alert("소재를 선택하세요");
				document.getElementById("cmbthmno").focus();
				return false;
			}
			submitChange();
		}

		function submitChange() {
			var frm = document.forms[0];
			frm.action = "/cust/outdoor/process/db_theme.asp";
			frm.method = "post";
			frm.submit();
		}

		window.onload = function () {
			self.focus();
		}

		window.onunload = function () {
//			try {
//				window.opener.getcontactdetail();
//			} catch(e) {
//				window.close();
//			}
		}
	//-->
	</script>
</head>

<body>
<form>
<table width="550" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="absmiddle"> 광고 소재 현황 </td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="550" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="22"><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td height="458" valign='top'>
		<div style='overflow-y:scroll;width:500px;height:428px;'>
<!--  -->
		<table border="0" cellpadding="0" cellspacing="0">
			<tr height='20'>
				<td width='270' valign='top'><a href="#" onclick="getaddtheme(); return false;"><img src='/images/m_add.gif' width='14' height='15' alt="소재 추가"></a> 소재추가</td>
				<td width='270' align='right' valign='top'><img src='/images/m_edit.gif' width='16' height='15' alt="소재 정보 수정" > 수정 <img src='/images/m_delete.gif' width='16' height='15' alt="매체 정보 삭제" hspace=2> 삭제 </td>
			</tr>
		</table>
		<table border="0" cellpadding="0" cellspacing="0">
			<tr>
				<th class='normal' width="20">No</th>
				<th class='normal' width="50">년도</th>
				<th class='normal' width="50">월</th>
				<th class='normal' width="120">브랜드</th>
				<th class='normal' width="120">서브브랜드</th>
				<th class='normal' width="136">소재</th>
				<th class='normal' width="44">관리</th>
			</tr>
			<%
				Do Until rs.eof
			%>
			<tr>
				<td class='normal' style='text-align:center;vertical-align:middle;'><%=no%></td>
				<td class='normal'><%=cyear%></td>
				<td class='normal'><%=cmonth%></td>
				<td class='normal'><%=seqname%></td>
				<td class='normal'><%=subname%></td>
				<td class='normal'><%=thmname%></td>
				<td class='normal'><a href="#" onclick="getEdittheme('<%=cyear%>','<%=cmonth%>', '<%=thmno%>',<%=no%>,<%=seq%>); return false;"><img src='/images/m_edit.gif' width='16' height='15' alt="소재 정보 수정" hspace=1></a><a href="#" onclick="getDeletetheme(<%=no%>, <%=seq%>); return false;"><img src='/images/m_delete.gif' width='16' height='15' alt="매체 정보 삭제" ></a></td>
			</tr>
			<%
					rs.movenext
				Loop
			%>
		</table>
		</div>
<!--  -->
	</td>
    <td background="/images/pop_right_middle_bg.gif"></td>
  </tr>
  <tr>
    <td><img src="/images/pop_left_bottom_bg.gif" width="22" height="25"></td>
    <td background="/images/pop_center_bottom_bg.gif">&nbsp;</td>
    <td><img src="/images/pop_right_bottom_bg.gif" width="23" height="25"></td>
  </tr>
</table>
<input type="hidden" id="custcode" name="custcode" value="<%=pcustcode%>" />
<input type="hidden" id="mdidx" name="mdidx" value="<%=pmdidx%>" />
<input type="hidden" id="side" name="side" value="<%=pside%>" />
<input type="hidden" id="seq" name='seq'/>
<input type="hidden" id="seqno" />
<input type="hidden" id="subno" />
<input type="hidden" id="thmno" />
<input type="hidden" id="crud" name='crud'/>

<div id='buttonLayer' style='left:470px; top:505px;width:57px;height:18px;position:absolute; z-index:9;background-image:url(/images/btn_close.gif);cursor:hand;' onclick="window.close();"></div>
<div id="themeLayer"style="LEFT:21px; TOP:67px; width:495px; height:435px;POSITION:absolute; z-index:10; background-color:#CCCCCC;filter:alpha(opacity='100'); border: 1px solid #333333;color:#000000;font-weight:bolder;display:none;">
<table border="0" cellpadding="0" cellspacing="0" >
	<tr height='30'>
		<td colspan='2' style='padding-left:10px;'><strong>소재 관리</strong> </td>
		<td align='right' style='padding-right:10px;'> <strong><a href="#"  onclick="setclose(); return false;">44X</a></strong> </td>
	</tr>
	<tr>
		<td style='padding-left:10px;'><div id="brandview"></div></td>
		<td ><div id="subbrandview"></div></td>
		<td style='padding-right:10px;'><div id="themeview"></div></td>
	</tr>
	<tr height='30'>
		<td style='padding-left:10px;'> <strong>소재번호 </strong><input type="text" id="txtno" name='txtno' style='width:50px;' /></td>
		<td colspan='2'> <strong>집행년월 </strong> : <%Call getyear(pcyear)%> <%Call getmonth(pcmonth)%> </td>
	</tr>
	<tr height='30'>
		<td colspan='3' align='center'> <a href="#" onclick="setCRUDtheme();return false;"><strong> 저장 </strong></a> | <a href="#" onclick="setclose(); return false;"><strong> 닫기 </strong></a> </td>
	</tr>
</table>
</div>
</form>
</body>
</html>