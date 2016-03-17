<script type="text/javascript" src="/js/ajax.js"></script>
<script type="text/javascript" src="/js/script.js"></script>
<script type="text/javascript">
<!--
	function getbrandcode() {
	// 광고주를 선택 했을때 실행
		var highcustcode = "<%=highcustcode%>";
		var seqno = "" ;
		var params = "highcustcode="+highcustcode+"&seqno="+seqno ;
		_sendRequest("/hq/outdoor/inc/getbrandcode.asp", params, _getbrandcode, "GET");
		_sendRequest("/hq/outdoor/inc/getsubbrandcode.asp",  null, _getsubbrandcode, "GET");
		_sendRequest("/hq/outdoor/inc/getthemecode.asp", null, _getthemecode, "GET");
	}

	function _getbrandcode() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var brandview = document.getElementById("brandview");
				if (brandview) {
					brandview.innerHTML = xmlreq.responseText ;
					document.getElementById("cmbseqno").attachEvent("onchange", getsubbrandcode);
					document.getElementById("cmbseqno").style.width = "155px";
					document.getElementById("cmbseqno").style.height = "240px";
				}
			}
		}
	}

	function getsubbrandcode() {
		// 브랜드를 선택 했을때 실행
		var seqno = document.getElementById("cmbseqno").value;
		var subno = "" ;
		var params = "seqno="+seqno+"&subno="+subno ;
		_sendRequest("/hq/outdoor/inc/getsubbrandcode.asp", params, _getsubbrandcode, "GET");
		_sendRequest("/hq/outdoor/inc/getthemecode.asp", null, _getthemecode, "GET");
	}

	function  _getsubbrandcode() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var subbrandview = document.getElementById("subbrandview");
				if (subbrandview) {
					subbrandview.innerHTML = xmlreq.responseText ;
					document.getElementById("cmbsubno").attachEvent("onchange", getthemecode);
					document.getElementById("cmbsubno").style.width = "155px";
					document.getElementById("cmbsubno").style.height = "240px";
				}
			}
		}
	}

	function getthemecode() {
		//tj 브랜드를 선택 했을때 실행
		var subno = document.getElementById("cmbsubno").value;
		var thmno = "" ;
		var params = "subno="+subno+"&thmno="+thmno ;
		sendRequest("/hq/outdoor/inc/getthemecode.asp", params, _getthemecode, "GET");
	}

	function _getthemecode() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var themeview = document.getElementById("themeview");
				if (themeview) {
					themeview.innerHTML = xmlreq.responseText ;
					document.getElementById("cmbthmno").style.width = "185px";
					document.getElementById("cmbthmno").style.height = "240px";
					document.getElementById("cmbthmno").attachEvent("ondblclick",
						function () {
							var clickElement = event.srcElement;
							document.getElementById("txttheme").value= clickElement.options[clickElement.selectedIndex].text;
							document.getElementById("hdnthmno").value = clickElement.value;
							setclose();
						}
					);
				}
			}
		}
	}

	function _getthemecode2() {
		if (document.getElementById("cmbthmno").value != "") {
		document.getElementById("txttheme").value= document.getElementById("cmbthmno").options[document.getElementById("cmbthmno").selectedIndex].text;
		document.getElementById("hdnthmno").value = document.getElementById("cmbthmno").value;
		setclose();
		} else {
			alert("집행할 소재를 선택하세요");
			return false;
		}
	}

	window.attachEvent("onload", getbrandcode);

//-->
</script>
<table border="0" cellpadding="0" cellspacing="0" align='center'>
	<tr>
		<th>브랜드</th>
		<th>서브 브랜드</th>
		<th>집행 소재</th>
	</tr>
	<tr>
		<td><div id="brandview"></div></td>
		<td><div id="subbrandview"></div></td>
		<td><div id="themeview"></div></td>
	</tr>
	<tr>
		<td align='right' colspan='3'> <strong><a href="#"  onclick="_getthemecode2(); return false;"> 소재선택 </a></strong> | <strong><a href="#"  onclick="setclose(); return false;"> 닫기 </a></strong> &nbsp;</td>
	</tr>
</table>
