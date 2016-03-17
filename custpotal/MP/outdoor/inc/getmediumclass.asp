<script type="text/javascript" src="/js/ajax.js"></script>
<script type="text/javascript" src="/js/script.js"></script>
<script type="text/javascript">
<!--
	function gethighclass() {
		_sendRequest("/MP/outdoor/inc/gethighclass.asp", null, _gethighclass, "GET");
		_sendRequest("/MP/outdoor/inc/getmiddleclass.asp", null, _getmiddleclass, "GET");
		_sendRequest("/MP/outdoor/inc/getlowclass.asp", null, _getlowclass, "GET");
		_sendRequest("/MP/outdoor/inc/getdetailclass.asp", null, _getdetailclass, "GET");
		}

	function _gethighclass() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var highclass = document.getElementById("highclass");
				highclass.innerHTML = xmlreq.responseText ;
				document.getElementById("cmbhighclass").attachEvent("onchange", getmiddleclass);
				document.getElementById("cmbhighclass").style.width = "110px";
				document.getElementById("cmbhighclass").style.height = "240px";
			}
		}
	}

	function getmiddleclass(crud) {
	// 중분류 코드
		if (typeof crud == "object") crud = 'r';
		var highclasscode = document.getElementById("cmbhighclass").value;
		var params = "crud="+crud+"&highclasscode="+highclasscode ;
		_sendRequest("/MP/outdoor/inc/getmiddleclass.asp", params, _getmiddleclass, "GET");
		_sendRequest("/MP/outdoor/inc/getlowclass.asp", null, _getlowclass, "GET");
		_sendRequest("/MP/outdoor/inc/getdetailclass.asp", null, _getdetailclass, "GET");
	}

	function _getmiddleclass() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var middleclass = document.getElementById("middleclass");
				middleclass.innerHTML = xmlreq.responseText ;
				document.getElementById("cmbmiddleclass").attachEvent("onchange", getlowclass);
				document.getElementById("cmbmiddleclass").style.width = "120px";
				document.getElementById("cmbmiddleclass").style.height = "240px";
			}
		}
	}

	function getlowclass(crud) {
	// 소분류 코드
		if (typeof crud == "object") crud = 'r';
		var middleclasscode = document.getElementById("cmbmiddleclass").value;
		var params = "crud="+crud+"&middleclasscode="+middleclasscode;
		_sendRequest("/MP/outdoor/inc/getlowclass.asp", params, _getlowclass, "GET");
		_sendRequest("/MP/outdoor/inc/getdetailclass.asp", null, _getdetailclass, "GET");
	}

	function _getlowclass() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var lowclass = document.getElementById("lowclass");
				lowclass.innerHTML = xmlreq.responseText ;
				document.getElementById("cmblowclass").attachEvent("onchange", getdetailclass);
				document.getElementById("cmblowclass").attachEvent("ondblclick", setvalue);
				document.getElementById("cmblowclass").style.width = "120px";
				document.getElementById("cmblowclass").style.height = "240px";
			}
		}
	}

	function getdetailclass(crud) {
	// 세분류 코드
		if (typeof crud == "object") crud = 'r';
		var lowclasscode = document.getElementById('cmblowclass').value;
		var params = "crud="+crud+"&lowclasscode="+lowclasscode;
		sendRequest("/MP/outdoor/inc/getdetailclass.asp", params, _getdetailclass, "GET");

	}

	function _getdetailclass() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var detailclass = document.getElementById("detailclass");
				detailclass.innerHTML = xmlreq.responseText ;
				document.getElementById("cmbdetailclass").attachEvent("ondblclick", setvalue);
				document.getElementById("cmbdetailclass").style.width = "145px";
				document.getElementById("cmbdetailclass").style.height = "240px";
			}
		}
	}

	function setvalue() {
		var clickElement = event.srcElement;
		document.getElementById("txtcategoryname").value = clickElement.options[clickElement.selectedIndex].text;
		document.getElementById("hdncategoryidx").value = clickElement.value;
		setclose();
	}

	function setvalue2() {
		var setvalue, setdesc ;
		if (!(document.getElementById("cmbdetailclass").value || document.getElementById("cmblowclass").value)) {alert("등록할 매체분류를 선택하세요"); return false;}
		if (document.getElementById("cmbdetailclass").value) {
			setdesc = document.getElementById("cmbdetailclass").options[document.getElementById("cmbdetailclass").selectedIndex].text;
			setvalue = document.getElementById("cmbdetailclass").value ;
		} else {
			setdesc = document.getElementById("cmblowclass").options[document.getElementById("cmblowclass").selectedIndex].text;
			setvalue = document.getElementById("cmblowclass").value;
		}
		document.getElementById("txtcategoryname").value = setdesc ;
		document.getElementById("hdncategoryidx").value = setvalue ;
		setclose();
	}

	window.attachEvent("onload", gethighclass);

//-->
</script>
<table border="0" cellpadding="0" cellspacing="0" align='center'>
	<tr>
		<th>대분류</td>
		<th>중분류</th>
		<th>소분류</th>
		<th>세분류</th>
	</tr>
	<tr>
		<td><div id="highclass"></div></td>
		<td><div id="middleclass"></div></td>
		<td><div id="lowclass"></div></td>
		<td><div id="detailclass"></div></td>
	</tr>
	<tr>
		<td align='right' colspan='4'> <strong><a href="#" onclick="setvalue2(); return false;">매체선택</a></strong> | <strong><a href="#"  onclick="setclose(); return false;"> 닫기 </a></strong> &nbsp;</td>
	</tr>
</table>