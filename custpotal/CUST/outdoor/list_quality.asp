<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<html>
<head>
<title>�Ƣ� SK M&C | Media Management System �Ƣ� </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/cust/outdoor/style.css" rel="stylesheet" type="text/css">
<script type='text/javascript' src='/js/ajax.js'></script>
<script type='text/javascript' src='/js/script.js'></script>
<script type="text/javascript">
<!--
	function getquality(crud) {
		var name = document.getElementById("txtquality").value;
		var value = document.getElementById("hdnquality").value;
		if (name == "") {
			 switch (crud) {
				case "c":
					alert("�߰��Ϸ��� ������ �Է��ϼ���.");
					break;
				case "u":
					alert("������ ������ �����ϼ���.");
					break;
				case "d":
					alert("�����Ϸ��� ������ �����ϼ���");
					break;
			 }
			 return false;
		 }
		 if (crud == "d") {if (!confirm("������ ������ �����Ͻðڽ��ϱ�?")) return false;}
		var params = "crud="+crud+"&name="+encodeURIComponent(name)+"&value="+value;
		sendRequest("/cust/outdoor/inc/getquality.asp", params, _getquality, "GET");

	}

	function _getquality(element, index, arrary) {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var qualityview = document.getElementById("qualityview");
				qualityview.innerHTML = xmlreq.responseText ;
				document.getElementById("cmbquality").attachEvent("onchange", getcode);
				document.getElementById("txtquality").value = "";
				document.getElementById("hdnquality").value = "";
			}
		}
	}

	function getcode() {
		var code = document.getElementById("cmbquality").options[document.getElementById("cmbquality").selectedIndex].text;
		document.getElementById("txtquality").value = code;
		document.getElementById("hdnquality").value = document.getElementById("cmbquality").value ;
	}

	window.onload = function() {
		var params = "crud=r";
		sendRequest("/cust/outdoor/inc/getquality.asp", params, _getquality, "GET");
	}

	window.onunload = function () {
		document.getElementById("cmbquality").detachEvent("onchange", getcode);
	}

	function _debug() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var debugConsole = document.getElementById("debugConsole");
					debugConsole.innerHTML = xmlreq.responseText ;
			}
		}
	}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<form name="form1" method="post" action="">
<INPUT TYPE="hidden" NAME="menunum" value="<%=request("menunum")%>">
<!--#include virtual="/cust/top.asp" -->
  <table id="Table_01" width="1240"  border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/cust/left_outdoor_menu.asp"--></td>
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
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="boldtitle"> �������� </span></TD>
				<TD width="50%" align="right"><span class="navigator" > ���ܱ��� &gt; ��ü���� &gt; ��������</span></TD>
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
		  <td height="30" valign='middle'> <img src='/images/m_add.gif' width='14' height='15' alt="�߰�" align='absmiddle'> �߰�  <img src='/images/m_edit.gif' width='16' height='15' alt="����"> ���� <img src='/images/m_delete.gif' width='15' height='15' alt="����"> ���� </td>
          </tr>
          <tr>
            <td >
			<table border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td height='30' ><img src='/images/m_arw.gif' width='5' height='8' hspace='2'> ������Ȳ </td>
				</tr>
				<tr>
					<td width='270'><div id='qualityview'></div><input type="text" id="txtquality" style='width:208px;' maxlength='20'/>   <a href="#" onclick="getquality('c'); return false;"><img src='/images/m_add.gif' width='14' height='15' alt="�߰�" ></a>  <a href="#" onclick="getquality('u'); return false;"><img src='/images/m_edit.gif' width='16' height='15' alt="����"></a> <a href="#" onclick="getquality('d'); return false;"><img src='/images/m_delete.gif' width='15' height='15' alt="����"></a> <input type="hidden" id="hdnquality" name='hdnquality' /></td>
				</tr>
			</table>
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

<div id='debugConsole'></div>