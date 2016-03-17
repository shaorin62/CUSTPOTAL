<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<%
		Dim sql : sql = "select highcustcode, custname from sc_cust_hdr where medflag='A' and use_flag=1 order by custname"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		Set cmd = Nothing

		Sub getcustcode
			Do Until rs.eof
				response.write "<option value='" & rs("highcustcode") & "'>" & rs("custname") & "</option>"
				rs.movenext
			Loop
		End Sub
%>
<html>
<head>
<title>�Ƣ� SK M&C | Media Management System �Ƣ� </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/cust/outdoor/style.css" rel="stylesheet" type="text/css">
<script type='text/javascript' src='/js/ajax.js'></script>
<script type='text/javascript' src='/js/script.js'></script>
<script type="text/javascript">
<!--
	function getbrandcode() {
	// �����ָ� ���� ������ ����
		var custcode = document.getElementById("cmbcustcode").value;
		var seqno = "" ;
		var params = "custcode="+custcode+"&seqno="+seqno ;
		_sendRequest("/cust/outdoor/inc/getbrandcode.asp", params, _getbrandcode, "GET");
		_sendRequest("/cust/outdoor/inc/getsubbrandcode.asp",  null, _getsubbrandcode, "GET");
		_sendRequest("/cust/outdoor/inc/getthemecode.asp", null, _getthemecode, "GET");
		document.getElementById("txtsubname").value = "";
		document.getElementById("txtthmname").value = "";
		document.getElementById("hdnsubno").value= "";
		document.getElementById("hdnthmno").value = "";
	}

	function _getbrandcode() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var displayseqno = document.getElementById("displayseqno");
				if (displayseqno) {
					displayseqno.innerHTML = xmlreq.responseText ;
					document.getElementById("cmbseqno").attachEvent("onchange", getsubbrandcode);
				}
			}
		}
	}

	function getsubbrandcode() {
		// �귣�带 ���� ������ ����
		var seqno = document.getElementById("cmbseqno").value;
		var subno = "" ;
		var params = "seqno="+seqno+"&subno="+subno ;
		_sendRequest("/cust/outdoor/inc/getsubbrandcode.asp", params, _getsubbrandcode, "GET");
		_sendRequest("/cust/outdoor/inc/getthemecode.asp", null, _getthemecode, "GET");
		document.getElementById("txtsubname").value = "";
		document.getElementById("txtthmname").value = "";
		document.getElementById("hdnsubno").value= "";
		document.getElementById("hdnthmno").value = "";
	}

	function  _getsubbrandcode() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var displaysubno = document.getElementById("displaysubno");
				if (displaysubno) {
					displaysubno.innerHTML = xmlreq.responseText ;
					document.getElementById("cmbsubno").attachEvent("onchange", getthemecode);
					document.getElementById("cmbsubno").attachEvent("onchange", function () {
		document.getElementById("txtsubname").value = document.getElementById("cmbsubno").options[document.getElementById("cmbsubno").selectedIndex].text;
		document.getElementById("hdnsubno").value = document.getElementById("cmbsubno").value;});
				}
			}
		}
	}

	function getthemecode() {
		//sub �귣�带 ���� ������ ����
		var subno = document.getElementById("cmbsubno").value;
		var thmno = "" ;
		var params = "subno="+subno+"&thmno="+thmno ;
		sendRequest("/cust/outdoor/inc/getthemecode.asp", params, _getthemecode, "GET");
		document.getElementById("txtthmname").value = "";
		document.getElementById("hdnthmno").value = "";
	}

	function _getthemecode() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var displaythmno = document.getElementById("displaythmno");
				if (displaythmno) {
					displaythmno.innerHTML = xmlreq.responseText ;
					document.getElementById("cmbthmno").attachEvent("onchange", function () {
		document.getElementById("txtthmname").value = document.getElementById("cmbthmno").options[document.getElementById("cmbthmno").selectedIndex].text;
		document.getElementById("hdnthmno").value = document.getElementById("cmbthmno").value;});
				}
			}
		}
	}

	function setsubno(crud) {
		if (crud == 'd') {
			if (isdelete()) {
				alert('�ش� ���簡 ���� �ϹǷ� �����Ͻ� �� �����ϴ�.\n\n�ش� ���縦 ���� �����Ͻʽÿ�');
				return false;
			} else {
				if (!confirm("������ ����귣�带 �����Ͻðڽ��ϱ�?")) {return false;}
			}
		}
		var seqno = document.getElementById("cmbseqno").value;
		if (!seqno) {alert("�귣�带 �����ϼ���"); return false;}
		var subno = document.getElementById("cmbsubno").value;
		if (crud == 'c') {
			if (!document.getElementById("txtsubname").value) { alert("����� ���� �귣�带 �Է��ϼ���"); document.getElementById('txtsubname').focus(); return false;}
			var i = document.getElementById("cmbsubno").length;
			if (i>0) 	subno = document.getElementById("cmbsubno").options[i-1].value;
		}
		var subname = document.getElementById("txtsubname").value;
		var params = "crud="+crud+"&subno="+subno+"&subname="+encodeURIComponent(subname)+"&seqno="+seqno ;
//		document.getElementById("debugConsole").innerHTML = params;
		sendRequest("/cust/outdoor/inc/setsubno.asp", params, getsubbrandcode, "GET");
		document.getElementById("txtsubname").value = "";
	}


	function setthmno(crud) {
		if (crud == 'd') {
			if (!confirm("���縦 �����Ͻðڽ��ϱ�?")) {return false;}
		}
		var subno = document.getElementById("cmbsubno").value;
		if (!subno) {alert("���� �귣�带 �����ϼ���"); return false;}
		var thmno = document.getElementById("cmbthmno").value;
		if (crud == 'c') {
			var i = document.getElementById("cmbthmno").length;
			if (i>0) 	thmno = document.getElementById("cmbthmno").options[i-1].value;
		}
		var thmname = document.getElementById("txtthmname").value;
		var params = "crud="+crud+"&subno="+subno+"&thmname="+encodeURIComponent(thmname)+"&thmno="+thmno ;
		sendRequest("/cust/outdoor/inc/setthmno.asp", params, getthemecode, "GET");
		document.getElementById("txtthmname").value = "";
	}


	function isdelete() {
		return document.getElementById("cmbthmno").options.length
	}

	function _debug() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var debugConsole = document.getElementById("debugConsole");
					debugConsole.innerHTML = xmlreq.responseText ;
			}
		}
	}

	window.onload = function () {
		document.getElementById("cmbcustcode").attachEvent("onchange", getbrandcode);
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
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="boldtitle"> ������� </span></TD>
				<TD width="50%" align="right"><span class="navigator" > ���ܱ��� &gt; ��ü���� &gt; �������</span></TD>
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
		  <td height="30" valign='middle'><!-- <img src='/images/m_add.gif' width='14' height='15' alt="�߰�" align='absmiddle'> �߰�  <img src='/images/m_edit.gif' width='16' height='15' alt="����"> ���� <img src='/images/m_delete.gif' width='15' height='15' alt="����"> ���� --> </td>
          </tr>
          <tr>
            <td >
			<table border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td width='270' height='30'><img src='/images/m_arw.gif' width='5' height='8' hspace='2'> ������ </td>
					<td width='230'><img src='/images/m_arw.gif' width='5' height='8' hspace='2'> ��ǥ �귣�� </td>
					<td width='230'><img src='/images/m_arw.gif' width='5' height='8' hspace='2'> ���� �귣�� </td>
					<td width='300'><img src='/images/m_arw.gif' width='5' height='8' hspace='2'> �귣�� ���� </td>
				</tr>
				<tr>
					<td><select size="20" id="cmbcustcode" name="cmbcustcode" style="width:265px;"><%Call getcustcode%></select></td>
					<td><div id="displayseqno"><select size='20' id='cmbseqno' name='cmbseqno' style='width:225px;'></select></div></td>
					<td><div id="displaysubno"><select size='20' id='cmbsubno' name='cmbsubno' style='width:225px;'></select></div></td>
					<td><div id="displaythmno"><select size='20' id='cmbthmno' name='cmbthmno' style='width:295px;'></select></div></td>
				</tr>
				<tr>
					<td height='40'>&nbsp;</td>
					<td>&nbsp;</td>
					<td align='bottom'> <input type="text" id="txtsubname"  name="txtsubname" style='width:225px;'/> <br>���� �귣�� [ <a href="#" onclick="setsubno('c'); return false;">�߰�</a> | <a href="#" onclick="setsubno('u'); return false;">����</a> | <a href="#" onclick="setsubno('d'); return false;">����</a><input type="hidden" id="hdnsubno" name="hdnsubno" /> ]</td>
					<td align='bottom'> <input type="text" id="txtthmname" name="txtthmname" style='width:290px;'/> <br>�귣�� ����� [ <a href="#" onclick="setthmno('c'); return false;"> �߰�</a>  | <a href="#" onclick="setthmno('u'); return false;">����</a> | <a href="#" onclick="setthmno('d'); return false;">����<input type="hidden" id="hdnthmno" name="hdnthmno" /></a> ]</td>
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
<script language="JavaScript">
<!--
	function go_medium_view(idx) {
		var url = "pop_job_view.asp?jobidx=" + idx + "&gotopage=<%=gotopage%>&searchstring=<%=searchstring%>";
		var name = "pop_job_view";
		var opt = "width=540, height=268, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function pop_job_reg() {
		var url = "pop_job_reg.asp";
		var name = "pop_job_reg";
		var opt = "width=540, height=268, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open(url, name, opt);

	}

	function go_page(url) {
		var frm = document.forms[0];
		location.href="job_list.asp?selcustcode="+frm.selcustcode.options[frm.selcustcode.selectedIndex].value;
	}

	function get_search() {
		var frm = document.forms[0];
		frm.action = "job_list.asp";
		frm.method = "post";
		frm.submit();
	}
//-->
</script>