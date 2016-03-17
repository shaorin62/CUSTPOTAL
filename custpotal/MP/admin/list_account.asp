<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<%
		Dim sql : sql = "select distinct a.highcustcode, a.custname from sc_cust_hdr a inner join sc_cust_dtl b on a.highcustcode=b.highcustcode where a.medflag='B' and a.use_flag=1 and b.med_out = 1 order by a.custname"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		Set cmd = Nothing

		Sub getmedcode
			Do Until rs.eof
				response.write "<option value='" & rs(0) & "'>" & rs(1) & "</option>"
				rs.movenext
			Loop
		End Sub
%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/hq/outdoor/style.css" rel="stylesheet" type="text/css">
<script type='text/javascript' src='/js/ajax.js'></script>
<script type='text/javascript' src='/js/script.js'></script>
<script type="text/javascript">
<!--

	// 매체사 직원 정보
	function getemployee(crud) {
		var medcode = document.getElementById("cmbmed").value ;
		var empid = document.getElementById("cmbemp").value;
		var name = "winEmployee";
		var left = screen.width / 2 - 550 / 2;
		var top = 10;
		var opt = "width=540, height=328, resizable=no, scrollbars=no, status=yes, left="+left+"&top="+top;
		switch (crud) {
			case "c":
		var url = "/hq/outdoor/popup/view_emp.asp?medcode="+medcode+"&crud="+crud ;
				if (medcode =="") {alert("등록할 매체사를 선택하세요"); return false;}
				window.open(url, name, opt);
				break;
			case "u":
		var url = "/hq/outdoor/popup/view_emp.asp?medcode="+medcode+"&empid="+empid+"&crud="+crud ;
				if (empid == '') {alert("수정할 직원을 선택하세요"); return false;}
				window.open(url, name, opt);
				break;
			case "d":
		var url = "/hq/outdoor/popup/view_emp.asp?medcode="+medcode+"&empid="+empid+"&crud="+crud ;
				if (empid == '') {alert("삭제할 직원을 선택하세요"); return false;}
				if (confirm("선택한 직원을 삭제하시겠습니까?")) {
					window.open(url, name, opt);
				}
				break ;
		}
	}

	function getmedemployee() {
		var medcode = document.getElementById("cmbmed").value ;
		var params = "medcode="+medcode ;
		sendRequest("/hq/outdoor/inc/getmedemployee.asp", params, _getmedemployee, "GET");

	}

	function _getmedemployee() {
		if (xmlreq.readyState == 4) {
			if (xmlreq.status == 200) {
				var employeeview = document.getElementById("employeeview");
				employeeview.innerHTML = xmlreq.responseText ;
				var cmbemp = document.getElementById("cmbemp");
				cmbemp.style.width = "200px";
				cmbemp.style.height = "260px";
				cmbemp.setAttribute("size", "20");
			}
		}
	}

	function msg(str) {
		alert(str);
		return false;
	}


	window.onload = function () {
		document.getElementById("cmbmed").attachEvent("onchange", getmedemployee);
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
<!--#include virtual="/hq/top.asp" -->
  <table id="Table_01" width="1240"  border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/hq/left_outdoor_menu.asp"--></td>
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
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="boldtitle"> 계정관리 </span></TD>
				<TD width="50%" align="right"><span class="navigator" > 옥외광고 &gt; 매체관리 &gt; 계정관리</span></TD>
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
		  <td height="30" valign='middle'><!-- <img src='/images/m_add.gif' width='14' height='15' alt="추가" align='absmiddle'> 추가  <img src='/images/m_edit.gif' width='16' height='15' alt="수정"> 수정 <img src='/images/m_delete.gif' width='15' height='15' alt="삭제"> 삭제  --></td>
          </tr>
          <tr>
            <td >
			<table border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td width='270' height='30'><img src='/images/m_arw.gif' width='5' height='8' hspace='2'> 매체사 </td>
					<td width='205'><img src='/images/m_arw.gif' width='5' height='8' hspace='2'> 직원계정 </td>
					<td width='555'><!-- <img src='/images/m_arw.gif' width='5' height='8' hspace='2'> 계정관리  --></td>
				</tr>
				<tr>
					<td><select size="20" id="cmbmed" name="cmbmed" style="width:265px;"><%Call getmedcode%></select></td>
					<td align='right'><div id="employeeview"><select size='20' id='cmbemp' name='cmbemp' style='width:200;height:260px;'></select></div> <br>
					 <a href="#" onclick="getemployee('c'); return false;"><!-- <img src='/images/m_add.gif' width='14' height='15' alt="추가" align='absmiddle' > --> 추가 </a> |  <a href="#" onclick="getemployee('u'); return false;"><!-- <img src='/images/m_edit.gif' width='16' height='15' alt="수정" hspace='2'> --> 수정 </a> | <a href="#" onclick="getemployee('d'); return false;"><!-- <img src='/images/m_delete.gif' width='15' height='15' alt="삭제" > --> 삭제 </a>
					 </td>
					<td valign='top'> </td>
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