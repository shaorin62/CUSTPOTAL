<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	Dim searchstring : searchstring = request("seltotalcustcode")
	Dim sql : sql = "select custcode as DEPTCODE, custname as DEPTNAME, highcustcode as CUSTCODE,companyname as CUSTNAME from sc_cust_temp where medflag= 'A' and attr10 = 1 and highcustcode like '" & searchstring & "%' order by highcustcode"

	Dim objrs : Set objrs = server.CreateObject("adodb.recordset")
	objrs.activeConnection = application("connectionstring")
	objrs.cursorlocation = aduseclient
	objrs.cursortype = adopenstatic
	objrs.locktype = adlockreadonly
	objrs.source = sql
	objrs.open
	Dim custcode, custname, deptcode, deptname
	If Not objrs.eof Then
		Set deptcode = objrs("DEPTCODE")
		Set deptname = objrs("DEPTNAME")
		Set custcode = objrs("CUSTCODE")
		Set custname = objrs("CUSTNAME")
	End if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<title>▒▒ SK M&C | Media Management System ▒▒  </title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>

<body background="/images/pop_bg.gif"  oncontextmenu="return false">
<form>

<table width="522" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="22"><img src="/images/pop_left_top_bg.gif" width="22" height="102" ></td>
    <td background="/images/pop_center_top.gif" style="padding-top:12px;color:#FFFFFF; font-size:16px;font-weight:bolder;" width="379"> <img src="/images/pop_title_dot.gif" width="5" height="14" align="top" > 사업부 / 광고주 검색 <p> <%call get_custcode_total(custcode, null, null)%>  <img src="/images/btn_search.gif" width="39" height="20" align="absmiddle" onClick="getSerch();" class="styleLink" ></td>
    <td width="121"><img src="/images/pop_right_top_bg.gif" width="121" height="102"></td>
  </tr>
</table>
<table width="522" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="22"><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td>
<!--  -->
<TABLE width="100%"  bgcolor="#ECECEC"  border="0" cellpadding="0" cellspacing="1">
  <TR bgcolor="#ECECEC">
	<TD class="thd" >광고주명</TD>
	<TD class="thd" >사업부명</TD>
  </TR>
  <% Do Until objrs.eof %>
  <TR class="stylelink" onclick="checkForDept('<%=deptcode%>','<%=deptname%>', '<%=custcode%>','<%=custname%>');" bgcolor="#FFFFFF">
	<TD style="padding-left:10px;"><%=custname%></TD>
	<TD height="29" style="padding-left:10px;"><%=deptname%></TD>
  </TR>
  <%
		objrs.movenext
	Loop
	objrs.close
	Set objrs = nothing
  %>
  </TABLE>
<!--  -->
	</td>
    <td background="/images/pop_right_middle_bg.gif">&nbsp;</td>
  </tr>
  <tr>
    <td><img src="/images/pop_left_bottom_bg.gif" width="22" height="25"></td>
    <td background="/images/pop_center_bottom_bg.gif">&nbsp;</td>
    <td><img src="/images/pop_right_bottom_bg.gif" width="23" height="25"></td>
  </tr>
</table>
</form>
</body>
</html>

<SCRIPT LANGUAGE="JavaScript">
<!--
	window.onload = function init() {
		self.focus();
	}

	function getSerch() {
		var frm = document.forms[0];
		frm.action = "pop_custcode.asp";
		frm.method = "post";
		frm.submit();
	}

	function checkForDept(dcode, dname, ccode, cname) {
		var frm = window.opener.document.forms[0];
		frm.txtdeptcode.value = dcode;
		frm.txtdeptname.value = dname ;
		frm.txtcustcode.value = ccode;
		frm.txtcustname.value = cname;
		this.close();
	}
//-->
</SCRIPT>
