<!--#include virtual="/inc/getdbcon.asp" -->

<%
	Dim searchstring : searchstring = request.Form("txtsearchstring")
	Dim sql : sql = "select CUSTCODE, CUSTNAME from dbo.SC_CUST_TEMP where MEDFLAG = 'B' AND CUSTCODE = HIGHCUSTCODE AND custname like '" & searchstring & "%' ORDER BY CUSTNAME "

	Dim objrs : Set objrs = server.CreateObject("adodb.recordset")
	objrs.activeConnection = application("connectionstring")
	objrs.cursorlocation = aduseclient
	objrs.cursortype = adopenstatic
	objrs.locktype = adlockreadonly
	objrs.source = sql
	objrs.open
	Dim custcode, custname
	If Not objrs.eof Then
		Set custcode = objrs("CUSTCODE")
		Set custname = objrs("CUSTNAME")
	End if
%>
<HTML>
 <HEAD>
<title>�Ƣ� SK M&C | Media Management System �Ƣ� </title>
<link href="../style.css" rel="stylesheet" type="text/css">
 </HEAD>

 <BODY  oncontextmenu="return false">

<FORM METHOD=POST ACTION="">
<TABLE border="1">
  <TR>
	<TD colspan="2"> �˻��� ��ü����� �Է��ϼ��� : <INPUT TYPE="text" NAME="txtsearchstring" size="15"> <img src="/images/btn_search.gif" width="44" height="22" align="absmiddle" onClick="getSerch();" class="styleLink" > </TD>
  </TR>
  <TR>
	<TD>��ü���ڵ�</TD>
	<TD>��ü���</TD>
  </TR>
  <% Do Until objrs.eof %>
  <TR class="stylelink" onclick="checkForDept('<%=custcode%>','<%=custname%>');">
	<TD height="22"><%=custcode%></TD>
	<TD><%=custname%></TD>
  </TR>
  <%
		objrs.movenext
	Loop
	objrs.close
	Set objrs = nothing
  %>
  </TABLE>
</FORM>

 </BODY>
</HTML>
<SCRIPT LANGUAGE="JavaScript">
<!--
	window.onload = function init() {
		self.focus();
	}

	function getSerch() {
		var frm = document.forms[0];
		frm.action = "employee_list.asp";
		frm.method = "post";
		frm.submit();
	}

	function checkForDept(code, name) {
		var frm = window.opener.document.forms[0];
		frm.txtdeptcode.value = "";
		frm.txtdeptname.value = "" ;
		frm.txtcustcode.value = code;
		frm.txtcustname.value = name;
		this.close();
	}
//-->
</SCRIPT>
