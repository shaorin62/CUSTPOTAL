<!--#include virtual="/inc/getdbcon.asp" -->

<%
	Dim deptname : deptname = request.Form("txtdeptname")
	Dim sql : sql = "SELECT E.EMPNO, E.EMP_NAME, E.CC_CODE, C.CC_NAME, E.E_MAIL FROM dbo.SC_EMPLOYEE_MST E LEFT OUTER JOIN dbo.SC_CC C ON E.CC_CODE = C.CC_CODE " &_
						" WHERE E.EMP_NAME LIKE '%"  & deptname & "%'  ORDER BY CC_NAME"

	Dim objrs : Set objrs = server.CreateObject("adodb.recordset")
	objrs.activeConnection = application("connectionstring")
	objrs.cursorlocation = aduseclient
	objrs.cursortype = adopenstatic
	objrs.locktype = adlockreadonly
	objrs.source = sql
	objrs.open
	Dim empno, empname, cc_code, cc_name, email
	If Not objrs.eof Then
		set empno = objrs("EMPNO")
		set empname = objrs("EMP_NAME")
		Set cc_code = objrs("CC_CODE")
		Set cc_name = objrs("CC_NAME")
		set email = objrs("E_MAIL")
	End if
%>
<HTML>
 <HEAD>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<link href="../style.css" rel="stylesheet" type="text/css">
 </HEAD>

 <BODY  oncontextmenu="return false">

<FORM METHOD=POST ACTION="">
<TABLE border="1">
  <TR>
	<TD colspan="5"> 사원명을 입력하세요 : <INPUT TYPE="text" NAME="txtdeptname"  size="15"> <img src="/images/btn_search.gif" width="44" height="22" align="absmiddle" onClick="getSerch();" class="styleLink" > </TD>
  </TR>
  <TR>
	<TD>사원코드</TD>
	<TD>사원명</TD>
	<TD>부서코드</TD>
	<TD>부서명</TD>
	<TD>Email</TD>
  </TR>
  <% Do Until objrs.eof %>
  <TR class="stylelink" onclick="checkForDept('<%=empno%>', '<%=empname%>','<%=cc_code%>','<%=cc_name%>','<%=email%>');">
	<TD height="22"><%=empno%></TD>
	<TD><%=empname%></TD>
	<TD height="22"><%=cc_code%></TD>
	<TD><%=cc_name%></TD>
	<TD><%=email%></TD>
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

	function checkForDept(empno, empname, code, name, email) {
		var frm = window.opener.document.forms[0];
		frm.txtempno.value = empno;
		frm.txtempname.value = empname ;
		frm.txtdeptcode.value = code;
		frm.txtdeptname.value = name ;
		frm.txtcustcode.value = "A00001";
		frm.txtcustname.value = "㈜SKM&C";
		this.close();
	}
//-->
</SCRIPT>
