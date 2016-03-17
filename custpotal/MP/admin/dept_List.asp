<!--#include virtual="/inc/getdbcon.asp" -->

<%
	Dim searchstring : searchstring = request.Form("txtsearchstring")
	Dim sql : sql = "select C2.CUSTCODE as DEPTCODE, C2.CUSTNAME as DEPTNAME, C.CUSTCODE, C.CUSTNAME from dbo.SC_CUST_TEMP C INNER JOIN dbo.SC_CUST_TEMP C2 on C.CUSTCODE = C2.HIGHCUSTCODE where C.MEDFLAG = 'A' AND C.CUSTNAME like '" & searchstring & "%' "

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
<HTML>
 <HEAD>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<link href="../style.css" rel="stylesheet" type="text/css">
 </HEAD>

 <BODY  oncontextmenu="return false">

<FORM METHOD=POST ACTION="">
<TABLE border="1">
  <TR>
	<TD colspan="4"> 검색할 사업부명을 입력하세요 : <INPUT TYPE="text" NAME="txtsearchstring" size="15"> <img src="/images/btn_search.gif" width="44" height="22" align="absmiddle" onClick="getSerch();" class="styleLink" > </TD>
  </TR>
  <TR>
	<TD>사업부코드</TD>
	<TD>사업부명</TD>
	<TD>광고주코드</TD>
	<TD>광고주명</TD>
  </TR>
  <% Do Until objrs.eof %>
  <TR class="stylelink" onclick="checkForDept('<%=deptcode%>','<%=deptname%>', '<%=custcode%>','<%=custname%>');">
	<TD height="22"><%=deptcode%></TD>
	<TD><%=deptname%></TD>
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
