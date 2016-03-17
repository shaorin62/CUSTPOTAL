<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	Dim searchstring : searchstring = request.Form("selcustcode")
	Dim sql : sql = "select C2.CUSTCODE as DEPTCODE, C2.CUSTNAME as DEPTNAME, C.CUSTCODE, C.CUSTNAME from dbo.SC_CUST_TEMP C INNER JOIN dbo.SC_CUST_TEMP C2 on C.CUSTCODE = C2.HIGHCUSTCODE where C.MEDFLAG = 'B' AND C.CUSTCODE LIKE '" & searchstring & "%'"

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
  <TITLE> SKM&C 부서관리 </TITLE>
<link href="/style.css" rel="stylesheet" type="text/css">
 </HEAD>

 <BODY  oncontextmenu="return false">

<FORM>
<TABLE border="1">
  <TR>
	<TD colspan="4" > 광고주검색 : <%call get_custcode_mst()%>  <img src="/images/btn_search.gif" width="39" height="20" align="absmiddle" onClick="getSerch();" class="styleLink" > </TD>
  </TR>
  <TR bgcolor="#ECECEC">
	<TD height="31">사업부명</TD>
	<TD>광고주명</TD>
  </TR>
  <% Do Until objrs.eof %>
  <TR class="stylelink" onclick="checkForDept('<%=deptcode%>','<%=deptname%>', '<%=custcode%>','<%=custname%>');">
	<TD><%=deptname%></TD>
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
