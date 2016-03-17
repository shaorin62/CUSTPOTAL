<!--#include virtual="/inc/getdbcon.asp" -->

<%
	Dim searchstring : searchstring = request.Form("txtsearchstring")
	Dim sql : sql = "select deptcode, deptname,  custcode, custname from dbo.VW_CUST where MEDFLAG = 'A' AND CUSTCODE <> 'A00000' AND deptname like '%" & searchstring & "%' "
	'"select deptcode, deptname,  custcode, custname from dbo.VW_CUST where custcode like '%" & searchcode & "%' AND deptname LIKE '%" & searchstring & "%' "

'	response.write searchstring
'	response.end

	Dim objrs : Set objrs = server.CreateObject("adodb.recordset")
	objrs.activeConnection = application("connectionstring")
	objrs.cursorlocation = aduseclient
	objrs.cursortype = adopenstatic
	objrs.locktype = adlockreadonly
	objrs.source = sql
	objrs.open
	Dim deptcode, deptname, custcode, custname
	If Not objrs.eof Then
		set deptcode = objrs("deptcode")
		set deptname = objrs("deptname")
		Set custcode = objrs("custcode")
		Set custname = objrs("custname")
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
	<TD colspan="2"> 광고주 검색 <input type="text" name="txtsearchstring"> <img src="/images/btn_search.gif" width="44" height="22" align="absmiddle" vspace="5" onClick="getSerch();" class="styleLink" > </TD>
  </TR>
  <TR>
	<TD height="30">사업부명</TD>
	<TD>광고주명</TD>
  </TR>
  <% Do Until objrs.eof %>
  <TR class="stylelink" onclick="checkForDept('<%=deptcode%>','<%=deptname%>');">
	<TD height="22"><%=deptname%></TD>
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
		frm.action = "customer_list2.asp";
		frm.method = "post";
		frm.submit();
	}

	function checkForDept(code, name) {
		var frm = window.opener.document.forms[0];
		frm.txtdeptcode.value = code;
		frm.txtdeptname.value = name ;
		this.close();
	}
//-->
</SCRIPT>

