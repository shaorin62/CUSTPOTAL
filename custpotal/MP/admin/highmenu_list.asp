<!--#include virtual="/inc/getdbcon.asp" -->

<%
	Dim deptcode : deptcode = request.querystring("deptcode")
	Dim sql : sql = "select menuidx, menuname from dbo.web_board_menu where custcode = '" & deptcode & "'"

'	response.write searchstring
'	response.end

	Dim objrs : Set objrs = server.CreateObject("adodb.recordset")
	objrs.activeConnection = application("connectionstring")
	objrs.cursorlocation = aduseclient
	objrs.cursortype = adopenstatic
	objrs.locktype = adlockreadonly
	objrs.source = sql
	objrs.open
	Dim menuidx, menuname
	If Not objrs.eof Then
		set menuidx = objrs("menuidx")
		set menuname = objrs("menuname")
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
	<TD height="30">메뉴명</TD>
  </TR>
  <% Do Until objrs.eof %>
  <TR class="stylelink" onclick="checkForMenu('<%=menuidx%>','<%=menuname%>');">
	<TD height="22"><%=menuname%></TD>
  </TR>
  <%
		objrs.movenext
	Loop
	objrs.close
	Set objrs = nothing
  %>
  <TR onclick="checkForMenu('','');">
	<TD height="30">메뉴취소</TD>
  </TR>
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

	function checkForMenu(code, name) {
		var frm = window.opener.document.forms[0];
		frm.txthighmenuidx.value = code;
		frm.txthighmenuname.value = name ;
		this.close();
	}
//-->
</SCRIPT>

