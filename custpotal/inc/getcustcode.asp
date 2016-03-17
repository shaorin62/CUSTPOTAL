<!--#include virtual = "/inc/getdbcon.asp" -->
<%

	Dim strdeptname : strdeptname = request.Form("txtdeptname")
	
	Dim objrs : Set objrs = server.CreateObject("adodb.recordset")
	objrs.activeconnection = application("connectionstring")
	objrs.cursorlocation = aduseclient
	objrs.cursortype = adopenforwardonly
	objrs.locktype = adlockreadonly
	objrs.source = "SELECT C.CUSTCODE, C.CUSTNAME, C2.CUSTCODE, C2.CUSTNAME FROM dbo.SC_CUST_TEMP C INNER JOIN dbo.SC_CUST_TEMP C2 ON C.CUSTCODE = C2.HIGHCUSTCODE WHERE C.MEDFLAG = 'A' AND C2.CUSTNAME LIKE '%" & strdeptname & "%' AND 1 = 1"
	objrs.open 

	Dim custcode, custname, deptcode, deptname
	If Not objrs.eof Then 
		Set custcode = objrs(0)
		Set custname = objrs(1)
		Set deptcode = objrs(2)
		Set deptname = objrs(3)
	End if
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
 <HEAD>
  <TITLE> 옥외 광고주 선택 </TITLE>
  <META NAME="Generator" CONTENT="EditPlus">
  <META NAME="Author" CONTENT="">
  <META NAME="Keywords" CONTENT="">
  <META NAME="Description" CONTENT="">
 </HEAD>
<link type="text/css" rel="stylesheet" href="/style.css">
 <BODY>  
<FORM METHOD=POST ACTION="">
검색할 사업부명을 입력 하세요 : <INPUT TYPE="text" NAME="txtdeptname"> <IMG SRC="/images/go.gif" WIDTH="34" HEIGHT="16" BORDER="0" ALT="" class="stylelink" onclick="getdeptcode()">
			<TABLE border=1>
	<TR>
		<TD>광고주코드</TD>
		<TD>광고주</TD>
		<TD>사업부코드</TD>
		<TD>사업부</TD>
		<TD>&nbsp;</TD>
	</TR>
	<% Do Until objrs.eof %>
	<TR>
		<TD><%=custcode%></TD>
		<TD><%=custname%></TD>
		<TD><%=deptcode%></TD>
		<TD><%=deptname%></TD>
		<TD><span class="stylelink" onclick="putdeptcode('<%=custcode%>','<%=custname%>','<%=deptcode%>','<%=deptname%>')">선택</span></TD>
	</TR>
	<%
		objrs.movenext
		Loop
		
		objrs.close
		Set objrs = Nothing
	%>
	</TABLE>
	</FORM>
 </BODY>
</HTML>
<script language="javascript">
	function putdeptcode(ccode, cname, dcode, dname) {
		var frm = window.opener.document.forms[0];
		frm.txtdeptname.value = dname;
		frm.txtcustname.value = cname;
		frm.txtcustcode.value = ccode;
		frm.txtdeptcode.value = dcode;

		frm.txtmdname.focus();
		this.close();
	}

	function getdeptcode() {
		var frm = document.forms[0];
		frm.method = "POST";
		frm.action = "getcustcode.asp";
		frm.submit();
	}
</script>