<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
	Dim selcustcode : selcustcode = request.Form("selcustcode")

	Dim strUserid : strUserid = request("strUserid")
	Dim strCustcode : strCustcode = request("strCustcode")
	

	dim objrs, sql, Timcode, Timname
	sql = "select custcode Timcode, custname Timname from dbo.sc_cust_dtl where highcustcode <> custcode and highcustcode ='" & strcustcode &"'  and  isnull(gbnflag,'0')='0' and medflag='A'  and custcode not in ( select timcode from dbo.wb_account_tim where userid = '"& strUserid &"' group by timcode) "
	sql = sql & " order by custname,custcode"


	call get_recordset(objrs, sql)
	Dim custcode, custname
	If Not objrs.eof Then
		Set Timcode = objrs("Timcode")
		Set Timname = objrs("Timname")
	End if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
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
<INPUT TYPE="hidden" NAME="strUserid" value=<%=strUserid%>> 
<INPUT TYPE="hidden" NAME="strCustcode" value=<%=strCustcode%>> 
  <tr>
    <td width="22"><img src="/images/pop_left_top_bg.gif" width="22" height="102" ></td>
    <td background="/images/pop_center_top.gif" style="padding-top:12px;color:#FFFFFF; font-size:16px;font-weight:bolder;" > <img src="/images/pop_title_dot.gif" width="5" height="14" align="top" > 팀 저장 <p> <span style="font-size:11px;color:#333333;"> <A onclick='submitchange();  return false;' style="cursor:hand">[선택저장] </a></span> </td>
    <td width="121"><img src="/images/pop_right_top_bg.gif" width="121" height="102" ></td>
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
			<div id='#contents' style='margin-top:0px;width:460px; height: 310px; overflow-y:scroll'>
			<TABLE width="100%"  bgcolor="#ECECEC"  border="0" cellpadding="0" cellspacing="1">
			<TR bgcolor="#ECECEC">
				 <td class="thd"  width = "30">
						<table width="458" border="1" cellspacing="0" cellpadding="0">
							<tr>
								<td class="thd" width = "35" ><INPUT TYPE="checkbox" NAME="toggle" id='toggle' onclick='gettoggle();'></td>
								<TD class="thd" width = "440" >팀명</TD>
							</tr>
						</table>
					</td>
			  </TR>
			  <% Do Until objrs.eof %>
			  <TR class="stylelink"  bgcolor="#FFFFFF">
				<TD style="padding-left:10px;" height="31"><INPUT TYPE="checkbox" NAME="timidx"  value="<%=Timcode%>">  &nbsp; 
				<b>|</b>&nbsp; <%=Timcode%>&nbsp; <b>|</b> &nbsp;<%=Timname%></TD>
			  </TR>
			  <%
					objrs.movenext
				Loop
				objrs.close
				Set objrs = nothing
			  %>
			  </TABLE>
			   </div>
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
		frm.action = "pop_timcode.asp";
		frm.method = "post";
		frm.submit();
	}

//	function check_deptcode(ccode, cname) {
//
//		var frm = window.opener.document.forms[0];
//		frm.txtcustcode.value = ccode;
//		frm.txtcustname.value = cname;
//		this.close();
//	}

	function gettoggle() {
			var bln = document.getElementById("toggle").checked;
			var checkElement = document.getElementsByTagName("input");
			for (var i=0; i<checkElement.length;i++) {
				if (checkElement[i].getAttribute("type") == "checkbox") checkElement[i].checked = bln;
			}
		}
	

	function submitchange() {
		var cnt
		var bln = document.getElementById("toggle").checked;
		var checkElement = document.getElementsByTagName("input");

		cnt=0;
		for (var i=0; i<checkElement.length;i++) {
			if (checkElement[i].checked) {
				cnt = 1;
			}
		}	

		if (cnt == 0)
		{
			alert("선택된 내역이 없습니다.");
			return false;
		}

		var frm = document.forms[0];
		frm.action = "pop_timcode_proc.asp";
		frm.method = "post";
		frm.submit();
	}


//-->
</SCRIPT>
