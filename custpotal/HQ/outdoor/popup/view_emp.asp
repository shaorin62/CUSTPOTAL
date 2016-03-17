<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<!--#include virtual="/inc/secure.inc.asp"-->
<%
'	Dim item 
'	For Each item In request.querystring
'		response.write item & " : "& request.querystring(item) & "<br>"
'	Next
	
	Dim atag : atag = ""	
	Dim pmedcode : pmedcode = clearXSS(request("medcode"), atag)
	Dim pempid : pempid = clearXSS(request("empid"), atag)
	Dim pcrud : pcrud = request("crud")
	If pcrud = "C" Or pcrud = "D" Then pempid = "" 
	Dim sql : sql = "select empid, medcode, empname, emppwd, useflag, custname from wb_med_employee a inner join sc_cust_hdr b on a.medcode=b.highcustcode where a.empid=? and useflag = '1'"
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	cmd.parameters.append cmd.createparameter("empid", adChar, adParamInput, 9)
	cmd.parameters("empid").value = pempid
	Dim rs : Set rs = cmd.Execute
	Set cmd = Nothing 

	If Not rs.eof Then 
		Dim empid : empid = rs(0)
		Dim medcode : medcode = rs(1)
		Dim empname : empname = rs(2)
		Dim emppwd : emppwd = rs(3)
		Dim useflag : useflag = rs(4)
		Dim medname : medname = rs(5)
	End If 
	If medname = "" Then medname = getmedname(pmedcode)
	If empid = "" Then empid = getempid(pmedcode)
	
%>
<html>
<head>
	<title>▒▒ SK M&C | Media Management System ▒▒  </title>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<link href="/hq/outdoor/style.css" rel="stylesheet" type="text/css">
	<script type='text/javascript' src='/js/ajax.js'></script>
	<script type='text/javascript' src='/js/script.js'></script>
	<script type="text/javascript">
	<!--
		var crud = "<%=pcrud%>";
		function submitchange() {
			var frm = document.forms[0];
			if (crud != 'd') {
				if (frm.empname.value.replace(/\s/g, '') == "") {
					alert("직원명을 입력하세요");
					frm.empname.focus();
					return false;
				}
				if (frm.emppwd.value != frm.reemppwd.value) {
					alert("비밀번호가 서로 다릅니다.");
					frm.reemppwd.select();
					return false;
				}		
			}

			frm.action = "/hq/outdoor/process/db_emp.asp";
			frm.method = "post";
			frm.submit(); 
		}
		window.onload = function () {
		if (crud=='d') {submitchange(); }		
			self.focus();
			document.getElementById("coverLayer").style.display='none';
		}
	//-->
	</script>
</head>
<body>
<form>
<input type="hidden" id="medcode" name='medcode' value='<%=pmedcode%>' />
<input type="hidden" id="crud" name='crud' value='<%=pcrud%>' />
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="absmiddle"> 매체사 직원 정보 관리 </td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="22"><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td>
<!--  -->
			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td class="hdr h">매체사명 </td>
				<td class="sc"><input name="medname" type="text" id="medname"maxlength="100" style="width:350px" value="<%=medname%>" readonly></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">직원계정 </td>
				<td class="sc"><input name="empid" type="text" id="empid"maxlength="100" style="width:350px" value="<%=empid%>" readonly></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">직원이름 </td>
				<td class="sc"><input name="empname" type="text" id="empname"maxlength="15" style="width:350px" value="<%=empname%>"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">비밀번호 </td>
				<td class="sc"><input name="emppwd" type="text" id="emppwd"maxlength="15" style="width:350px" value="<%=emppwd%>"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">비밀번호확인 </td>
				<td class="sc"><input name="reemppwd" type="text" id="reemppwd"maxlength="15" style="width:350px" ></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h"> 접근여부</td>
				<td class="sc"> <input type="radio" id="agree" name='useflag' value='1' <%If useflag = "1" Or useflag = "" Then response.write "checked"%>> 허용 <input type="radio" id="deny" name='useflag' value='0' <%If useflag = "0" Then response.write "checked"%>> 거부 </td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
                  <td  align="right" valign="bottom" width='495' height='50'><a href="#" onclick="submitchange(); return false;"><img src="/images/btn_save.gif" width="59" height="18"  vspace="5" border=0 hspace='5'></a><a href="#" onclick="window.close(); return false;"><img src="/images/btn_close.gif" width="57" height="18" vspace="5" border=0 ></a></td>
                </tr>
			</table>
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
<div id="coverLayer" class='select-free'><div id='bd'> delete data ...</div></div>