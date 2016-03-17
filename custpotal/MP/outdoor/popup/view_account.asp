<!--#include virtual="/MP/outdoor/inc/Function.asp" -->
<%
	On Error Resume Next
'	Call getquerystringparameter
	Dim pcontidx : pcontidx = request("contidx")
	Dim pmdidx : pmdidx = request("mdidx")
	Dim pside : pside = request("side")
	Dim pflag : pflag = request("flag")
	If pside = "" Then pside = "F"

	Dim sql :  sql = "select a.seq, a.cyear, a.cmonth, a.monthly, a.expense, b.ishold from wb_contact_exe a left outer join (select a.cyear, a.cmonth, a.contidx, a.isHold, b.mdidx from wb_contact_trans a left outer join wb_contact_md b on a.contidx=b.contidx) as b on a.mdidx=b.mdidx and a.cyear=b.cyear and a.cmonth=b.cmonth where a.mdidx=? and a.side=?"
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	cmd.parameters.append cmd.createparameter("mdidx", adInteger, adParaminput)
	cmd.parameters.append cmd.createparameter("side", adChar, adParamInput, 1)
	cmd.parameters("mdidx").value = pmdidx
	cmd.parameters("side").value = pside
	Dim rs_ : Set rs_ = cmd.execute
	Set cmd = Nothing
	Dim rs_seq
	Dim rs_cyear
	Dim rs_cmonth
	Dim rs_monthly
	Dim rs_expense
	Dim rs_ishold
	If Not rs_.eof Then
		Set rs_seq = rs_(0)
		Set rs_cyear = rs_(1)
		Set rs_cmonth = rs_(2)
		Set rs_monthly = rs_(3)
		Set rs_expense = rs_(4)
		Set rs_ishold = rs_(5)
	End If
%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<link href="/MP/outdoor/style.css" rel="stylesheet" type="text/css">
	<title>▒▒ SK M&C | Media Management System ▒▒  </title>
	<script type='text/javascript' src='/js/ajax.js'></script>
	<script type='text/javascript' src='/js/script.js'></script>
    <script type="text/javascript" src="/js/calendar.js"></script>
	<script type="text/javascript">
	<!--

		function get() {
			var params = "";
			sendRequest(url, params, _get, "GET");
		}

		function _get() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
				}
			}
		}

		function calculate() {
			// 내수액(율) 자동 계산
			var monthly = document.forms[0].monthly;
			var expense = document.forms[0].expense;
			var income = document.forms[0].income;
			var rate = document.forms[0].rate;

			var totmonthly = totexpense = totincome = 0;

			var elemCount = monthly.length;
			for (var i = 0 ; i < elemCount ;i++) {

				mon = parseInt(monthly[i].value.replace(/,/g,""));
				exp = parseInt(expense[i].value.replace(/,/g, ""));
				icm = (mon-exp);
				income[i].value = icm.toLocaleString().slice(0,-3);

				if (icm == 0)
					rate[i].value = "0.00";
				else
					if (parseFloat(icm /mon * 100).toFixed(2) > 0 ) rate[i].value = parseFloat(icm /mon * 100).toFixed(2);
					else rate[i].value = "0.00";

				totmonthly += mon;
				totexpense += exp;
			}
			totincome = parseFloat(totmonthly-totexpense);
			if ((totincome == 0) || (totincome < 0) )	totrate= "0.00";
			else	totrate = parseFloat(totincome/totmonthly * 100).toFixed(2) ;

			document.getElementById("monthlyview").innerText = totmonthly.toLocaleString().slice(0,-3);
			document.getElementById("expenseview").innerText = totexpense.toLocaleString().slice(0,-3);
			document.getElementById("incomeview").innerText = totincome.toLocaleString().slice(0,-3);
			document.getElementById("rateview").innerText = totrate;
		}

		function submitchange() {
			var frm = document.forms[0];
			frm.action = "/MP/outdoor/process/db_account.asp";
			frm.method = "post";
			frm.submit();
		}

		window.onload = function () {
			self.focus();
			calculate();
		}

		window.onunload = function () {
//			try {
//				window.opener.getcontactdetail();
//			} catch(e) {
//				window.close();
//			}
		}
	//-->
	</script>
</head>

<body>
<form onsubmit="return submitchange();">
<input type="hidden" id="contidx" name="contidx" value="<%=pcontidx%>" />
<input type="hidden" id="mdidx" name="mdidx" value="<%=pmdidx%>" />
<input type="hidden" id="side" name="side" value="<%=pside%>" />
<input type="hidden" id="flag" name="flag" value="<%=pflag%>" />
<table width="550" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="absmiddle"> 광고 비용 관리 </td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="550" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="22"><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td height="458" valign='top'>
		<div style='overflow-y:scroll;width:500px;height:428px;'>
<!--  -->
		<table border="0" cellpadding="0" cellspacing="0">
			<tr height='20'>
				<td width='0' valign='top'>&nbsp;</td>
				<td width='0' align='right' valign='top'>&nbsp;</td>
			</tr>
		</table>
		<table border="0" cellpadding="0" cellspacing="0">
			<tr>
				<th class='normal' width="20"><img src='/images/lock.gif' width='16' height='16' alt='결산집행여부'></th>
				<th class='normal' width="60">연도</th>
				<th class='normal' width="60">월</th>
				<th class='normal' width="110">월광고료</th>
				<th class='normal' width="110">월지급액</th>
				<th class='normal' width="90">내수액</th>
				<th class='normal' width="50">내수율</th>
			</tr>
			<%
				Dim income_
				Dim rate_
				Do Until rs_.eof
				income_ = rs_monthly - rs_expense
				If rs_monthly = 0 Then rate_ = 0 Else rate_ = income_/rs_monthly * 100
			%>
			<tr>
				<td class='normal'><%If rs_ishold = "Y"  Then response.write "<img src='/images/hold.gif' width='16' height='16' alt='정산완료' hspace=2>" Else If rs_ishold= "N" Then  response.write "<img src='/images/lock.gif' width='16' height='16' alt='정산요청'>" Else response.write "<img src='/images/unlock.gif' width='16' height='16' alt='미정산'>" End If %></td>
				<td class='normal'><%=rs_cyear%><input type="hidden" id="cyear" name="cyear" value="<%=rs_cyear%>"/></td>
				<td class='normal'><%=rs_cmonth%><input type="hidden" id="cmonth" name="cmonth" value="<%=rs_cmonth%>"/></td>
				<td class='normal'><input type="text" id="monthly" name='monthly' value="<%=FormatNumber(rs_monthly,0)%>" style="width:107px;" maxlength= '15' <%If Not IsNull(rs_ishold) Then response.write " readonly style='border:0px; padding-right:3px;'"%> class="currency" onclick='comma(this);' onkeyup='comma(this); calculate();'/></td>
				<td class='normal'><input type="text" id="expense" name='expense' value="<%=FormatNumber(rs_expense,0)%>" style="width:107"  maxlength= '15'  <%If Not IsNull(rs_ishold) Then response.write " readonly style='border:0px;padding-right:3px'"%> class="currency" onclick='comma(this);' onkeyup='comma(this); calculate();'/></td>
				<td class='normal' style='text-align:right; padding-right:3px;'><input type="text" id='income' name='income' value='<%=FormatNumber(income_,0)%>' readonly style='border:0px; padding-right:3px; text-align:right; width=87'/></td>
				<td class='normal' style='text-align:right; padding-right:3px; '><input type="text" id='rate' name='rate' value='<%=FormatNumber(rate_,2)%>' readonly style='border:0px; padding-right:3px;text-align:right;width=47px;'/></td>
			</tr>
			<%
					rs_.movenext
				Loop
			%>
			<tr>
				<th class='normal' colspan='3'>총 계</th>
				<th class='normal' width="110"><div id='monthlyview'></div></th>
				<th class='normal' width="110"><div id='expenseview'></div></th>
				<th class='normal' width="90"><div id='incomeview'></div></th>
				<th class='normal' width="50"><div id='rateview'></div></th>
			</tr>
		</table>
		</div>
<!--  -->
	</td>
    <td background="/images/pop_right_middle_bg.gif"></td>
  </tr>
  <tr>
    <td><img src="/images/pop_left_bottom_bg.gif" width="22" height="25"></td>
    <td background="/images/pop_center_bottom_bg.gif">&nbsp;</td>
    <td><img src="/images/pop_right_bottom_bg.gif" width="23" height="25"></td>
  </tr>
</table>

<div id='buttonLayer' style='left:408px; top:505px;width:120px;height:18px;position:absolute; z-index:9;' ><input type='image'  src='/images/btn_save.gif' width='57' height='18'> <a href="#" onclick='window.close(); return false;'><img src='/images/btn_close.gif' width='57' height='18'></a> </div>
</form>
</body>
</html>
<%
	If Err.number <> 0 Then
		Call Debug
	End If
%>