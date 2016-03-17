<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim idx : idx = request("idx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")

	dim objrs, sql
	sql = "select m.contidx, title from dbo.wb_contact_mst m inner join dbo.wb_contact_md m2 on m.contidx = m2.contidx inner join dbo.wb_contact_md_dtl d on m2.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on d.idx = a.idx where a.idx="& idx
	call get_recordset(objrs, sql)

	dim contidx : contidx = objrs("contidx")
	dim title : title = objrs("title")

	objrs.close

	sql = "select cyear, cmonth, monthprice, expense, isPerform from dbo.wb_contact_md_dtl_account where idx = " & idx
	call get_recordset(objrs, sql)
	objrs.sort = "cyear, cmonth"

	dim cyear2, cmonth2, monthprice, expense, income, incomeratio, isPerform, totalMonthPrice, totalExpense

	if not objrs.eof then
		set cyear2 = objrs("cyear")
		set cmonth2 = objrs("cmonth")
		set monthprice = objrs("monthprice")
		set expense = objrs("expense")
		set isPerform = objrs("isPerform")
	end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<title>▒▒ SK M&C | Media Management System ▒▒  </title>
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
<form >
<INPUT TYPE="hidden" NAME="idx" value="<%=idx%>">
<table width="522" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> <%=title%></td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="522" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="22"><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr bgcolor="#FFFFFF">
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td>
	<!--  -->
	  <table border="0" cellpadding="0" cellspacing="0" align="center" >
          <tr>
            <td class="thd2" width="100" align="center">년.월</td>
            <td  class="thd2" width="100" align="center">월광고료</td>
            <td  class="thd2" width="100" align="center">월지급액</td>
            <td  class="thd2" width="100" align="center">내수액</td>
            <td  class="thd2" width="100" align="center">내수율</td>
          <tr>
          </tr>
			<td colspan="5" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
		  <%
			dim intLoop
			intLoop = 0
			do until objrs.eof
			income = monthprice - expense

			if income = 0 or monthprice = 0 then incomeratio = "0.00" else incomeratio = income/monthprice*100

		  %>
          <tr>
            <td class="tbd" ><%=cyear2%>.<%=cmonth2%>&nbsp;</td>
            <td  class="tbd" ><input type="text" name="txtmonthprice" value="<%=formatnumber(monthprice,0)%>" size="12" style="text-align:right;" onkeyup="comma(<%=intLoop%>);" id="txtmonthprice" onblur="checkZero();" <%if isPerform then response.write "readonly"%>></td>
            <td  class="tbd" ><input type="text" name="txtexpense" value="<%=formatnumber(expense,0)%>" size="12" style="text-align:right;" onkeyup="comma2(<%=intLoop%>);" id="txtexpense" onblur="checkZero();" <%if isPerform then response.write "readonly"%>></td>
            <td  class="tbd" width="100" align="right"><span id="income" ><%=formatnumber(income,0)%></span>&nbsp;</td>
            <td  class="tbd" width="100" align="right"><span id="incomeratio" ><%=formatnumber(incomeratio,2)%></span>&nbsp;</td>
          <tr>
		  <tr>
			<td colspan="5" bgcolor="#E7E7DE" height="1"><input type="hidden" name="txtcyear" value="<%=cyear%>" ><input type="hidden" name="txtcmonth" value="<%=cmonth%>"></td>
		  </tr>
		  <%
			totalMonthPrice = totalMonthPrice + monthprice
			totalExpense = totalExpense + expense
			income = 0
			incomeratio = 0
			intLoop = intLoop + 1
			objrs.movenext
			loop
			objrs.close
			set objrs = nothing
		  %>
		  <tr>
			<td colspan="5" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td  class="tbd" > 총광고비 </td>
            <td  class="tbd" ><span id="totalMonthPrice"><%=formatnumber(totalMonthPrice,0)%></span></td>
            <td  class="tbd" ><span id="totalExpense"><%=formatnumber(totalExpense,0)%></span></td>
            <td  class="tbd" ><span id="totalIncome"><%=formatnumber(totalMonthPrice - totalExpense,0)%></span></td>
            <td  class="tbd" ><span id="totalIncomeRatio"><% if totalMonthPrice <> 0 then response.write  formatnumber((totalMonthPrice - totalExpense)/totalMonthPrice*100,2) else response.write "0.00"%></span></td>
          <tr>
		  <tr>
			<td colspan="5" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
		  <tr>
				<td colspan="5"  height="50" valign="bottom" align="right"> <img src="/images/btn_save.gif" width="59" height="18"  vspace="5" style="cursor:hand" onClick="check_submit();" ><img src="/images/btn_init.gif" width="67" height="18" vspace="5" style="cursor:hand" onClick="set_reset();"  hspace="10"> <img src="/images/btn_close.gif" width="57" height="18" vspace="5" style="cursor:hand" onClick="set_close();" >
	</td>
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
	<script language="JavaScript">
	<!--
		function check_submit() {
			var frm = document.forms[0];
			frm.action = "pop_contact_account_edit_proc.asp";
			frm.method = "post";
			frm.submit();

		}

		function chk_account() {
			var frm = document.forms[0];
			for (var i = 0 ; i <frm.chkperform.length ; i++) {
				if (frm.chkperform[i].checked) frm.txtchk[i].value = "1"
				else frm.txtchk[i].value="0";
			}
		}

		function set_reset() {
			document.location.reload();
		}

		function set_close() {
			this.close();
		}

		window.onload=function () {
			self.focus();
		}

		function comma(p){
			var ctl =0 ;
			var input = document.getElementsByTagName("input");
			for (var i = 0 ; i < input.length ; i++) {
				if (!input[i].getAttribute("id")) continue;
				if (input[i].getAttribute("id") == "txtmonthprice") {
					if (ctl == p) input[i].value =  Number(String(input[i].value).replace(/[^\d]/g,"")).toLocaleString().toLocaleString().slice(0,-3);
					ctl++;
				}
			}
		}
		function comma2(p){
			var ctl =0 ;
			var input = document.getElementsByTagName("input");
			for (var i = 0 ; i < input.length ; i++) {
				if (!input[i].getAttribute("id")) continue;
				if (input[i].getAttribute("id") == "txtexpense") {
					if (ctl == p) input[i].value =  Number(String(input[i].value).replace(/[^\d]/g,"")).toLocaleString().toLocaleString().slice(0,-3);
					ctl++;
				}
			}
		}

		function checkZero() {
			var frm = document.forms[0];
			var input = document.getElementsByTagName("input");

			var monthprice = new Array();
			var expense = new Array()
			var income = new Array();
			var incomeratio = new Array();

			var m = n = 0;

			var tmp_income, tmp_incomeratio ;

			for (var i = 0 ; i < input.length ; i++) {
				if (!input[i].getAttribute("id")) continue;

				if (input[i].getAttribute("id") == "txtmonthprice")  {
					monthprice[m] = parseInt(input[i].value.replace(/,/g, ""));
					m++;
				}

				if (input[i].getAttribute("id") == "txtexpense") {
					expense[n] = parseInt(input[i].value.replace(/,/g, ""));
					n++;
				}

			}

			m = n = 0 ;

			var span = document.getElementsByTagName("span") ;
			for (var i = 0 ; i < span.length ; i++) {
				if (!span[i].getAttribute("id")) continue;

				if (span[i].getAttribute("id") == "income") {
					income[m] = span[i];
					m++;
				}

				if (span[i].getAttribute("id") == "incomeratio") {
					incomeratio[n] = span[i];
					n++;
				}
			}

			for (var i = 0 ; i <monthprice.length; i++) {
				tmp_income = monthprice[i] - expense[i] ;
				income[i].innerText = Number(String(tmp_income).replace(/[^\d-]/g,"")).toLocaleString().toLocaleString().slice(0,-3);
				tmp_incomeratio = tmp_income / monthprice[i] * 100 ;
				if (tmp_income == 0 || monthprice[i] == 0 ) incomeratio[i].innerText = "0.00";
				else incomeratio[i].innerText = tmp_incomeratio.toFixed(2);
			}

			var totalmonthprice = totalexpense = totalincome = totalincomeratio = 0 ;
			for (var i = 0 ; i < frm.txtmonthprice.length ; i++) {
				totalmonthprice = totalmonthprice + parseInt(frm.txtmonthprice[i].value.replace(/,/g,""));
				totalexpense = totalexpense + parseInt(frm.txtexpense[i].value.replace(/,/g,""));
			}
			totalincome = totalmonthprice - totalexpense;
			if (totalincome == 0) totalincomeratio =0;
			else totalincomeratio = totalincome / totalmonthprice * 100 ;
			document.getElementById("totalMonthPrice").innerText = Number(String(totalmonthprice).replace(/[^\d-]/g,"")).toLocaleString().toLocaleString().slice(0,-3);
			document.getElementById("totalExpense").innerText = Number(String(totalexpense).replace(/[^\d-]/g,"")).toLocaleString().toLocaleString().slice(0,-3);
			document.getElementById("totalIncome").innerText = Number(String(totalincome).replace(/[^\d-]/g,"")).toLocaleString().toLocaleString().slice(0,-3);
			if (totalmonthprice == 0) document.getElementById("totalIncomeRatio").innerText = "0.00";
			else  document.getElementById("totalIncomeRatio").innerText = totalincomeratio.toFixed(2);

		}
	//-->
	</script>