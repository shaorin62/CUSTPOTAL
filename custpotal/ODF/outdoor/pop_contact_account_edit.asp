<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim contidx : contidx = request("contidx")
	dim sidx : sidx = request("sidx")

	dim objrs, sql
	sql = "select title, cyear, cast(cmonth as int) as cmonth, monthprice, expense, perform from dbo.wb_contact_md m inner join dbo.wb_contact_md_dtl d on m.sidx = d.sidx where d.contidx = " & contidx &" and d.sidx = " & sidx
	call get_recordset(objrs, sql)
	objrs.sort = "cyear, cmonth"

	dim title, cyear, cmonth, monthprice, expense, income, incomeratio, perform, summonthprice, sumexpense

	if not objrs.eof then
		title = objrs("title")
		set cyear = objrs("cyear")
		set cmonth = objrs("cmonth")
		set monthprice = objrs("monthprice")
		set expense = objrs("expense")
		set perform = objrs("perform")
	end if
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
<form >
<table width="522" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> <%=title%> 광고비 집행 현황 </td>
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
	  <input type="hidden" name="contidx" value="<%=contidx%>">
	  <input type="hidden" name="sidx" value="<%=sidx%>">
	  <table border="0" cellpadding="0" cellspacing="0" align="center" >
          <tr>
            <td  class="thd2" width="30" align="center">&nbsp;</td>
            <td class="thd2" width="70" align="center">년.월</td>
            <td  class="thd2" width="100" align="center">월광고료</td>
            <td  class="thd2" width="100" align="center">월지급액</td>
            <td  class="thd2" width="70" align="center">내수액</td>
            <td  class="thd2" width="70" align="center">내수율</td>
          <tr>
          </tr>
			<td colspan="6" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
		  <%
			do until objrs.eof
			income = monthprice - expense
			if income = 0 then incomeratio = "0.00" else incomeratio = income/monthprice*100
		  %>
          <tr>
            <td  class="tbd" ><input type="checkbox" name="chkperform" <%if perform then response.write " checked " %> onclick="chk_account();" disabled></td>
            <td class="tbd" ><%=cyear%>.<%if len(cmonth) = 1 then response.write "0"&cmonth else response.write cmonth%>&nbsp;</td>
            <td  class="tbd" ><%=formatnumber(monthprice,0)%>&nbsp;</td>
            <td  class="tbd" ><%=formatnumber(expense,0)%>&nbsp;</td>
            <td  class="tbd" ><%=formatnumber(income,0)%>&nbsp;</td>
            <td  class="tbd" ><%=formatnumber(incomeratio,2)%>&nbsp;</td>
          <tr>
		  <tr>
			<td colspan="6" bgcolor="#E7E7DE" height="1"><input type="hidden" name="txtcyear" value="<%=cyear%>"><input type="hidden" name="txtcmonth" value="<%=cmonth%>"><input type="hidden" name="txtchk" value="<%if perform then response.write "1" else response.write "0"%>"></td>
		  </tr>
		  <%
			summonthprice = summonthprice + monthprice
			sumexpense = sumexpense + expense
			income = 0
			incomeratio = 0
			objrs.movenext
			loop
			objrs.close
			set objrs = nothing
		  %>
		  <tr>
			<td colspan="6" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td  class="tbd" colspan="2"> 총광고비 </td>
            <td  class="tbd" ><%=formatnumber(summonthprice,0)%>&nbsp;</td>
            <td  class="tbd" ><%=formatnumber(sumexpense,0)%>&nbsp;</td>
            <td  class="tbd" ><%=formatnumber(summonthprice-sumexpense,0)%>&nbsp;</td>
            <td  class="tbd" ><% if summonthprice <> 0 then response.write  formatnumber((summonthprice-sumexpense)/summonthprice*100,2) else response.write "0.00"%>&nbsp;</td>
          <tr>
		  <tr>
			<td colspan="6" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
		  <tr>
				<td colspan="6"  height="50" valign="bottom" align="right"> <!-- <img src="/images/btn_save.gif" width="59" height="18"  vspace="5" style="cursor:hand" onClick="check_submit();" ><img src="/images/btn_init.gif" width="67" height="18" vspace="5" style="cursor:hand" onClick="set_reset();"  hspace="10"> --><img src="/images/btn_close.gif" width="57" height="18" vspace="5" style="cursor:hand" onClick="set_close();" >
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
			document.forms[0].reset();
		}

		function set_close() {
			this.close();
		}

		window.onload=function () {
			self.focus();
		}
	//-->
	</script>