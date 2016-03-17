<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim contidx : contidx = request("contidx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	dim objrs , sql, title
	sql = "select title from dbo.wb_contact_mst where contidx=" &contidx
	call get_recordset(objrs, sql)

	title = objrs(0).value

	objrs.close

	sql = "select m.contidx, m2.sidx,
	sql = "select m.contidx, m.sidx, m.title, m.qty, m.standard, m.quality, m.custcode, m.side, d.monthprice, d.expense, c.custname, d.perform, d.closing from dbo.wb_contact_md m inner join dbo.wb_contact_md_dtl d on m.contidx = d.contidx and m.sidx = d.sidx inner join dbo.sc_cust_temp c on m.custcode = c.custcode where m.contidx = "&contidx&" and d.cyear ='"&cyear&"' and d.cmonth='"&cmonth&"' "

	call get_recordset(objrs, sql)

	dim sidx, mdtitle, qty, standard, quality, custname3, monthprice, expense, side, perform, closing

	if not objrs.eof then
		set sidx = objrs("sidx")
		set mdtitle = objrs("title")
		set qty = objrs("qty")
		set side = objrs("side")
		set standard = objrs("standard")
		set quality = objrs("quality")
		set custname3 = objrs("custname")
		set monthprice = objrs("monthprice")
		set expense = objrs("expense")
		set perform = objrs("perform")		'	1 : 정산확인	0 : 미정산
		set closing = objrs("closing")			'	1 : RMS 마감	0 : RMS 미마감
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

<body bgcolor="#857C7A"  oncontextmenu="return false">
<form>
<input type="hidden" name="contidx" value="<%=contidx%>">
<input type="hidden" name="cyear" value="<%=cyear%>">
<input type="hidden" name="cmonth" value="<%=cmonth%>">
<table width="1000" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> <%=cyear%>.<%=cmonth%> &nbsp;<%=title%> 광고비 현황 </td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="1000" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="22"><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td bgcolor="#FFFFFF">
<!--  -->
	<table width="950" border="0" cellspacing="0" cellpadding="0" class="header" >
		<tr>
			<td class="hdbd" width="250">매체명</td>
                    <td class="hdbd" width="50">수량</td>
                    <td class="hdbd" width="50">면</td>
                    <td class="hdbd"  >규격/재질</td>
                    <td class="hdbd" width="90">월광고액</td>
                    <td class="hdbd" width="70">월지급액</td>
                    <td class="hdbd" width="70">내수액</td>
                    <td class="hdbd" width="50">내수율</td>
                    <td class="hdbd" width="130">매체사</td>
		</tr>
	     <%
			dim p_title, p_custname3, i, total_perform, total_closing
			i = 0
			do until objrs.eof
		%>
		<tr>
                    <td class="tbd" ><%if p_title <> mdtitle.value then response.write mdtitle.value%></td>
                    <td class="tbd"><%=qty.value%></td>
                    <td class="tbd"><%=side.value%></td>
                    <td class="tbd"><%=standard.value%> <%if not isnull(quality.value) then response.write "/" & quality.value %></td>
                    <td class="tbd"><% if not (perform and closing) then %><input type="text" name="txtmonthprice<%=i%>" id="txtmonthprice<%=i%>" style="width:90px;text-align:right" value="<%=formatnumber(monthprice.value,0)%>" onkeyup="comma(document.getElementById('txtmonthprice<%=i%>'));"><%else%> <%=formatnumber(monthprice.value,0)%><%end if%></td>
                    <td class="tbd"><% if not (perform and closing) then %><input type="text" name="txtexpense<%=i%>" id="txtexpense<%=i%>" value="<%=formatnumber(expense.value,0)%>" style="width:70px;text-align:right" onkeyup="comma(document.getElementById('txtexpense<%=i%>'));" onblur="calculation_income(document.getElementById('txtmonthprice<%=i%>'),document.getElementById('txtexpense<%=i%>'),document.getElementById('txtincome<%=i%>'),document.getElementById('txtincomeratio<%=i%>'));"><%else%><%=formatnumber(expense.value,0)%><%end if%></td>
                    <td class="tbd"><% if not (perform and closing) then %><input type="text" name="txtincome<%=i%>" id="txtincome<%=i%>" value="<%=formatnumber(monthprice.value-expense.value,0)%>" style="width:70px;text-align:right" readonly><%else%><%=formatnumber(monthprice.value-expense.value,0)%><%end if%></td>
                    <td class="tbd"><% if not (perform and closing) then %><input type="text" name="txtincomeratio<%=i%>" id="txtincomeratio<%=i%>"value="<%if monthprice <> 0 then response.write formatnumber(((monthprice.value-expense.value)/monthprice.value*100),2) else response.write "0.00"%>" style="width:50px;text-align:right;" readonly ><%else%><%if monthprice <> 0 then response.write formatnumber(((monthprice.value-expense.value)/monthprice.value*100),2) else response.write "0.00"%><%end if%> </td>
                    <td class="tbd"><%if p_custname3 <> custname3.value then response.write custname3.value%></td>
					<input type="hidden" name="sidx<%=i%>" value="<%=sidx%>">
		</tr>
		<%
				i = i + 1
				p_title = mdtitle.value
				p_custname3 = custname3.value
				total_perform = perform
				total_closing = closing
				objrs.movenext
			loop
			objrs.close
			set objrs = nothing
		%>
            <tr>
              <td  height="50" align="left" valign="bottom"><img src="/images/space.gif" width="59" height="18" border="0"></td>
              <td  align="right" valign="bottom" colspan="8"><% if not (total_perform and total_closing) then %><img src="/images/btn_save.gif" width="59" height="18"  vspace="5" style="cursor:hand" onClick="check_submit();"><img src="/images/btn_init.gif" width="67" height="18" vspace="5" style="cursor:hand" onClick="set_reset();" hspace="10"><% end if%><img src="/images/btn_close.gif" width="57" height="18" vspace="5" style="cursor:hand" onClick="set_close();" ></td>
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
<input type="hidden" name="txtcount" value="<%=i%>">
</form>
</body>
</html>
<script language="JavaScript">
<!--
	function set_close() {
		this.close();
	}

	function set_reset() {
		document.forms[0].reset();
	}

	function check_submit() {
		var frm = document.forms[0];
		frm.action = "execution_edit_proc.asp";
		frm.method = "post";
		frm.submit();
	}

	function comma(n){
			n.value =  Number(String(n.value).replace(/[^\d]/g,"")).toLocaleString().toLocaleString().slice(0,-3);
	}

	function calculation_income(pmonthprice, pexpense, pincome, pincomeratio) {
		var frm = document.forms[0];
		var monthprice = pmonthprice.value.replace(/[^\d]/g, "") ;
		var expense = pexpense.value.replace(/[^\d]/g, "") ;
		var income = monthprice-expense ;
		var incomeratio
		if (income==0) {
			pincomeratio.value = "0.00";
		} 	else {
			incomeratio = income / monthprice * 100 ;
			pincomeratio.value = Number(String(incomeratio)).toLocaleString();
		}
		pincome.value = Number(String(income).replace(/[^\d]/g,"")).toLocaleString().slice(0,-3);
	}

	window.onload = function () {
		self.focus();
	}
//-->
</script>