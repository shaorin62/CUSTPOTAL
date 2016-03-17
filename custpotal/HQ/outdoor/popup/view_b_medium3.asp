<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<%
	Dim pcontidx : pcontidx = request("contidx")
	Dim pmdidx : pmdidx = request("mdidx")
	Dim pcrud : pcrud = request("crud")
	Dim pcyear : pcyear = request("cyear")
	Dim pcmonth : pcmonth = request("cmonth")
	Dim pside : pside = request("side")
	Dim cmd_ : Set cmd_ = server.CreateObject("adodb.command")
	Dim sql_ : sql_ = "select title, b.mdidx, c.highcustcode, unit , isnull(qty,1) qty , standard, quality, isnull(monthly,0)  monthly, isnull(expense,0) expense, f.thmno, thmname, f.seq, f.no from wb_contact_mst a  left outer join wb_contact_md b on a.contidx = b.contidx  left outer join sc_cust_dtl c on c.custcode = a.custcode left outer join wb_contact_md_dtl d on d.seq = (select max(seq) from wb_contact_md_dtl where mdidx=? and side=? and cyear+cmonth <= ?)  left outer join wb_contact_exe e on b.mdidx = e.mdidx and d.side = e.side and e.cyear =? and e.cmonth =  ?  left outer join wb_subseq_exe f on b.mdidx = f.mdidx and d.side = f.side and f.seq = (select max(seq) from wb_subseq_exe where cyear+cmonth <= ? and mdidx=?) left outer join wb_subseq_dtl g on f.thmno = g.thmno where a.contidx = ?"
	cmd_.activeconnection = application("connectionstring")
	cmd_.commandText = sql_
	cmd_.parameters.append cmd_.createparameter("mdidx", adInteger, adParamInput, , pmdidx)
	cmd_.parameters.append cmd_.createparameter("side", adChar, adParamInput, 1, pside)
	cmd_.parameters.append cmd_.createparameter("cyearmonth", adChar, adParamInput, 6, pcyear&pcmonth)
'	cmd_.parameters.append cmd_.createparameter("side2", adChar, adParamInput, 1, pside)
	cmd_.parameters.append cmd_.createparameter("cyear", adChar, adParamInput, 4, pcyear)
	cmd_.parameters.append cmd_.createparameter("cmonth", adChar, adParamInput, 2, pcmonth)
	cmd_.parameters.append cmd_.createparameter("cyearmonth2", adChar, adParamInput, 6, pcyear&pcmonth)
	cmd_.parameters.append cmd_.createparameter("mdidx2", adInteger, adParamInput, , pmdidx)
	cmd_.parameters.append cmd_.createparameter("contidx", adInteger, adParamInput, , pcontidx)
	cmd_.commandType = adCmdText


	Dim rs : Set rs = cmd_.execute
	Dim mdidx
	Dim title
	Dim custcode
	Dim unit
	Dim qty
	Dim standard
	Dim quality
	Dim monthly
	Dim expense
	Dim thmno
	Dim thmname
	Dim subexeseq
	dim no
	If Not rs.eof Then
		title = rs(0)
		mdidx = rs(1)
		highcustcode = rs(2)
		unit = rs(3)
		qty = rs(4)
		standard = rs(5)
		quality = rs(6)
		monthly = rs(7)
		expense = rs(8)
		thmno = rs(9)
		thmname = rs(10)
		subexeseq = rs(11)
		no = rs(12)
	End If

	clearParameter(cmd_)
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
		function validElement() {
			var frm = document.forms[0];

			if (frm.txtstandard.value.replace(/\s/g, "") == "") {
				alert("매체 규격을 입력하세요");
				frm.txtstandard.focus();
				return false;
			}

			var frm = document.forms[0];
			frm.action = "/hq/outdoor/process/db_b_medium3.asp";
			frm.method = "post";
			frm.submit();
		}

		window.onload = function () {
			self.focus();
		}
	//-->
	</script>
</head>
<body>
<form>
<input type="hidden" id="crud" name="crud" value="<%=pcrud%>" />
<input type="hidden" id="cyear" name="cyear" value="<%=pcyear%>" />
<input type="hidden" id="cmonth" name="cmonth" value="<%=pcmonth%>" />
<input type="hidden" id="side" name="side" value="<%=pside%>" />
<input type="hidden" id="mdidx" name="mdidx" value="<%=pmdidx%>" />
<input type='hidden' id='contidx' name='contidx' value='<%=pcontidx%>' />
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="absmiddle"> <%=title%> :  면 정보 수정 </td>
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
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h" style="width:80px;">규격 </td>
				<td class="sc"><input name="txtstandard" type="text" id="txtstandard"maxlength="100" style="width:370px" value="<%=standard%>"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">재질 </td>
				<td class="sc"><% Call getquality(quality)%></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
                  <td  align="right" valign="bottom" width='495' height='50'><% if contidx <> "" then %> * 변경된 내역은 선택된 년월 이후 모든 내역에 반영됩니다. <% end if %><br><a href="#" onclick="validElement(); return false;"><img src="/images/btn_save.gif" width="59" height="18"  vspace="5" border=0></a> <a href="#" onclick="window.close(); return false;"><img src="/images/btn_close.gif" width="57" height="18" vspace="5" border=0 ></a></td>
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


