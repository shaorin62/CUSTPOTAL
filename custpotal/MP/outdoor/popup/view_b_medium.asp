<!--#include virtual="/MP/outdoor/inc/Function.asp" -->
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
	<link href="/MP/outdoor/style.css" rel="stylesheet" type="text/css">
	<script type='text/javascript' src='/js/ajax.js'></script>
	<script type='text/javascript' src='/js/script.js'></script>
	<script type="text/javascript">
	<!--
		function validElement() {
			var frm = document.forms[0];
			var chk = true;
			for (var i = 0; i< frm.rdoside.length ; i++) {
				if (frm.rdoside[i].checked) chk = false;
			}
			if (chk) {alert('면의 위치를 선택하세요'); return false;}

			if (frm.txtstandard.value.replace(/\s/g, "") == "") {
				alert("매체 규격을 입력하세요");
				frm.txtstandard.focus();
				return false;
			}

			for (var i = 0 ; i<frm.rdoside.length;i++) {
				frm.rdoside[i].disabled = false;
				if (frm.rdoside[i].checked) document.getElementById("txtside").value = frm.rdoside[i].value;
			}
			submitchange();
		}
		function submitchange() {
			var frm = document.forms[0];
			frm.action = "/MP/outdoor/process/db_b_medium.asp";
			frm.method = "post";
			frm.submit();
		}

		function calculation() {
			// 내수액(율) 자동 계산, 월광고료 선입력 필수 체크
			var monthly = parseFloat(document.getElementById("txtmonthly").value.replace(/,/g, ""));
			var expense = parseFloat(document.getElementById("txtexpense").value.replace(/,/g, ""));

			var income = (monthly - expense).toLocaleString().slice(0, -3) ;
			if (monthly > 0 ) var rate =  ((monthly - expense)/monthly*100).toLocaleString();
			else rate = "0.00" ;

			document.getElementById("incomeview").innerHTML = income + " ("+rate+")";
		}
		function setclose() {
			document.getElementById("themeLayer").style.display='none';
		}



		window.onload = function () {
			self.focus();
			var crud = "<%=pcrud%>";
			if (crud == "d") {
				submitchange();
			}
			document.getElementById("themeLayer").style.display='none';
			document.getElementById("txttheme").attachEvent("onfocus", function() {document.getElementById("themeLayer").style.display='block';});
			calculation();
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
<input type='hidden' id='subexeseq' name='subexeseq' value='<%=subexeseq%>' />
<input type='hidden' id='no' name='no' value='<%=no%>' />
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="absmiddle"> <%=title%> : <%if contidx ="" then %> 면 정보 등록 <% else %> <%=getside(pside)%> 면 <% end if %></td>
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
				<td class="hdr h" width='95'>면위치 </td>
				<td class="sc" width='400'>
				<%
					sql_ = "select distinct side from wb_contact_exe where mdidx =?"
					cmd_.commandText = sql_
					cmd_.parameters.append cmd_.createparameter("mdidx", adinteger, adparaminput, ,mdidx)
					Set rs_ = cmd_.execute
					Do Until rs_.eof
						If Trim(rs_(0)) = "R" Then disabled1 = " disabled"
						If Trim(rs_(0)) = "L" Then disabled2 = " disabled"
						If Trim(rs_(0)) = "F" Then disabled3 = " disabled"
						If Trim(rs_(0)) = "B" Then disabled4 = " disabled"
						rs_.movenext
					Loop
					Set cmd_ = nothing
				%>
				<input type='radio' id='rdoside' name='rdoside' value='L' <%=disabled2%> <%If pside="L" Then response.write " checked "%>> 좌측 <input type='radio' id='rdoside' name='rdoside' value='R' <%=disabled1%> <%If pside="R" Then response.write " checked "%>> 우측  <input type='radio' id='rdoside' name='rdoside' value='F' <%=disabled3%> <%If pside="F" Then response.write " checked "%>> 정면 <input type='radio' id='rdoside' name='rdoside' value='B' <%=disabled4%> <%If pside="B" Then response.write " checked "%>> 후면 <input type="hidden" id="txtside" name='txtside' />
				</td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">수량 </td>
				<td class="sc"><input name="txtqty" type="text" id="txtqty" maxlength="5"  class="number" onclick='comma(this);' onkeyup="comma(this);" value="<%=qty%>"> (<%=unit%>) </td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h">규격 </td>
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
			<tr>
				<td class="hdr h"> 월광고료</td>
				<td class="sc"> <input name="txtmonthly" type="text" id="txtmonthly" maxlength="20"  class="currency"  onclick='comma(this);'  onkeyup="comma(this); calculation();" value="<%=FormatNumber(monthly,0)%>" ></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h"> 월지급액</td>
				<td class="sc"> <input name="txtexpense" type="text" id="txtexpense" maxlength="20"  class="currency"  value="<%=FormatNumber(expense,0)%>" onkeyup="comma(this); calculation();" onclick='comma(this);' ></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h"> 내수액(율) </td>
				<td class="sc"> <div id="incomeview">0 (0.00)</div></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="hdr h" width='95'> 집행소재 </td>
				<td class="sc" width='400'> <input name="txttheme" type="text" id="txttheme" style="width:370px" style="ime-mode:active;" value="<%=thmname%>"><input type='hidden' name='hdnthmno' id='hdnthmno' value="<%=thmno%>"> </td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			</table>
			<table border="0" cellpadding="0" cellspacing="0">
			<tr>
                  <td  align="right" valign="bottom" width='495' height='50'><% if contidx <> "" then %> * 변경된 내역은 선택된 년월 이후 모든 내역에 반영됩니다. <% end if %><br><a href="#" onclick="validElement(); return false;"><img src="/images/btn_save.gif" width="59" height="18"  vspace="5" border=0></a><a href="#" onclick="reset(); return false;"><img src="/images/btn_init.gif" width="67" height="18" vspace="5" border=0 hspace="10" ></a><a href="#" onclick="window.close(); return false;"><img src="/images/btn_close.gif" width="57" height="18" vspace="6" border=0 ></a></td>
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
<div id="themeLayer"style="LEFT:10px; TOP:67px; width:520px; height:297px;position:absolute; z-index:10; background-color:#CCCCCC;border: 1px solid #333333;color:#000000;padding:10 10 10 10;font-weight:bolder;display:none;overflow:hidden;"><!--#include virtual="/MP/outdoor/inc/getsubseq.asp" --></div>


