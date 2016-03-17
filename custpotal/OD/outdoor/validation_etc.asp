<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim tidx : tidx = request("tidx")
	dim mdidx : mdidx = request("mdidx")
	dim objrs, sql
	sql = "select title from dbo.wb_medium_mst where mdidx = " & mdidx
	call get_recordset(objrs, sql)

	dim title : title = objrs(0).value
	objrs.close

	sql = "select v.tidx, v.code, v.value from dbo.wb_validation_value v inner join dbo.wb_validation_class  c on c.tidx = v.tidx where c.mdidx = " & mdidx & " and isuse = 1 "

	call get_recordset(objrs, sql)
	dim sel1_1, val1_1
	dim sel1_2, val1_2
	dim sel1_3, val1_3
	dim sel1_4, val1_4
	dim sel1_5, val1_5

	dim sel2_1, val2_1
	dim sel2_2, val2_2
	dim sel2_3, val2_3
	dim sel2_4, val2_4
	dim sel2_5, val2_5
	dim sel2_6, val2_6

	dim sel3_1, val3_1
	dim sel3_2, val3_2
	dim sel3_3, val3_3

	dim sel4_1, val4_1
	dim sel4_2, val4_2
	dim sel4_3, val4_3

	if not objrs.eof then
		do until objrs.eof
		if objrs("code") = "sel1_1" then
			sel1_1 = objrs("code").value
			val1_1 = objrs("value").value
		end if
		if objrs("code") = "sel1_2" then
			sel1_2 = objrs("code").value
			val1_2 = objrs("value").value
		end if
		if objrs("code") = "sel1_3" then
			sel1_3 = objrs("code").value
			val1_3 = objrs("value").value
		end if
		if objrs("code") = "sel1_4" then
			sel1_4 = objrs("code").value
			val1_4 = objrs("value").value
		end if
		if objrs("code") = "sel1_5" then
			sel1_5 = objrs("code").value
			val1_5 = objrs("value").value
		end if
'2
		if objrs("code") = "sel2_1" then
			sel2_1 = objrs("code").value
			val2_1 = objrs("value").value
		end if
		if objrs("code") = "sel2_2" then
			sel2_2 = objrs("code").value
			val2_2 = objrs("value").value
		end if
		if objrs("code") = "sel2_3" then
			sel2_3 = objrs("code").value
			val2_3 = objrs("value").value
		end if
		if objrs("code") = "sel2_4" then
			sel2_4 = objrs("code").value
			val2_4 = objrs("value").value
		end if
		if objrs("code") = "sel2_5" then
			sel2_5 = objrs("code").value
			val2_5 = objrs("value").value
		end if
		if objrs("code") = "sel2_6" then
			sel2_6 = objrs("code").value
			val2_6 = objrs("value").value
		end if
'3
		if objrs("code") = "sel3_1" then
			sel3_1 = objrs("code").value
			val3_1 = objrs("value").value
		end if
		if objrs("code") = "sel3_2" then
			sel3_2 = objrs("code").value
			val3_2 = objrs("value").value
		end if
		if objrs("code") = "sel3_3" then
			sel3_3 = objrs("code").value
			val3_3 = objrs("value").value
		end if
'4
		if objrs("code") = "sel4_1" then
			sel4_1 = objrs("code").value
			val4_1 = objrs("value").value
		end if
		if objrs("code") = "sel4_2" then
			sel4_2 = objrs("code").value
			val4_2 = objrs("value").value
		end if
		if objrs("code") = "sel4_3" then
			sel4_3 = objrs("code").value
			val4_3 = objrs("value").value
		end if
		objrs.movenext
		loop
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

<body  oncontextmenu="return false">
<form>
<INPUT TYPE="hidden" NAME="tidx" value="<%=tidx%>">
<INPUT TYPE="hidden" NAME="mdidx" value="<%=mdidx%>">
<table width="820" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top">  <%=title%> 평가 기준표 </td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="820" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="22"><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td>
<!--  -->
<table width="775" border="1" cellpadding="3" cellspacing="1" bordercolor="#CCCCCC">
  <tr>
    <td colspan="3" rowspan="2" align="center"><strong>평가항목</strong></td>
    <td rowspan="2" align="center"><strong>가중치</strong></td>
    <td height="25" colspan="4" align="center"><strong>평가기준</strong></td>
    <td rowspan="2" align="center"><strong>평가</strong></td>
    <td rowspan="2" align="center"><strong>환산점수<br>
      (가중치X평가)</strong></td>
  </tr>
  <tr>
    <td width="100" height="25" align="center"><strong>A(4)</strong></td>
    <td width="100" align="center"><strong>B(3)</strong></td>
    <td width="100" align="center"><strong>C(2)</strong></td>
    <td width="100" align="center"><strong>D(1)</strong></td>
  </tr>
  <tr>
    <td rowspan="6" align="center">1. 지역환경</td>
    <td rowspan="2" align="center">Coverage</td>
    <td align="center">설치지역</td>
    <td align="center"><span class="style2">5</span></td>
    <td height="30" align="center">전략지역</td>
    <td align="center">수도권/광역시</td>
    <td align="center">국지성</td>
    <td align="center">-<BR></td>
    <td align="center"><select name="sel1_1" id="sel1_1" onChange="convert_sum();">
      <option value="0"  <%if sel1_1 = "sel1_1" and val1_1 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel1_1 = "sel1_1" and val1_1 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel1_1 = "sel1_1" and val1_1 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel1_1 = "sel1_1" and val1_1 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel1_1 = "sel1_1" and val1_1 = "1" then response.write "selected"%>>D</option>
    </select>    </td>
    <td align="center"><span id="1_1">0</span></td>
  </tr>
  <tr>
    <td align="center">설치범위</td>
    <td align="center">5</td>
    <td height="30" align="center">복수집행(5site 이상) </td>
    <td align="center">복수집행(2-5site)</td>
    <td align="center">단수집행</td>
    <td align="center">-</td>
    <td align="center"><select name="sel1_2" id="sel1_2"  onChange="convert_sum();">
      <option value="0"  <%if sel1_2 = "sel1_2" and val1_2 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel1_2 = "sel1_2" and val1_2 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel1_2 = "sel1_2" and val1_2 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel1_2 = "sel1_2" and val1_2 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel1_2 = "sel1_2" and val1_2 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="1_2">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">설치위치</td>
    <td align="center"><span class="style2">5</span></td>
    <td align="center">강제노출</td>
    <td align="center">자연노출</td>
    <td align="center">의도가시</td>
    <td align="center">-</td>
    <td align="center"><select name="sel1_3" id="sel1_3"  onChange="convert_sum();">
      <option value="0" <%if sel1_3 = "sel1_3" and val1_3 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel1_3 = "sel1_3" and val1_3 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel1_3 = "sel1_3" and val1_3 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel1_3 = "sel1_3" and val1_3 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel1_3 = "sel1_3" and val1_3 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="1_3">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">유동인구(차량)</td>
    <td align="center"><span class="style2">10</span></td>
    <td align="center">고밀도</td>
    <td align="center">적정</td>
    <td align="center">저밀도</td>
    <td align="center">-</td>
    <td align="center"><select name="sel1_4" id="sel1_4"  onChange="convert_sum();">
      <option value="0" <%if sel1_4 = "sel1_4" and val1_4 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel1_4 = "sel1_4" and val1_4 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel1_4 = "sel1_4" and val1_4 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel1_4 = "sel1_4" and val1_4 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel1_4 = "sel1_4" and val1_4 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="1_4">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">상품적합성<br>
      (Relevant)</td>
    <td align="center"><span class="style2">5</span></td>
    <td align="center">고</td>
    <td align="center">중</td>
    <td align="center">저</td>
    <td align="center">-</td>
    <td align="center"><select name="sel1_5" id="sel1_5"  onChange="convert_sum();">
      <option value="0" <%if sel1_5 = "sel1_5" and val1_5 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel1_5 = "sel1_5" and val1_5 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel1_5 = "sel1_5" and val1_5 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel1_5 = "sel1_5" and val1_5 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel1_5 = "sel1_5" and val1_5 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center" ><span id="1_5">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center" bgcolor="#FFFF99"><strong>소 계</strong></td>
    <td align="center" bgcolor="#FFFF99"><span class="style2">30</span></td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99"><span id="sum_1">0</span></td>
  </tr>
  <tr>
    <td rowspan="7" align="center">2.매체사양</td>
    <td height="30" colspan="2" align="center">광고물 규격</td>
    <td align="center"><span class="style2">10</span></td>
    <td align="center">20㎡ 이상 </td>
    <td align="center">9㎡</td>
    <td align="center">5㎡이상</td>
    <td align="center">5㎡미만</td>
    <td align="center"><select name="sel2_1" id="sel2_1"  onChange="convert_sum();">
      <option value="0" <%if sel2_1 = "sel2_1" and val2_1 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel2_1 = "sel2_1" and val2_1 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel2_1 = "sel2_1" and val2_1 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel2_1 = "sel2_1" and val2_1 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel2_1 = "sel2_1" and val2_1 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="2_1">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">광고면비율</td>
    <td align="center">5</td>
    <td align="center">전면</td>
    <td align="center">3/4이상</td>
    <td align="center">1/2이상</td>
    <td align="center">-</td>
    <td align="center"><select name="sel2_2" id="sel2_2"  onChange="convert_sum();">
      <option value="0" <%if sel2_2 = "sel2_2" and val2_2 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel2_2 = "sel2_2" and val2_2 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel2_2 = "sel2_2" and val2_2 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel2_2 = "sel2_2" and val2_2 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel2_2 = "sel2_2" and val2_2 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="2_2">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">광고면소재</td>
    <td align="center"><span class="style2">5</span></td>
    <td align="center">특수소재</td>
    <td align="center">조명시트</td>
    <td align="center">단순시트</td>
    <td align="center">-</td>
    <td align="center"><select name="sel2_3" id="sel2_3"  onChange="convert_sum();">
      <option value="0" <%if sel2_3 = "sel2_3" and val2_3 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel2_3 = "sel2_3" and val2_3 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel2_3 = "sel2_3" and val2_3 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel2_3 = "sel2_3" and val2_3 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel2_3 = "sel2_3" and val2_3 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="2_3">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">품의(Prestige)</td>
    <td align="center"><span class="style2">10</span></td>
    <td align="center">고</td>
    <td align="center">중</td>
    <td align="center">저</td>
    <td align="center">-</td>
    <td align="center"><select name="sel2_4" id="sel2_4"  onChange="convert_sum();">
      <option value="0" <%if sel2_4 = "sel2_4" and val2_4 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel2_4 = "sel2_4" and val2_4 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel2_4 = "sel2_4" and val2_4 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel2_4 = "sel2_4" and val2_4 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel2_4 = "sel2_4" and val2_4 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="2_4">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">환경친화성</td>
    <td align="center"><span class="style2">5</span></td>
    <td align="center">고</td>
    <td align="center">중</td>
    <td align="center">저</td>
    <td align="center">-</td>
    <td align="center"><select name="sel2_5" id="sel2_5"  onChange="convert_sum();">
      <option value="0" <%if sel2_5 = "sel2_5" and val2_5 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel2_5 = "sel2_5" and val2_5 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel2_5 = "sel2_5" and val2_5 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel2_5 = "sel2_5" and val2_5 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel2_5 = "sel2_5" and val2_5 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="2_5">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">공익성</td>
    <td align="center">5</td>
    <td align="center">고</td>
    <td align="center">중</td>
    <td align="center">저</td>
    <td align="center">-</td>
    <td align="center"><select name="sel2_6" id="sel2_6"  onChange="convert_sum();">
      <option value="0" <%if sel2_6 = "sel2_6" and val2_6 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel2_6 = "sel2_6" and val2_6 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel2_6 = "sel2_6" and val2_6 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel2_6 = "sel2_6" and val2_6 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel2_6 = "sel2_6" and val2_6 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="2_6">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center" bgcolor="#FFFF99"><strong>소 계</strong></td>
    <td align="center" bgcolor="#FFFF99"><span class="style2">40</span></td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99"><span id="sum_2">0</span></td>
  </tr>
  <tr>
    <td rowspan="4" align="center">3.가시환경</td>
    <td height="30" colspan="2" align="center">가독성</td>
    <td align="center"><span class="style2">5</span></td>
    <td align="center">중거리가독</td>
    <td align="center">근거리가독</td>
    <td align="center">근접거리가독</td>
    <td align="center">-</td>
    <td align="center"><select name="sel3_1" id="sel3_1"  onChange="convert_sum();">
      <option value="0" <%if sel3_1 = "sel3_1" and val3_1 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel3_1 = "sel3_1" and val3_1 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel3_1 = "sel3_1" and val3_1 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel3_1 = "sel3_1" and val3_1 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel3_1 = "sel3_1" and val3_1 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="3_1">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">가시장애요인</td>
    <td align="center"><span class="style2">5</span></td>
    <td align="center">전무</td>
    <td align="center">부분장애</td>
    <td align="center">복합장애</td>
    <td align="center">-</td>
    <td align="center"><select name="sel3_2" id="sel3_2"  onChange="convert_sum();">
      <option value="0" <%if sel3_2 = "sel3_2" and val3_2 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel3_2 = "sel3_2" and val3_2 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel3_2 = "sel3_2" and val3_2 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel3_2 = "sel3_2" and val3_2 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel3_2 = "sel3_2" and val3_2 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="3_2">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">주변매체 경쟁력 </td>
    <td align="center"><span class="style2">5</span></td>
    <td align="center">고</td>
    <td align="center">중</td>
    <td align="center">저</td>
    <td align="center">-</td>
    <td align="center"><select name="sel3_3" id="sel3_3"  onChange="convert_sum();">
      <option value="0" <%if sel3_3 = "sel3_3" and val3_3 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel3_3 = "sel3_3" and val3_3 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel3_3 = "sel3_3" and val3_3 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel3_3 = "sel3_3" and val3_3 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel3_3 = "sel3_3" and val3_3 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="3_3">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center" bgcolor="#FFFF99"><strong>소 계</strong></td>
    <td align="center" bgcolor="#FFFF99"><span class="style2">15</span></td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99"><span id="sum_3">0</span></td>
  </tr>
  <tr>
    <td rowspan="4" align="center">4.기타항목</td>
    <td height="30" colspan="2" align="center">업체관리능력</td>
    <td align="center"><span class="style2">5</span></td>
    <td align="center">우수</td>
    <td align="center">보통</td>
    <td align="center">-</td>
    <td align="center">-</td>
    <td align="center"><select name="sel4_1" id="sel4_1"  onChange="convert_sum();">
      <option value="0" <%if sel4_1 = "sel4_1" and val4_1 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel4_1 = "sel4_1" and val4_1 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel4_1 = "sel4_1" and val4_1 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel4_1 = "sel4_1" and val4_1 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel4_1 = "sel4_1" and val4_1 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="4_1">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">허가여건</td>
    <td align="center"><span class="style2">5</span></td>
    <td align="center">합법</td>
    <td align="center">-</td>
    <td align="center">불법</td>
    <td align="center">-</td>
    <td align="center"><select name="sel4_2" id="sel4_2"  onChange=" convert_sum();">
      <option value="0" <%if sel4_2 = "sel4_2" and val4_2 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel4_2 = "sel4_2" and val4_2 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel4_2 = "sel4_2" and val4_2 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel4_2 = "sel4_2" and val4_2 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel4_2 = "sel4_2" and val4_2 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="4_2">0</span></td>
  </tr>
  <tr>
    <td height="14" colspan="2" align="center">경제성</td>
    <td align="center">5</td>
    <td align="center">저렴</td>
    <td align="center">적정</td>
    <td align="center">고가</td>
    <td align="center">-</td>
    <td align="center"><select name="sel4_3" id="sel4_3"  onChange=" convert_sum();">
      <option value="0" <%if sel4_3 = "sel4_3" and val4_3 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel4_3 = "sel4_3" and val4_3 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel4_3 = "sel4_3" and val4_3 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel4_3 = "sel4_3" and val4_3 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel4_3 = "sel4_3" and val4_3 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="4_3">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center" bgcolor="#FFFF99"><strong>소 계</strong></td>
    <td align="center" bgcolor="#FFFF99"><span class="style2">15</span></td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99"><span id="sum_4">0</span></td>
  </tr>
  <tr>
    <td height="48" colspan="3" align="center"><strong>총계</strong></td>
    <td align="center"><span class="style3">&nbsp;</span></td>
    <td align="center">&nbsp;</td>
    <td align="center">&nbsp;</td>
    <td align="center">매체등급</td>
      <td align="center"><span id="md_class">&nbsp;</span><input type="hidden" name="txtclass"></td>
      <td align="center"><span id="sum_avg">0</span><input type="hidden" name="txtavg"></td>
    <td align="center"><span id="sum_total">0</span></td>
  </tr>
  <tr>
    <td align="right" colspan="10" height="50"><img src="/images/btn_save.gif" width="59" height="18" vspace="5" onClick="check_submit();" style="cursor:hand" hspace="10"><img src="/images/btn_close.gif" width="57" height="18" vspace="5" style="cursor:hand" onClick="set_close();" ></td>
  </tr>
</table>
<!--  --></td>
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

	function check_submit() {
		var frm = document.forms[0];
		frm.action = "validation_etc_proc.asp";
		frm.method = "post";
		frm.submit();
	}


	function  convert_sum() {
		var frm = document.forms[0];
		document.getElementById('1_1').innerText = frm.sel1_1.options[frm.sel1_1.selectedIndex].value * 5;
		document.getElementById('1_2').innerText = frm.sel1_2.options[frm.sel1_2.selectedIndex].value * 5;
		document.getElementById('1_3').innerText = frm.sel1_3.options[frm.sel1_3.selectedIndex].value * 5;
		document.getElementById('1_4').innerText = frm.sel1_4.options[frm.sel1_4.selectedIndex].value * 10;
		document.getElementById('1_5').innerText = frm.sel1_4.options[frm.sel1_4.selectedIndex].value * 5;

		document.getElementById('2_1').innerText = frm.sel2_1.options[frm.sel2_1.selectedIndex].value * 10;
		document.getElementById('2_2').innerText = frm.sel2_2.options[frm.sel2_2.selectedIndex].value * 5;
		document.getElementById('2_3').innerText = frm.sel2_3.options[frm.sel2_3.selectedIndex].value * 5;
		document.getElementById('2_4').innerText = frm.sel2_4.options[frm.sel2_4.selectedIndex].value * 10;
		document.getElementById('2_5').innerText = frm.sel2_5.options[frm.sel2_5.selectedIndex].value * 5;
		document.getElementById('2_6').innerText = frm.sel2_6.options[frm.sel2_6.selectedIndex].value * 5;

		document.getElementById('3_1').innerText = frm.sel3_1.options[frm.sel3_1.selectedIndex].value * 5;
		document.getElementById('3_2').innerText = frm.sel3_2.options[frm.sel3_2.selectedIndex].value * 5;
		document.getElementById('3_3').innerText = frm.sel3_3.options[frm.sel3_3.selectedIndex].value * 5;

		document.getElementById('4_1').innerText = frm.sel4_1.options[frm.sel4_1.selectedIndex].value * 5;
		document.getElementById('4_2').innerText = frm.sel4_2.options[frm.sel4_2.selectedIndex].value * 5;
		document.getElementById('4_3').innerText = frm.sel4_3.options[frm.sel4_3.selectedIndex].value * 5;

		var _1_1 = document.getElementById("1_1").innerText;
		var _1_2 = document.getElementById("1_2").innerText;
		var _1_3 = document.getElementById("1_3").innerText;
		var _1_4 = document.getElementById("1_4").innerText;
		var _1_5 = document.getElementById("1_5").innerText;
		var _sum_1 = parseInt(_1_1) + parseInt(_1_2) + parseInt(_1_3) + parseInt(_1_4) + parseInt(_1_5);
		document.getElementById("sum_1").innerText = _sum_1 ;

		var _2_1 = document.getElementById("2_1").innerText;
		var _2_2 = document.getElementById("2_2").innerText;
		var _2_3 = document.getElementById("2_3").innerText;
		var _2_4 = document.getElementById("2_4").innerText;
		var _2_5 = document.getElementById("2_5").innerText;
		var _2_6 = document.getElementById("2_6").innerText;
		var _sum_2 = parseInt(_2_1) + parseInt(_2_2) + parseInt(_2_3) + parseInt(_2_4) + parseInt(_2_5) + parseInt(_2_6);
		document.getElementById("sum_2").innerText = _sum_2 ;

		var _3_1 = document.getElementById("3_1").innerText;
		var _3_2 = document.getElementById("3_2").innerText;
		var _3_3 = document.getElementById("3_3").innerText;
		var _sum_3 = parseInt(_3_1) + parseInt(_3_2) + parseInt(_3_3) ;
		document.getElementById("sum_3").innerText = _sum_3 ;

		var _4_1 = document.getElementById("4_1").innerText;
		var _4_2 = document.getElementById("4_2").innerText;
		var _4_3 = document.getElementById("4_3").innerText;
		var _sum_4 = parseInt(_4_1) + parseInt(_4_2) + parseInt(_4_3) ;
		document.getElementById("sum_4").innerText = _sum_4 ;

		document.getElementById("sum_total").innerText = parseInt(_sum_1) + parseInt(_sum_2) + parseInt(_sum_3) + parseInt(_sum_4);

		document.getElementById("sum_avg").innerText = parseInt(document.getElementById("sum_total").innerText) / 4;
		frm.txtavg.value = parseInt(document.getElementById("sum_total").innerText) / 4 ;

		var sum_avg = parseFloat(document.getElementById("sum_avg").innerText) ;
//		document.getElementById("md_avg").innerText = sum_avg;
		if (sum_avg >= 92.93 ) {
			document.getElementById("md_class").innerText = "SA" ;
			frm.txtclass.value = "SA" ;
//			document.getElementById("class_weight").innerText = "1.12" ;
		} else if (88.71 <= sum_avg && sum_avg <= 92.92) {
			document.getElementById("md_class").innerText = "A" ;
			frm.txtclass.value = "A" ;
//			document.getElementById("class_weight").innerText = "1.06" ;
		} else if (80.26 <= sum_avg && sum_avg <= 88.70) {
			document.getElementById("md_class").innerText = "B" ;
			frm.txtclass.value = "B" ;
//			document.getElementById("class_weight").innerText = "1.00" ;
		} else if (76.03 <= sum_avg && sum_avg <= 80.25) {
			document.getElementById("md_class").innerText = "C" ;
			frm.txtclass.value = "C" ;
//			document.getElementById("class_weight").innerText = "0.94" ;
		} else if (76.02 >= sum_avg ) {
			document.getElementById("md_class").innerText = "D" ;
			frm.txtclass.value = "D" ;
//			document.getElementById("class_weight").innerText = "0.88" ;
		}
	}
//
//	function set_unitprice(p) {
//		document.getElementById("unitprice").innerText = Number(String(p.options[p.selectedIndex].value).replace(/[^\d]/g,"")).toLocaleString().slice(0,-3);
//	}
//
//	function check_led() {
//		var frm = document.forms[0];
//		if (frm.sel_led.selectedIndex == 0) {
//			alert("LED 종류를 먼저 선택하세요");
//			frm.sel_led.focus();
//			return false ;
//		}
//	}
//
//	function sum_stand_price() {
//		var frm = document.forms[0];
//		var class_weight = parseFloat(document.getElementById("class_weight").innerText);
//		var unitprice = parseInt(document.getElementById("unitprice").innerText.replace(/,/g,""));
//		var qty = parseFloat(frm.txt_led.value);
//		var properprice = class_weight * unitprice * qty ;
//		document.getElementById("properprice").innerText = properprice.toLocaleString().slice(0,-3);
//		var monthprice = document.getElementById("monthprice").innerText ;
//		monthprice = isNaN(monthprice)?0:monthprice ;
//		var validation = document.getElementById("validation");
//		if (monthprice != 0) validation = properprice / monthprice * 100 ;
//		else validation = 120 ;
//
//		if (validation >= 116) document.getElementById("validation_class").innerText = "SA" ;
//		else if (106 <= validation && validation <=115) document.getElementById("validation_class").innerText = "A" ;
//		else if (96 <= validation && validation <=105) document.getElementById("validation_class").innerText = "B" ;
//		else if (86 <= validation && validation <=95) document.getElementById("validation_class").innerText = "C" ;
//		else if (validation <=85) document.getElementById("validation_class").innerText = "D" ;
//	}

	function set_reset() {
		document.forms[0].reset();
	}

	function set_close() {
		this.close();
	}

	window.onload = function () {
		self.focus();
		convert_sum();
	}

//-->
</SCRIPT>
