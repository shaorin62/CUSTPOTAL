<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim tidx : tidx = request("tidx")
	dim monthprice : monthprice = request("monthprice")
	dim mdidx , avg, mclass, objrs, sql, panaqty, panaprice, neonqty, neonprice, validation, validationclass

	sql = "select c.mdidx, c.avg, c.class, t.pana, t.panaprice, t.neon, t.neonprice, validfee, validclass from dbo.wb_validation_class c left outer  join dbo.wb_validation_tool t on c.tidx = t.tidx where c.tidx = " & tidx

	call get_recordset(objrs, sql)

	if not objrs.eof then
		avg = objrs("avg").value
		mclass = objrs("class").value
		mdidx = objrs("mdidx").value
		panaqty = objrs("pana").value
		panaprice = objrs("panaprice").value
		neonqty = objrs("neon").value
		neonprice = objrs("neonprice").value
		validation = objrs("validfee").value
		validationclass = objrs("validclass").value
		if isnull(panaqty) then panaqty = 0
		if isnull(panaprice) then panaprice = 0
		if isnull(neonqty) then neonqty = 0
		if isnull(neonprice) then neonprice = 0
	end if
	objrs.close


	sql = "select title from dbo.wb_medium_mst where mdidx = " & mdidx
	call get_recordset(objrs, sql)

	dim title : title = objrs(0).value
	objrs.close

	sql = "select v.tidx, v.code, v.value from dbo.wb_validation_value v inner join dbo.wb_validation_class  c on c.tidx = v.tidx where c.tidx = " & tidx

	call get_recordset(objrs, sql)
	dim sel1_1, val1_1
	dim sel1_2, val1_2
	dim sel1_3, val1_3
	dim sel1_4, val1_4
	dim sel1_5, val1_5
	dim sel1_6, val1_6
	dim sel2_1, val2_1
	dim sel2_2, val2_2
	dim sel2_3, val2_3
	dim sel2_4, val2_4
	dim sel2_5, val2_5
	dim sel3_1, val3_1
	dim sel3_2, val3_2
	dim sel3_3, val3_3
	dim sel3_4, val3_4
	dim sel3_5, val3_5
	dim sel4_1, val4_1
	dim sel4_2, val4_2
	dim sel4_3, val4_3
	dim sel5_1, val5_1
	dim sel5_2, val5_2
	dim sel5_3, val5_3
	dim sel5_4, val5_4
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
		if objrs("code") = "sel1_6" then
			sel1_6 = objrs("code").value
			val1_6 = objrs("value").value
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
		if objrs("code") = "sel3_4" then
			sel3_4 = objrs("code").value
			val3_4 = objrs("value").value
		end if
		if objrs("code") = "sel3_5" then
			sel3_5 = objrs("code").value
			val3_5 = objrs("value").value
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
'5
		if objrs("code") = "sel5_1" then
			sel5_1 = objrs("code").value
			val5_1 = objrs("value").value
		end if
		if objrs("code") = "sel5_2" then
			sel5_2 = objrs("code").value
			val5_2 = objrs("value").value
		end if
		if objrs("code") = "sel5_3" then
			sel5_3 = objrs("code").value
			val5_3 = objrs("value").value
		end if
		if objrs("code") = "sel5_4" then
			sel5_4 = objrs("code").value
			val5_4 = objrs("value").value
		end if
		objrs.movenext
		loop
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

<body  oncontextmenu="return false">
<form>
<INPUT TYPE="hidden" NAME="tidx" value="<%=tidx%>">
<table width="875" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top">  <%=title%> 효용성 평가 TOOL </td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="875" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="22"><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td>
<!--  -->
  <table width="830" border="1" cellpadding="0" cellspacing="1" bordercolor="#CCCCCC">
    <tr>
      <td width="100" height="25" align="center"><strong>매체구분</strong></td>
      <td width="100" align="center" bgcolor="#FFFFFF">옥탑</td>
      <td colspan="5" rowspan="2" align="center" bgcolor="#FFFFFF"><strong>규 격(㎡/구좌) </strong></td>
      <td width="0" rowspan="2" align="center"><strong>광고단가(㎡)</strong></td>
      <td align="center" bgcolor="#FFFFFF"><span id="unitprice">&nbsp;</span></td>
    </tr>
    <tr>
      <td height="25" align="center"><strong>매체등급</strong></td>
      <td width="100" align="center" bgcolor="#FFFFFF"><span id="md_class">&nbsp;</span></td>
      <td width="100" align="center" bgcolor="#FFFFFF"><span id="unitprice2">&nbsp;</span></td>
    </tr>
    <tr>
      <td height="25" align="center"><strong>매체평점</strong></td>
      <td align="center" bgcolor="#FFFFFF"><span id="md_avg">&nbsp;</span></td>
      <td width="80" align="center"><strong>네온</strong></td>
      <td width="80" align="center"><strong>파나</strong></td>
      <td width="80" align="center"><strong>야립</strong></td>
      <td width="80" align="center"><strong>LED</strong></td>
      <td width="80" align="center"><strong>화공</strong></td>
      <td width="0" align="center"><strong>적정광고료</strong></td>
      <td width="100" align="center" bgcolor="#FFFFFF"><span id="properprice">&nbsp;</span></td>
    </tr>
    <tr>
      <td height="25" align="center"><strong>가 중 치</strong></td>
      <td align="center" bgcolor="#FFFFFF"><span id="class_weight">&nbsp;</span></td>
      <td bgcolor="#FFFFFF"><select name="sel_neon" id="sel_neon" style="width:80px;height:25px"  onChange="set_unitprice(this);">
          <option  value="">선택</option>
          <option value="200000" <%if neonprice = "200000" then response.write "selected"%>>알파네온</option>
          <option value="100000" <%if neonprice = "100000" then response.write "selected"%>>점멸네온</option>
          <option value="55000" <%if neonprice = "55000" then response.write "selected"%>>단순네온</option>
      </select></td>
      <td bgcolor="#FFFFFF"><select name="sel_pana" id="sel_pana" style="width:80;height:25px"  onChange="set_unitprice2(this);">
          <option  value="">선택</option>
          <option value="55000" <%if panaprice = "55000" then response.write "selected"%>>파나</option>
          <option value="125000" <%if panaprice = "125000" then response.write "selected"%>>기금광고</option>
      </select></td>
      <td bgcolor="#FFFFFF"><select name="sel_board" id="sel_board" style="width:80;height:25px" disabled>
          <option>선택</option>
          <option value="75000">야립</option>
      </select></td>
      <td bgcolor="#FFFFFF"><select name="sel_led" id="sel_led"  style="width:80;height:25px"  disabled>
          <option value="" >선택</option>
          <option value="10000000" >SA</option>
          <option value="8000000" >A</option>
          <option value="6000000" >B</option>
          <option value="4000000" >C</option>
      </select></td>
      <td bgcolor="#FFFFFF"><select name="sel_paint" id="sel_paint"  style="width:80;height:25px" disabled>
          <option value="">선택</option>
          <option value="0">화공</option>
      </select></td>
      <td width="0" align="center"><strong>효용성</strong></td>
      <td width="100" align="center" bgcolor="#FFFFFF"><span id="validation">&nbsp;</span><input type="hidden" name="txtvalidation"></td>
    </tr>
    <tr>
      <td height="25" align="center"><strong>집행 광고료(월)</strong></td>
      <td align="center" bgcolor="#FFFFFF"><span id="monthprice"><%=formatnumber(monthprice,0)%></span></td>
      <td><input type="text" name="txt_neon" id="txt_neon" style="width:80;height:25px" onFocus="check_neon();"  onblur="sum_stand_price();"></td>
      <td><input type="text" name="txt_pana" id="txt_pana" style="width:80;height:25px" onFocus="check_pana();"  onblur="sum_stand_price();"></td>
      <td><input type="text" name="txt_board" id="txt_board" style="width:80;height:25px;text-align:right;" readonly></td>
      <td><input type="text" name="txt_led" id="txt_led" style="width:80;height:25px;text-align:center" readonly></td>
      <td><input type="text" name="txt_paint" id="txt_paint" style="width:80;height:25px" readonly></td>
      <td width="0" align="center"><strong>효용성 등급</strong></td>
      <td width="100" align="center" bgcolor="#FFFFFF"><span id="validation_class">&nbsp;</span><input type="hidden" name="txtvalidationclass"></td>
    </tr>
  </table>
<table width="830" border="1" cellpadding="3" cellspacing="1" bordercolor="#CCCCCC">
  <tr>
    <td colspan="3" rowspan="2" align="center"><strong>평가항목</strong></td>
    <td rowspan="2" align="center"><strong>가중치</strong></td>
    <td height="25" colspan="4" align="center"><strong>평가기준</strong></td>
    <td rowspan="2" align="center"><strong>평가</strong></td>
    <td rowspan="2" align="center"><strong>환산점수<br>
      (가중치X평가)</strong></td>
  </tr>
  <tr>
    <td width="90" height="25" align="center"><strong>A(4)</strong></td>
    <td width="90" align="center"><strong>B(3)</strong></td>
    <td width="90" align="center"><strong>C(2)</strong></td>
    <td width="90" align="center"><strong>D(1)</strong></td>
  </tr>
  <tr>
    <td rowspan="8" align="center">1. 지역환경</td>
    <td height="30" colspan="2" align="center">설치지역</td>
    <td align="center"><span class="style2">5</span></td>
    <td align="center">수도권/광역시</td>
    <td align="center">도별거점도시</td>
    <td align="center">일반시지역</td>
    <td align="center">읍면지역</td>
    <td align="center"><select name="sel1_1" id="sel1_1" onChange="convert_sum();" disabled>
      <option value="0"  <%if sel1_1 = "sel1_1" and val1_1 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel1_1 = "sel1_1" and val1_1 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel1_1 = "sel1_1" and val1_1 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel1_1 = "sel1_1" and val1_1 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel1_1 = "sel1_1" and val1_1 = "1" then response.write "selected"%>>D</option>
    </select>    </td>
    <td align="center"><span id="1_1">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">지역상권력</td>
    <td align="center"><span class="style2">5</span></td>
    <td align="center">핵심상권<br>
      /전략지역</td>
    <td align="center">지역상권(역세권)</td>
    <td align="center">일반상권</td>
    <td align="center">비상권</td>
    <td align="center"><select name="sel1_2" id="sel1_2"  onChange="convert_sum();" disabled>
      <option value="0"  <%if sel1_2 = "sel1_2" and val1_2 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel1_2 = "sel1_2" and val1_2 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel1_2 = "sel1_2" and val1_2 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel1_2 = "sel1_2" and val1_2 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel1_2 = "sel1_2" and val1_2 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="1_2">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">유동인구</td>
    <td align="center"><span class="style2">5</span></td>
    <td align="center">고밀도지역</td>
    <td align="center">적정지역</td>
    <td align="center">저밀도지역</td>
    <td align="center">-</td>
    <td align="center"><select name="sel1_3" id="sel1_3"  onChange="convert_sum();" disabled>
      <option value="0" <%if sel1_3 = "sel1_3" and val1_3 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel1_3 = "sel1_3" and val1_3 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel1_3 = "sel1_3" and val1_3 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel1_3 = "sel1_3" and val1_3 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel1_3 = "sel1_3" and val1_3 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="1_3">0</span></td>
  </tr>
  <tr>
    <td rowspan="2" align="center">유동<br>
      차량</td>
    <td height="30" align="center">서울,부산,대구</td>
    <td rowspan="2" align="center"><span class="style2">10</span></td>
    <td align="center">10만 이상</td>
    <td align="center">8∼10만</td>
    <td align="center">5∼8만</td>
    <td align="center">6만이하</td>
    <td align="center" rowspan=2><select name="sel1_4" id="sel1_4"  onChange="convert_sum();" disabled>
      <option value="0" <%if sel1_4 = "sel1_4" and val1_4 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel1_4 = "sel1_4" and val1_4 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel1_4 = "sel1_4" and val1_4 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel1_4 = "sel1_4" and val1_4 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel1_4 = "sel1_4" and val1_4 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center" rowspan="2"><span id="1_4">0</span></td>
  </tr>
  <tr>
    <td height="30" align="center">기타 지역</td>
    <td align="center">20%이내</td>
    <td align="center">30%이내</td>
    <td align="center">40%이내</td>
    <td align="center">40%이상</td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">도로폭</td>
    <td align="center"><span class="style2">3</span></td>
    <td align="center">8차선이상</td>
    <td align="center">6차선</td>
    <td align="center">4차선</td>
    <td align="center">4차선미만</td>
    <td align="center"><select name="sel1_5" id="sel1_5"  onChange="convert_sum();" disabled>
      <option value="0" <%if sel1_5 = "sel1_5" and val1_5 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel1_5 = "sel1_5" and val1_5 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel1_5 = "sel1_5" and val1_5 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel1_5 = "sel1_5" and val1_5 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel1_5 = "sel1_5" and val1_5 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="1_5">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">차량 정체도</td>
    <td align="center"><span class="style2">2</span></td>
    <td align="center">시속15㎞미만</td>
    <td align="center">30km이하</td>
    <td align="center">60km이하</td>
    <td align="center">60km이상</td>
    <td align="center"><select name="sel1_6" id="sel1_6"  onChange="convert_sum();" disabled>
      <option value="0" <%if sel1_6 = "sel1_6" and val1_6 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel1_6 = "sel1_6" and val1_6 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel1_6 = "sel1_6" and val1_6 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel1_6 = "sel1_6" and val1_6 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel1_6 = "sel1_6" and val1_6 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="1_6">0</span></td>
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
    <td rowspan="6" align="center">2.매체사양</td>
    <td height="30" colspan="2" align="center">광고물 규격</td>
    <td align="center"><span class="style2">10</span></td>
    <td align="center">170㎡</td>
    <td align="center">140㎡</td>
    <td align="center">100㎡</td>
    <td align="center">100㎡</td>
    <td align="center"><select name="sel2_1" id="sel2_1"  onChange="convert_sum();" disabled>
      <option value="0" <%if sel2_1 = "sel2_1" and val2_1 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel2_1 = "sel2_1" and val2_1 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel2_1 = "sel2_1" and val2_1 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel2_1 = "sel2_1" and val2_1 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel2_1 = "sel2_1" and val2_1 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="2_1">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">건물 높이</td>
    <td align="center">2</td>
    <td align="center">8∼10층</td>
    <td align="center">6∼7층/11∼13층</td>
    <td align="center">4∼5층</td>
    <td align="center">기타</td>
    <td align="center"><select name="sel2_2" id="sel2_2"  onChange="convert_sum();" disabled>
      <option value="0" <%if sel2_2 = "sel2_2" and val2_2 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel2_2 = "sel2_2" and val2_2 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel2_2 = "sel2_2" and val2_2 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel2_2 = "sel2_2" and val2_2 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel2_2 = "sel2_2" and val2_2 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="2_2">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">광고면 소재</td>
    <td align="center"><span class="style2">5</span></td>
    <td align="center">점멸네온</td>
    <td align="center">파나/단순</td>
    <td align="center">외부조명</td>
    <td align="center">비조명</td>
    <td align="center"><select name="sel2_3" id="sel2_3"  onChange="convert_sum();" disabled>
      <option value="0" <%if sel2_3 = "sel2_3" and val2_3 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel2_3 = "sel2_3" and val2_3 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel2_3 = "sel2_3" and val2_3 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel2_3 = "sel2_3" and val2_3 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel2_3 = "sel2_3" and val2_3 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="2_3">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">조명 여부</td>
    <td align="center"><span class="style2">5</span></td>
    <td align="center">3면이상</td>
    <td align="center">2면이상</td>
    <td align="center">1면이상</td>
    <td align="center">비조명</td>
    <td align="center"><select name="sel2_4" id="sel2_4"  onChange="convert_sum();" disabled>
      <option value="0" <%if sel2_4 = "sel2_4" and val2_4 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel2_4 = "sel2_4" and val2_4 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel2_4 = "sel2_4" and val2_4 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel2_4 = "sel2_4" and val2_4 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel2_4 = "sel2_4" and val2_4 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="2_4">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">매체 경쟁력</td>
    <td align="center"><span class="style2">3</span></td>
    <td align="center">우세</td>
    <td align="center">비등</td>
    <td align="center">열세</td>
    <td align="center">절대 열세</td>
    <td align="center"><select name="sel2_5" id="sel2_5"  onChange="convert_sum();" disabled>
      <option value="0" <%if sel2_5 = "sel2_5" and val2_5 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel2_5 = "sel2_5" and val2_5 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel2_5 = "sel2_5" and val2_5 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel2_5 = "sel2_5" and val2_5 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel2_5 = "sel2_5" and val2_5 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="2_5">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center" bgcolor="#FFFF99"><strong>소 계</strong></td>
    <td align="center" bgcolor="#FFFF99"><span class="style2">25</span></td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99"><span id="sum_2">0</span></td>
  </tr>
  <tr>
    <td rowspan="6" align="center">3.가시환경</td>
    <td height="30" colspan="2" align="center">가시 거리</td>
    <td align="center"><span class="style2">10</span></td>
    <td align="center">1km이상</td>
    <td align="center">500m이상</td>
    <td align="center">300m이상</td>
    <td align="center">300m미만</td>
    <td align="center"><select name="sel3_1" id="sel3_1"  onChange="convert_sum();" disabled>
      <option value="0" <%if sel3_1 = "sel3_1" and val3_1 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel3_1 = "sel3_1" and val3_1 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel3_1 = "sel3_1" and val3_1 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel3_1 = "sel3_1" and val3_1 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel3_1 = "sel3_1" and val3_1 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="3_1">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">가시 상태</td>
    <td align="center"><span class="style2">5</span></td>
    <td align="center">강제주시</td>
    <td align="center">자연가시</td>
    <td align="center">의도가시</td>
    <td align="center">-</td>
    <td align="center"><select name="sel3_2" id="sel3_2"  onChange="convert_sum();" disabled>
      <option value="0" <%if sel3_2 = "sel3_2" and val3_2 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel3_2 = "sel3_2" and val3_2 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel3_2 = "sel3_2" and val3_2 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel3_2 = "sel3_2" and val3_2 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel3_2 = "sel3_2" and val3_2 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="3_2">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">가시 장애 요인</td>
    <td align="center"><span class="style2">3</span></td>
    <td align="center">전무</td>
    <td align="center">장거리장애</td>
    <td align="center">중근거리장애</td>
    <td align="center">반복장애</td>
    <td align="center"><select name="sel3_3" id="sel3_3"  onChange="convert_sum();" disabled>
      <option value="0" <%if sel3_3 = "sel3_3" and val3_3 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel3_3 = "sel3_3" and val3_3 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel3_3 = "sel3_3" and val3_3 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel3_3 = "sel3_3" and val3_3 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel3_3 = "sel3_3" and val3_3 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="3_3">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">노출 시간</td>
    <td align="center"><span class="style2">2</span></td>
    <td align="center">1분이상</td>
    <td align="center">30초</td>
    <td align="center">10초</td>
    <td align="center">10초미만</td>
    <td align="center"><select name="sel3_4" id="sel3_4"  onChange="convert_sum();" disabled>
      <option value="0" <%if sel3_4 = "sel3_4" and val3_4 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel3_4 = "sel3_4" and val3_4 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel3_4 = "sel3_4" and val3_4 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel3_4 = "sel3_4" and val3_4 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel3_4 = "sel3_4" and val3_4 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="3_4">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">노출 방향</td>
    <td align="center"><span class="style2">5</span></td>
    <td align="center">3방향이상</td>
    <td align="center">2방향</td>
    <td align="center">퇴근1방향</td>
    <td align="center">출근1방향</td>
    <td align="center"><select name="sel3_5" id="sel3_5"  onChange="convert_sum();" disabled>
      <option value="0" <%if sel3_5 = "sel3_5" and val3_5 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel3_5 = "sel3_5" and val3_5 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel3_5 = "sel3_5" and val3_5 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel3_5 = "sel3_5" and val3_5 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel3_5 = "sel3_5" and val3_5 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="3_5">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center" bgcolor="#FFFF99"><strong>소 계</strong></td>
    <td align="center" bgcolor="#FFFF99"><span class="style2">25</span></td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99"><span id="sum_3">0</span></td>
  </tr>
  <tr>
    <td rowspan="4" align="center">4.경쟁환경</td>
    <td height="30" colspan="2" align="center">전략성</td>
    <td align="center"><span class="style2">4</span></td>
    <td align="center">높음</td>
    <td align="center">보통</td>
    <td align="center">낮음</td>
    <td align="center">-</td>
    <td align="center"><select name="sel4_1" id="sel4_1"  onChange="convert_sum();" disabled>
      <option value="0" <%if sel4_1 = "sel4_1" and val4_1 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel4_1 = "sel4_1" and val4_1 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel4_1 = "sel4_1" and val4_1 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel4_1 = "sel4_1" and val4_1 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel4_1 = "sel4_1" and val4_1 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="4_1">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">상징성</td>
    <td align="center"><span class="style2">3</span></td>
    <td align="center">지역화제성</td>
    <td align="center">일반상징성</td>
    <td align="center">보편성</td>
    <td align="center">진부</td>
    <td align="center"><select name="sel4_2" id="sel4_2"  onChange=" convert_sum();" disabled>
      <option value="0" <%if sel4_2 = "sel4_2" and val4_2 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel4_2 = "sel4_2" and val4_2 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel4_2 = "sel4_2" and val4_2 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel4_2 = "sel4_2" and val4_2 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel4_2 = "sel4_2" and val4_2 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="4_2">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">소구 대상</td>
    <td align="center">3</td>
    <td align="center">O.L / 실고객층</td>
    <td align="center">잠재고객층</td>
    <td align="center">비구매층</td>
    <td align="center">-</td>
    <td align="center"><select name="sel4_3" id="sel4_3"  onChange=" convert_sum();" disabled>
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
    <td align="center" bgcolor="#FFFF99"><span class="style2">10</span></td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99"><span id="sum_4">0</span></td>
  </tr>
  <tr>
    <td rowspan="5" align="center">5.기타항목</td>
    <td height="30" colspan="2" align="center">단위 광고료</td>
    <td align="center"><span class="style2">5</span></td>
    <td align="center">단가지수 0.8 ~ 1.0</td>
    <td align="center">1.0 ~ 1.2</td>
    <td align="center">미흡1.3이상</td>
    <td align="center">-</td>
    <td align="center"><select name="sel5_1" id="sel5_1"  onChange="convert_sum();" disabled>
      <option value="0" <%if sel5_1 = "sel5_1" and val5_1 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel5_1 = "sel5_1" and val5_1 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel5_1 = "sel5_1" and val5_1 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel5_1 = "sel5_1" and val5_1 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel5_1 = "sel5_1" and val5_1 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="5_1">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">업체 협력도</td>
    <td align="center"><span class="style2">2</span></td>
    <td align="center">우수</td>
    <td align="center">양호</td>
    <td align="center">미흡</td>
    <td align="center">불량</td>
    <td align="center"><select name="sel5_2" id="sel5_2"  onChange="convert_sum();" disabled>
      <option value="0" <%if sel5_2 = "sel5_2" and val5_2 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel5_2 = "sel5_2" and val5_2 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel5_2 = "sel5_2" and val5_2 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel5_2 = "sel5_2" and val5_2 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel5_2 = "sel5_2" and val5_2 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="5_2">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">허가여건</td>
    <td align="center">1</td>
    <td align="center">준수</td>
    <td align="center">미준수</td>
    <td align="center">-</td>
    <td align="center">미허가</td>
    <td align="center"><select name="sel5_3" id="sel5_3"  onChange="convert_sum();" disabled>
      <option value="0" <%if sel5_3 = "sel5_3" and val5_3 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel5_3 = "sel5_3" and val5_3 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel5_3 = "sel5_3" and val5_3 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel5_3 = "sel5_3" and val5_3 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel5_3 = "sel5_3" and val5_3 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="5_3">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center">건물 및 주변환경</td>
    <td align="center">2</td>
    <td align="center">우수</td>
    <td align="center">양호</td>
    <td align="center">불량</td>
    <td align="center">- </td>
    <td align="center"><select name="sel5_4" id="sel5_4"  onChange="convert_sum();" disabled>
      <option value="0" <%if sel5_4 = "sel5_4" and val5_4 = "0" then response.write "selected"%>></option>
      <option value="4" <%if sel5_4 = "sel5_4" and val5_4 = "4" then response.write "selected"%>>A</option>
      <option value="3" <%if sel5_4 = "sel5_4" and val5_4 = "3" then response.write "selected"%>>B</option>
      <option value="2" <%if sel5_4 = "sel5_4" and val5_4 = "2" then response.write "selected"%>>C</option>
      <option value="1" <%if sel5_4 = "sel5_4" and val5_4 = "1" then response.write "selected"%>>D</option>
    </select></td>
    <td align="center"><span id="5_4">0</span></td>
  </tr>
  <tr>
    <td height="30" colspan="2" align="center" bgcolor="#FFFF99"><strong>소계</strong></td>
    <td align="center" bgcolor="#FFFF99"><span class="style2">10</span></td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99">&nbsp;</td>
    <td align="center" bgcolor="#FFFF99"><span id="sum_5">0</span></td>
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
		frm.action = "validation_tool_neon_proc.asp";
		frm.method = "post";
		frm.submit();
	}


	function  convert_sum() {
		var frm = document.forms[0];
		document.getElementById('1_1').innerText = frm.sel1_1.options[frm.sel1_1.selectedIndex].value * 5;
		document.getElementById('1_2').innerText = frm.sel1_2.options[frm.sel1_2.selectedIndex].value * 5;
		document.getElementById('1_3').innerText = frm.sel1_3.options[frm.sel1_3.selectedIndex].value * 5;
		document.getElementById('1_4').innerText = frm.sel1_4.options[frm.sel1_4.selectedIndex].value * 10;
		document.getElementById('1_5').innerText = frm.sel1_5.options[frm.sel1_5.selectedIndex].value * 3;
		document.getElementById('1_6').innerText = frm.sel1_6.options[frm.sel1_6.selectedIndex].value * 2;

		document.getElementById('2_1').innerText = frm.sel2_1.options[frm.sel2_1.selectedIndex].value * 10;
		document.getElementById('2_2').innerText = frm.sel2_2.options[frm.sel2_2.selectedIndex].value * 2;
		document.getElementById('2_3').innerText = frm.sel2_3.options[frm.sel2_3.selectedIndex].value * 5;
		document.getElementById('2_4').innerText = frm.sel2_4.options[frm.sel2_4.selectedIndex].value * 5;
		document.getElementById('2_5').innerText = frm.sel2_5.options[frm.sel2_5.selectedIndex].value * 3;

		document.getElementById('3_1').innerText = frm.sel3_1.options[frm.sel3_1.selectedIndex].value * 10;
		document.getElementById('3_2').innerText = frm.sel3_2.options[frm.sel3_2.selectedIndex].value * 5;
		document.getElementById('3_3').innerText = frm.sel3_3.options[frm.sel3_3.selectedIndex].value * 3;
		document.getElementById('3_4').innerText = frm.sel3_4.options[frm.sel3_4.selectedIndex].value * 2;
		document.getElementById('3_5').innerText = frm.sel3_5.options[frm.sel3_5.selectedIndex].value * 5;

		document.getElementById('4_1').innerText = frm.sel4_1.options[frm.sel4_1.selectedIndex].value * 4;
		document.getElementById('4_2').innerText = frm.sel4_2.options[frm.sel4_2.selectedIndex].value * 3;
		document.getElementById('4_3').innerText = frm.sel4_3.options[frm.sel4_3.selectedIndex].value * 3;

		document.getElementById('5_1').innerText = frm.sel5_1.options[frm.sel5_1.selectedIndex].value * 5;
		document.getElementById('5_2').innerText = frm.sel5_2.options[frm.sel5_2.selectedIndex].value * 2;
		document.getElementById('5_3').innerText = frm.sel5_3.options[frm.sel5_3.selectedIndex].value * 1;
		document.getElementById('5_4').innerText = frm.sel5_4.options[frm.sel5_4.selectedIndex].value * 2;

		var _1_1 = document.getElementById("1_1").innerText;
		var _1_2 = document.getElementById("1_2").innerText;
		var _1_3 = document.getElementById("1_3").innerText;
		var _1_4 = document.getElementById("1_4").innerText;
		var _1_5 = document.getElementById("1_5").innerText;
		var _1_6 = document.getElementById("1_6").innerText;
		var _sum_1 = parseInt(_1_1) + parseInt(_1_2) + parseInt(_1_3) + parseInt(_1_4) + parseInt(_1_5) + parseInt(_1_6);
		document.getElementById("sum_1").innerText = _sum_1 ;

		var _2_1 = document.getElementById("2_1").innerText;
		var _2_2 = document.getElementById("2_2").innerText;
		var _2_3 = document.getElementById("2_3").innerText;
		var _2_4 = document.getElementById("2_4").innerText;
		var _2_5 = document.getElementById("2_5").innerText;
		var _sum_2 = parseInt(_2_1) + parseInt(_2_2) + parseInt(_2_3) + parseInt(_2_4) + parseInt(_2_5);
		document.getElementById("sum_2").innerText = _sum_2 ;

		var _3_1 = document.getElementById("3_1").innerText;
		var _3_2 = document.getElementById("3_2").innerText;
		var _3_3 = document.getElementById("3_3").innerText;
		var _3_4 = document.getElementById("3_4").innerText;
		var _3_5 = document.getElementById("3_5").innerText;
		var _sum_3 = parseInt(_3_1) + parseInt(_3_2) + parseInt(_3_3) + parseInt(_3_4) + parseInt(_3_5);
		document.getElementById("sum_3").innerText = _sum_3 ;

		var _4_1 = document.getElementById("4_1").innerText;
		var _4_2 = document.getElementById("4_2").innerText;
		var _4_3 = document.getElementById("4_3").innerText;
		var _sum_4 = parseInt(_4_1) + parseInt(_4_2) + parseInt(_4_3);
		document.getElementById("sum_4").innerText = _sum_4 ;

		var _5_1 = document.getElementById("5_1").innerText;
		var _5_2 = document.getElementById("5_2").innerText;
		var _5_3 = document.getElementById("5_3").innerText;
		var _5_4 = document.getElementById("5_4").innerText;
		var _sum_5 = parseInt(_5_1) + parseInt(_5_2) + parseInt(_5_3) + parseInt(_5_4);
		document.getElementById("sum_5").innerText = _sum_5 ;

		document.getElementById("sum_total").innerText = parseInt(_sum_1) + parseInt(_sum_2) + parseInt(_sum_3) + parseInt(_sum_4) + parseInt(_sum_5) ;

		document.getElementById("sum_avg").innerText = parseInt(document.getElementById("sum_total").innerText) / 4;
		frm.txtavg.value = parseInt(document.getElementById("sum_total").innerText) / 4 ;

		var sum_avg = parseFloat(document.getElementById("sum_avg").innerText) ;
		if (sum_avg >= 92.93 ) {
			document.getElementById("md_class").innerText = "SA" ;
			frm.txtclass.value = "SA" ;
		} else if (88.71 <= sum_avg && sum_avg <= 92.92) {
			document.getElementById("md_class").innerText = "A" ;
			frm.txtclass.value = "A" ;
		} else if (80.26 <= sum_avg && sum_avg <= 88.70) {
			document.getElementById("md_class").innerText = "B" ;
			frm.txtclass.value = "B" ;
		} else if (76.03 <= sum_avg && sum_avg <= 80.25) {
			document.getElementById("md_class").innerText = "C" ;
			frm.txtclass.value = "C" ;
		} else if (76.02 >= sum_avg ) {
			document.getElementById("md_class").innerText = "D" ;
			frm.txtclass.value = "D" ;
		}
	}

	function set_unitprice(p) {
		document.getElementById("unitprice").innerText = Number(String(p.options[p.selectedIndex].value).replace(/[^\d]/g,"")).toLocaleString().slice(0,-3);
		sum_stand_price();
	}
	function set_unitprice2(p) {
		document.getElementById("unitprice2").innerText = Number(String(p.options[p.selectedIndex].value).replace(/[^\d]/g,"")).toLocaleString().slice(0,-3);
		sum_stand_price();
	}

	function check_neon() {
		var frm = document.forms[0];
		if (frm.sel_neon.selectedIndex == 0) {
			alert("네온 종류를 먼저 선택하세요");
			frm.sel_neon.focus();
			return false ;
		}
	}

	function check_pana() {
		var frm = document.forms[0];
		if (frm.sel_pana.selectedIndex == 0) {
			alert("파나 종류를 먼저 선택하세요");
			frm.sel_pana.focus();
			return false ;
		}
	}

	function sum_stand_price() {
		var frm = document.forms[0];
		var class_weight = parseFloat(document.getElementById("class_weight").innerText);
		var unitprice = parseInt(document.getElementById("unitprice").innerText.replace(/,/g,""));
		var unitprice2 = parseInt(document.getElementById("unitprice2").innerText.replace(/,/g,""));
		var qty = parseFloat(frm.txt_neon.value);
		var qty2 = parseFloat(frm.txt_pana.value);
		if (isNaN(unitprice)) unitprice = 0;
		if (isNaN(unitprice2)) unitprice2 = 0;
		var properprice = (class_weight * unitprice * qty) + (class_weight * unitprice2 * qty2) ;
		document.getElementById("properprice").innerText = properprice.toLocaleString().slice(0,-3);
		var monthprice = parseInt(document.getElementById("monthprice").innerText.replace(/,/g,""));
		var validation;
		if (monthprice != 0)
			validation = properprice / monthprice * 100 ;
		else
			validation = 120 ;
		document.getElementById("validation").innerText = validation.toFixed(1);

		if (validation >= 116) document.getElementById("validation_class").innerText = "SA" ;
		else if (106 <= validation && validation <=115) document.getElementById("validation_class").innerText = "A" ;
		else if (96 <= validation && validation <=105) document.getElementById("validation_class").innerText = "B" ;
		else if (86 <= validation && validation <=95) document.getElementById("validation_class").innerText = "C" ;
		else if (validation <=85) document.getElementById("validation_class").innerText = "D" ;

		frm.txtvalidation.value = document.getElementById("validation").innerText ;
		frm.txtvalidationclass.value = document.getElementById("validation_class").innerText;
	}

	function set_reset() {
		document.forms[0].reset();
	}

	function set_close() {
		this.close();
	}

	window.onload = function () {
		self.focus();
		convert_sum();
		document.getElementById("md_class").innerText = "<%=mclass%>";
		document.getElementById("md_avg").innerText = "<%=avg%>";
		var weight
		switch (document.getElementById("md_class").innerText) {
			case "SA":
				weight = 1.12 ;
			break;
			case "A":
				weight = 1.06 ;
			break;
			case "B":
				weight = 1.00 ;
			break ;
			case "C":
				weight = 0.94 ;
			break ;
			case "D":
				weight = 0.88 ;
			break ;
		}
		document.getElementById("class_weight").innerText = weight ;
		document.getElementById("unitprice").innerText = "<%=formatnumber(neonprice,0)%>" ;
		document.getElementById("unitprice2").innerText = "<%=formatnumber(panaprice,0)%>" ;
		document.forms[0].txt_neon.value = "<%=neonqty%>";
		document.forms[0].txt_pana.value = "<%=panaqty%>";
		sum_stand_price();
	}

//-->
</SCRIPT>
