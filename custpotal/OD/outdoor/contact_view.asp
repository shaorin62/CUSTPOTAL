<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim gotopage : gotopage = request("gotopage")
	dim searchstring : searchstring = request.Form("txtsearchstring")

	dim contidx : contidx = request("contidx")
	dim sidx : sidx = request("sidx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	if cyear = "" then cyear = year(date)
	if cmonth = "" then cmonth = month(date)

	dim objrs, sql
	if sidx = "" then
		sql = "select min(sidx) as sidx from dbo.wb_contact_md where contidx="&contidx
		call get_recordset(objrs, sql)
		if not objrs.eof then sidx = objrs(0).value
		objrs.close
	end if
	dim org_sidx : org_sidx = sidx


	sql = "select m.title, c.custname, c2.custname as custname2, q.totalqty, v.monthprice as totalprice, m.firstdate, m.startdate, m.enddate, m2.map,"&_
			"s.monthprice, s.expense, d.photo_1, d.photo_2, d.photo_3, d.photo_4, m.regionmemo, m.mediummemo, m.comment "&_
			"from dbo.wb_contact_mst m left outer join dbo.wb_contact_md m2 on m.contidx = m2.contidx "&_
			"left outer join dbo.sc_cust_temp c2 on m.custcode = c2.custcode "&_
			"left outer join dbo.sc_cust_temp c on c2.highcustcode = c.custcode "&_
			"left outer join dbo.vw_contact_md_summaryprice s on (m2.contidx = s.contidx and s.cyear = "&cyear&" and s.cmonth = "&cmonth&") "&_
			"left outer join dbo.wb_contact_md_dtl d on (m2.contidx = d.contidx and d.cyear = "&cyear&" and d.cmonth = "&cmonth&") "&_
			"left outer join dbo.vw_contact_totalprice v on m.contidx = v.contidx "&_
			"left outer join dbo.vw_contact_totalqty q on m.contidx = q.contidx "&_
			"where m.contidx="&contidx

	call get_recordset(objrs, sql)

	dim  title, custname, custname2, totalqty, totalprice, firstdate, startdate, enddate, monthprice, expense, income, incomeratio, photo_1, photo_2, photo_3, photo_4, regionmemo, mediummemo, comment, map

	if not objrs.eof Then
		title = objrs("title").value
		custname = objrs("custname").value		'광고주 custname
		custname2 = objrs("custname2").value	'사업부 custname2
		totalqty = objrs("totalqty").value
		totalprice = objrs("totalprice").value
		firstdate = objrs("firstdate").value
		startdate = objrs("startdate").value
		enddate = objrs("enddate").value
		monthprice = objrs("monthprice").value
		expense = objrs("expense").value
		photo_1 = objrs("photo_1").value
		photo_2 = objrs("photo_2").value
		photo_3 = objrs("photo_3").value
		photo_4 = objrs("photo_4").value
		regionmemo = objrs("regionmemo").value
		mediummemo = objrs("mediummemo").value
		comment = objrs("comment").value
		map = objrs("map").value
		if isnull(monthprice) then monthprice = 0
		if isnull(expense) then expense = 0
		income = monthprice - expense
		if income = 0 then incomeratio = "0.00" else incomeratio = income/monthprice*100
		if isnull(totalqty) then totalqty = 0
		if isnull(totalprice) then totalprice = 0
	end if

	objrs.close
%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" >
<form>
<input type="hidden" name="contidx" value="<%=contidx%>">
<input type="hidden" name="sidx" value="<%=org_sidx%>">
<!--#include virtual="/od/top.asp" -->
  <table id="Table_01" width="1280" border="1" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top" width="210"><!--#include virtual="/od/left_outdoor_menu.asp"--></td>
      <td height="65" width="1070"><img src="/images/middle_navigater_outdoor.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td  align="left" valign="top"><table width="1030" border="1" cellspacing="0" cellpadding="0">
          <tr>
            <td height="24">&nbsp;</td>
          </tr>
          <tr>
            <td height="19" valign="top" class="navigator" align="right"> 옥외관리 &gt; 계약관리 &gt; <%=title%></td>
          </tr>
          <tr>
            <td height="17"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"><%=title%></span></td>
          </tr>
          <tr>
            <td height="27">&nbsp;</td>
          </tr>
          <tr>
            <td valign="top" class="bdpdd">
			<table width="976" height="35" border="1" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
                  <td width="50%" align="left" background="/images/bg_search.gif">
				            <img src="/images/icon_dot_search.gif" width="4" height="3" align="absmiddle"> 검색년월
				            <select name="cyear">
				<%
					dim intLoop
					for intLoop = 2000 to year(date) + 5
						response.write "<option value='" & intLoop &"' "
						if intLoop = cint(cyear) then response.write " selected "
						response.write ">" & intLoop & "</option>"
					next
				%>
				</select>
				<select name="cmonth">
				<%
					for intLoop = 1 to 12
						response.write "<option value='" & intLoop &"' "
						if intLoop = cint(cmonth) then response.write " selected "
						response.write ">" & intLoop & "</option>"
					next
				%>
				</select> <img src="/images/btn_search.gif" width="39" height="20" align="top" class="styleLink" onclick="go_search();"></td>
                  <td  align="right" valign="bottom" background="/images/bg_search.gif"><img src="/images/btn_contact_edit.gif" width="78" height="18" class="stylelink" onclick="pop_contact_edit();"><img src="/images/btn_contact_extension.gif" width="78" height="18" hspace="10"><img src="/images/btn_contact_cancel.gif" width="78" height="18"> </td>
                  <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table>
              <br>
            <table width="976" border="0" cellspacing="0" cellpadding="0">
			<tr>
				<td colspan="8" bgcolor="#cacaca" height="1"></td>
			</tr>
                <tr>
                  <td class="tdt">계약(매체)명</td>
                  <td colspan="7" class="header3" >&nbsp;<%=title%></td>
                </tr>
			<tr>
				<td colspan="8" bgcolor="#E7E7DE" height="1"></td>
			</tr>
                <tr>
                  <td class="tdhd s2">광고주</td>
                  <td class="tdbd s3">&nbsp;<%=custname%></td>
                  <td class="tdhd s2">사업부</td>
                  <td class="tdbd s3">&nbsp;<%=custname2%></td>
                  <td class="tdhd s2">총수량</td>
                  <td class="tdbd s3">&nbsp;<%=formatnumber(totalqty,0)%>&nbsp;</td>
                  <td class="tdhd s2">&nbsp;</td>
                  <td class="tdbd s3">&nbsp;</td>
                </tr>
			<tr>
				<td colspan="8" bgcolor="#E7E7DE" height="1"></td>
			</tr>
                <tr>
                  <td class="tdhd s2">총광고료</td>
                  <td class="tdbd s3">&nbsp;<%=formatnumber(totalprice, 0)%></td>
                  <td class="tdhd s2">최초계약일</td>
                  <td class="tdbd s3">&nbsp;<%=firstdate%></td>
                  <td class="tdhd s2">시작일</td>
                  <td class="tdbd s3">&nbsp;<%=startdate%></td>
                  <td class="tdhd s2">종료일</td>
                  <td class="tdbd s3">&nbsp;<%=enddate%></td>
                </tr>
			<tr>
				<td colspan="8" bgcolor="#E7E7DE" height="1"></td>
			</tr>
                <tr>
                  <td class="tdhd s2">월광고료</td>
                  <td class="tdbd s3">&nbsp;<%=formatnumber(monthprice,0)%></td>
                  <td class="tdhd s2">월지급액</td>
                  <td class="tdbd s3">&nbsp;<%=formatnumber(expense, 0)%></td>
                  <td class="tdhd s2">내수액</td>
                  <td class="tdbd s3">&nbsp;<%=formatnumber(income, 0)%></td>
                  <td class="tdhd s2">내수율</td>
                  <td class="tdbd s3">&nbsp;<%=formatnumber(incomeratio,2)%></td>
                </tr>
			<tr>
				<td colspan="8" bgcolor="#E7E7DE" height="1"></td>
			</tr>
                <tr>
				  <td height="50" colspan="8" align="right" valign="bottom"><% if not isnull(map) then %><img src="/images/btn_map.gif" width="78" height="18"  vspace="5" class="stylelink"  onClick="pop_medium_map();"><%end if%><img src="/images/btn_account_mng.gif" width="88" height="18" hspace="10"  vspace="5" class="stylelink"  onClick="pop_contact_account_edit(<%=sidx%>);"><img src="/images/btn_comment_mng.gif" width="78" height="18"  vspace="5" class="stylelink"  onClick="pop_contact_comment_edit()"><img src="/images/btn_photo_mng.gif" width="78" height="18"  vspace="5" class="stylelink" hspace="10" onClick="pop_contact_photo_edit(<%=org_sidx%>);"><img src="/images/btn_medium_delete.gif" width="78" height="18"  vspace="5" class="stylelink" onClick="" ><img src="/images/btn_medium_edit.gif" width="78" height="18"  vspace="5" class="stylelink" onClick="pop_contact_medium_edit(<%=org_sidx%>);" hspace="10"><img src="/images/btn_md_reg.gif" width="78" height="18" vspace="5"   class="stylelink" onClick="get_contact_medium_add(<%=contidx%>);"></td>
                </tr>
                <tr>
                  <td height="14" colspan="8">
				  <%
					sql = "select m.sidx, p.mdname, m.side, m.qty, m.locate, m.standard, m.quality, j.thema, d.monthprice, d.expense, c.custname, sj.seqname  from dbo.wb_contact_md m inner join dbo.wb_contact_md_dtl d on (m.contidx = d.contidx and m.sidx = d.sidx)  inner join dbo.sc_cust_temp c on m.custcode = c.custcode  inner join dbo.vw_medium_category p on (m.categoryidx = p.mdidx) inner join dbo.wb_jobcust j on d.jobidx = j.jobidx inner join dbo.sc_jobcust sj on j.seqno = sj.seqno where m.contidx="&contidx&" and cyear = "&cyear&" and cmonth = "&cmonth

					call get_recordset(objrs, sql)

					dim side, categoryname, qty, locate, standard, quality, thema, custname3, seqname

					dim old_sidx : old_sidx = sidx

					if not objrs.eof then
						set sidx = objrs("sidx")
						set categoryname = objrs("mdname")
						set side = objrs("side")
						set qty = objrs("qty")
						set locate = objrs("locate")
						set standard = objrs("standard")
						set quality = objrs("quality")
						set thema = objrs("thema")
						set monthprice = objrs("monthprice")
						set expense = objrs("expense")
						set custname3 = objrs("custname")
						set seqname = objrs("seqname")
					end if
				  %>
				  <table width="976" border="0" cellspacing="1" cellpadding="0">
                    <tr>
                      <td class="hdbd" width="">분류</td>
                      <td class="hdbd" width="">면</td>
                      <td class="hdbd" width="">수량</td>
                      <td class="hdbd"  width="">세부위치</td>
                      <td class="hdbd" width="">규격(M)/재질</td>
                      <td class="hdbd" width="">브랜드</td>
                      <td class="hdbd" width="">소재명</td>
                      <td class="hdbd" width="">월광고액</td>
                      <td class="hdbd" width="">월지급액</td>
                      <td class="hdbd" width="">매체사</td>
                      </tr>
					<%
						dim prev_locate , prev_custname3
						do until objrs.eof
					%>
                    <tr  onClick="get_contact_medium_view('<%=contidx%>', '<%=sidx.value%>', '<%=cyear%>', '<%=cmonth%>');" class="styleLink <%if cstr(org_sidx) = cstr(sidx.value) then Response.write "b"%>">
                      <td class="tbd" ><%=categoryname.value%></td>
                      <td class="tbd"><%=side.value%></td>
                      <td class="tbd"><%=qty.value%></td>
                      <td class="tbd"><%if prev_locate <> locate.value then response.write locate.value%></td>
                      <td class="tbd"><%=standard.value%> <%if not isnull(quality.value) then response.write "(" & quality.value & ")" %></td>
                      <td class="tbd"><%=seqname.value%>&nbsp;</td>
                      <td class="tbd"><%=thema.value%>&nbsp;</td>
                      <td class="tbd" align="right"><%=formatnumber(monthprice.value,0)%></td>
                      <td class="tbd" align="right"><%=formatnumber(expense.value,0)%></td>
                      <td class="tbd"><%if prev_custname3 <> custname3.value then response.write custname3.value%></td>
                      </tr>
				  <%
						prev_locate = locate.value
						prev_custname3 = custname3.value
						objrs.movenext
						loop
						objrs.close

						sql = "select d2.filename from dbo.wb_contact_md m inner join dbo.wb_contact_md_dtl d on m.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on d.idx = a.idx left outer join dbo.wb_contact_photo_mst p on a.idx = p.dtlidx left outer join dbo.wb_contact_photo_dtl d2 on p.idx = d2.mstidx where m.contidx = "&contidx&" and a.cyear = '"&cyear&"' and a.cmonth = '"&cmonth&"' and d2.chk = 1"
						call get_recordset(objrs, sql)

						dim intLoop
						intLoop = 0

				%>
                  </table>
				  </td>
                </tr>
                <tr>
                  <td height="1" colspan="8" align="right"  bgcolor="#ECECEC"></td>
                </tr>
                <tr>
                  <td height="30" colspan="8" style="padding-top:10;">
				  <table width="976" border="0" cellpadding="0" cellspacing="5" bgcolor="#EEEEEF">
                    <tr>
					<% do until objrs.eof or intLoop > 4%>
                      <td   align="center" valign="top">a<img src="<%if objrs("filename") <> "" then response.write "/pds/media/"&objrs("filename") &""" class='stylelink' " else response.write "/images/noimage.gif"%>" width="230" height="132" border="0" onclick="pop_medium_photo('<%=objrs("filename")%>');" >a</td>
					<%
						intLoop = intLoop + 1
						objrs.movenext
						Loop

						objrs.close
						set objrs = nothing
					%>
                    </tr>
                  </table>
				  </td>
                </tr>
                <tr>
                  <td height="30" colspan="8">&nbsp;</td>
                </tr>
                <tr>
                  <td height="30" colspan="8"><table width="976" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td height="1" colspan="2" bgcolor="#E7E7DE"></td>
                      </tr>
                    <tr>
                      <td class="tdhd s2">&nbsp; 매체특성</td>
                      <td class="comment"><%if not isnull(mediummemo) then response.write replace(mediummemo, chr(13)&chr(10), "<br>") %>&nbsp;</td>
                    </tr>
                    <tr>
                      <td height="1" colspan="2" bgcolor="#E7E7DE"></td>
                      </tr>
                    <tr>
                      <td class="tdhd s2">&nbsp; 지역특성</td>
                      <td class="comment"><%if not isnull(regionmemo) then response.write replace(regionmemo, chr(13)&chr(10), "<br>") %>&nbsp;</td>
                    </tr>
                    <tr>
                      <td height="1" colspan="2" bgcolor="#E7E7DE"></td>
                      </tr>
                    <tr>
                      <td class="tdhd s2">&nbsp; 특이사항</td>
                      <td class="comment"><%if not isnull(comment) then response.write replace(comment, chr(13)&chr(10), "<br>") %>&nbsp;</td>
                    </tr>
                    <tr>
                      <td height="1" colspan="2" bgcolor="#E7E7DE"></td>
                      </tr>
                    <tr>
                  </table></td>
                </tr>
            </table></td></tr>
      </table>
	  </td>
    </tr>
  </table>
<!--#include virtual="bottom.asp" -->
</form>
</body>
</html>


<script type="text/javascript" src="/js/calendar.js"></script>
<script type="text/javascript" src="/js/script.js"></script>
<script language="javascript">
<!--

	// 계약 매체 등록
	function get_contact_medium_add(code) {
		var url = "/od/outdoor/pop_contact_medium_reg.asp?contidx="+code;
		var name = "get_contact_medium_add";
		var opt = "width=540, height=483, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	// 계약 매체 수정
	function pop_contact_medium_edit(sidx) {
		var url = "pop_contact_medium_edit.asp?sidx="+sidx+"&contidx=<%=contidx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
		var name = "contact_medium_edit";
		var opt = "width=540, height=483, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	// 계약 매체, 지역 특성 및 특이사항 관리
	function pop_contact_comment_edit() {
		var url = "pop_contact_comment_edit.asp?contidx=<%=contidx%>";
		var name = "pop_comment_edit";
		var opt = "width=540, height=385, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	// 계약 매체별 사진 관리
	function pop_contact_photo_edit(sidx) {
		var url = "pop_contact_photo_edit.asp?contidx=<%=contidx%>&sidx="+sidx+"&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
		var name = "pop_contact_photo_edit";
		var opt = "width=540, height=267, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	// 계약 매체별 광고비 관리
	function pop_contact_account_edit(sidx) {
		var url = "pop_contact_account_edit.asp?contidx=<%=contidx%>&sidx="+sidx;
		var name = "pop_contact_account_edit";
		var opt = "width=540, height=592, resizable=no, scrollbars=yes, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	// 계약 수정
	function pop_contact_edit() {
		var url = "pop_contact_edit.asp?contidx=<%=contidx%>";
		var name = "pop_contact_edit";
		var opt = "width=540, height=577, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	// 계약 검색
	function go_search() {
		var frm = document.forms[0];
		frm.action = "pop_contact_view.asp";
		frm.method = "post";
		frm.submit();
	}

	// 약도 팝업
	function pop_medium_map() {
		var url = "pop_medium_map.asp?sidx=<%=org_sidx%>&contidx=<%=contidx%>";
		var name = "pop_medium_map";
		var opt = "width=650, height=500, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	function pop_medium_photo(photo) {
		if (photo != "") {
			var url = "pop_medium_photo.asp?sidx=<%=org_sidx%>&contidx=<%=contidx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>&photo=" + photo ;
			var name = "pop_medium_photo";
			var opt = "width=668, height=550, resizable=no, scrollbars=yes, status=yes, left=100, top=100";
			window.open(url, name, opt);
		}
	}

	// 계약 정보 보기로 이동하기
	function get_contact_medium_view(contidx, sidx, cyear, cmonth) {
		location.href="/od/outdoor/pop_contact_view.asp?contidx="+contidx+"&sidx="+sidx+"&cyear="+cyear+"&cmonth="+cmonth;
	}

	window.onload = function init() {
		self.focus();
	}


//-->
</script>