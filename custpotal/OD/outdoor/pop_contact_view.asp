<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim idx					' 매체(면) 등록일련번호를 가진다.
	dim contidx				' 현재 선택된 계약의 계약 일련번호
	dim cyear				' 현재 선택된 년도
	dim cmonth			' 현재 선택된 월
	dim objrs				' 레코드셋을 생성하기 위한 레코드셋 변수
	dim sql					' 쿼리 문장을 지정하는 변수
	dim title					' 계약명
	dim firstdate			' 최초계약일
	dim startdate			' 계약 시작일
	dim enddate			' 계약 종료일
	dim regionmemo		' 계약 지역 특성
	dim mediummemo	' 계약 매체 특성
	dim comment			' 계약 변경 사항
	dim canceldate		' 계약 해지 일자
	dim cancel				' 계약 해지 여부 IsCancel
	dim totalprice			' 총광고료
	dim income				' 내수액
	dim incomeratio		' 내수율
	dim custname2		' 사업부서
	dim custname			' 광고주
	dim totalqty			' 현재 선택된 계약 년월에 해당 하는 매체 총갯수
	dim idx2					' 계약 목록에 나타나는 매체(면) 일련번호
	dim sidx					' 계약매체 일련번호
	dim mdname
	dim side
	dim custcode
	dim custcode2
	dim qty, unit, locate, standard, quality, seqname, thema, monthprice2, expense2, medname, map, monthprice, expense, photo_1, photo_2, photo_3, photo_4
	dim isPerform, searchstring
	dim tidx					' 효용성 번호
	dim medclass
	dim validclass
	dim medIdx

	contidx = request("contidx")
	cyear = request("cyear")
	cmonth = request("cmonth")
	custcode = request("custcode")
	custcode2 = request("custcode2")
	searchstring = request("searchstring")
	idx = request("idx")
	sidx = null

	if cyear = "" then cyear = year(date)
	if cmonth = "" then cmonth = month(date)
	if Len(cmonth) = 1 then cmonth = "0"&cmonth

	sql = "select m.contidx, m.title, m.firstdate, m.startdate, m.enddate, m.regionmemo, m.mediummemo, m.comment, m.cancel, m.canceldate, c.custname as custname2, c2.custname from dbo.wb_contact_mst m inner join dbo.sc_cust_temp c on m.custcode = c.custcode inner join dbo.sc_cust_temp c2 on c.highcustcode = c2.custcode where m.contidx = " & contidx

	call get_recordset(objrs, sql)

	if objrs.eof then
		response.write "<script> alert('계약기간이 종료된 년월입니다.'); history.back(); </script>"
		response.end
	end if


	title = objrs("title")									' 계약명
	firstdate = objrs("firstdate")						' 최초 계약일자
	startdate = objrs("startdate")					' 계약 시작일자
	enddate = objrs("enddate")						' 계약 종료일자
	regionmemo = objrs("regionmemo")			' 계약 지역특성
	mediummemo = objrs("mediummemo")		' 계약 매체특성
	comment = objrs("comment")					' 계약 특이사항(변경사항 이력)
	canceldate = objrs("canceldate")				' 계약 해지일자
	cancel = objrs("cancel")							' 계약 해지여부 -> isCancel로 변경하는것이 좋음
	custname2 = objrs("custname2")				' 사업부서
	custname = objrs("custname")					' 광고주명

	objrs.close

	' ********** 계약 등록 후 매체가 등록되기 전에 계약 정보를 확인 하기 위하여 페이지를 오픈했을 경우
	' ********** 계약 리스트에서 선택되어질때 최초로 등록된 매체(면)을 선택하도록 설정한다.
	if idx = "" or isnull(idx) then

		sql = "select isnull(min(a.idx),0) from dbo.wb_contact_md_dtl d inner join dbo.wb_contact_md m on m.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on  d.idx = a.idx where contidx = "&contidx & " and a.cyear = '" & cyear & "' and a.cmonth = '" & cmonth & "' " ' 해당 계약에서 최초로 등록된 매체(면) 정보

		call get_recordset(objrs, sql)

		idx = objrs(0)

		objrs.close

	End if

	' ********** 계약의 총광고료를 계산한다.
	sql = "select isnull(sum(monthprice),0) from dbo.wb_contact_md_dtl d inner join dbo.wb_contact_md m on m.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on d.idx = a.idx where m.contidx = " & contidx

	call get_recordset(objrs, sql)

	totalPrice = objrs(0)

	objrs.close


	' ********** 선택된 년월의 매체(면)의 총갯수를 구한다.
	sql = "select isnull(sum(qty),0) from dbo.wb_contact_md_dtl d inner join dbo.wb_contact_md m on m.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on a.idx = d.idx where m.contidx = " & contidx & " and a.cyear = '" & cyear & "' and a.cmonth = '" & cmonth & "' "

	call get_recordset(objrs, sql)

	totalqty = objrs(0)

	objrs.close

	' ********** 선택된 년월의 매체가 가지는 약도, 사진, 월별 광고료, 지급액, 내수액, 내수율을 구한다.
	sql = "select m.map, m.sidx, a.monthprice, a.expense, a.photo_1, a.photo_2, a.photo_3, a.photo_4, a.isPerform from dbo.wb_contact_md m inner join dbo.wb_contact_md_dtl d on m.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on d.idx = a.idx  where a.idx = " & idx & " and a.cyear = '" & cyear & "' and a.cmonth = '" & cmonth & "' "

	call get_recordset(objrs, sql)

	if not objrs.eof then
		map = objrs("map")
		sidx = objrs("sidx")
		monthprice = objrs("monthprice")
		expense = objrs("expense")
		photo_1 = objrs("photo_1")
		photo_2 = objrs("photo_2")
		photo_3 = objrs("photo_3")
		photo_4 = objrs("photo_4")
		isPerform = objrs("isPerform")

		income = monthprice - expense
		if monthprice <> 0 then incomeratio = income/monthprice*100 else incomeratio = "0.00" end if
	else
		monthprice =0
		expense = 0
		income = 0
		incomeRatio = 0
	end if

	objrs.close

	if isnull(photo_1) and isnull(photo_2) and isnull(photo_3) and isnull(photo_4)  then
		sql = "select max(photo_1), max(photo_2), max(photo_3), max(photo_4) from dbo.wb_contact_md_dtl_account where idx = "& idx &" and cyear+cmonth <= '" & cyear&cmonth &"' "
		call get_recordset(objrs, sql)
		if not objrs.eof then
			photo_1 = objrs(0)
			photo_2 = objrs(1)
			photo_3 = objrs(2)
			photo_4 = objrs(3)
		end if
		objrs.close

	end if

	' 매체등급, 효용성 등급을 조회
	sql = "select m.tidx, c.class, t.validclass, c.mdidx  from dbo.wb_validation_mst m inner join dbo.wb_validation_class c on m.tidx = c.tidx inner join dbo.wb_validation_tool t on m.tidx = t.tidx where  m.contidx = " & contidx
	call get_recordset(objrs, sql)

	if objrs.eof then
		tidx = null
		medclass = null
		validclass = null
		medIdx = null
	else
		tidx = objrs(0)
		medclass = objrs(1)
		validclass = objrs(2)
		medIdx = objrs(3)
	end if

	objrs.close

%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"   oncontextmenu="return false">
<form>
<INPUT TYPE="hidden" NAME="contidx" value="<%=contidx%>">
<table width="1240" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td height="24"><img src="/images/pop_top.gif" width="1240" height="60" align="absmiddle"></td>
  </tr>
  <tr>
    <td height="24">&nbsp;</td>
  </tr>
  <tr>
    <td height="17"  align="center"><table border="0" cellpadding="0" cellspacing="0" width="976">
    <tr>
		<td><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"><%=title%></span></td>
    </tr>
    </table></td>
  </tr>
  <tr>
    <td height="27">&nbsp;</td>
  </tr>
  <tr>
    <td valign="top">
	<table width="976" height="35" border="0" cellpadding="0" cellspacing="0" align="center">
      <tr>
        <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
        <td width="50%" align="left" background="/images/bg_search.gif"><img src="/images/icon_dot_search.gif" width="4" height="3" align="absmiddle">           <select name="cyear">
                <%
					dim intLoop
					for intLoop = 2005 to year(date) + 5
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
              </select>
              <img src="/images/btn_search.gif" width="39" height="20" align="top" class="styleLink" onClick="go_search();"></td>
        <td  align="right" valign="bottom" background="/images/bg_search.gif"><%if not cancel then %>
		<img src="/images/btn_md_validation.gif" width="78" height="18"  hspace="10" border="0" class="stylelink" onClick="pop_medium_validation(<%=contidx%>, '<%=custcode%>');" hspace="10"><img src="/images/btn_contact_edit.gif" width="78" height="18" class="stylelink" onClick="pop_contact_edit();" alt="계약 기간(시작, 종료), 광고주, 사업부, 입력 정보를 수정하시면 버튼을 누르세요. "  ><img src="/images/btn_contact_extension.gif" width="78" height="18" class="stylelink" onclick="pop_contact_extention();" alt="현재 등록된 계약과 동일(금액, 매체사)한 계약을 연장하시려면 누르세요. 금액 또는 매체의 변경이 있는 경우에는 새로 계약을 등록하세요"  hspace="10"><img src="/images/btn_contact_cancel.gif" width="78" height="18"  class="stylelink" onclick="set_contact_cancel();" alt="계약을 기간 중간에 해지할 경우에 계약에 대한 데이터는 삭제되지 않습니다."><img src="/images/btn_contact_delete.gif" width="78" height="18" class="stylelink" onclick="set_contact_delete();" alt="계약정보가 시스템에서 완전히 삭제 됩니다."  hspace="10" ><%end if%><img src="/images/btn_close.gif" width="57" height="18"  style="cursor:hand" onClick="set_close();" > </td>
        <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
      </tr>
    </table>
        <br>
        <table width="976" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td colspan="8" bgcolor="#cacaca" height="1"></td>
          </tr>
          <tr>
            <td class="tdt">계약(매체)명</td>
            <td colspan="7" class="header3" style="padding-left:10px;"><%=title%> <%if cancel then response.write "계약해지 (" & canceldate & ")" %></td>
          </tr>
          <tr>
            <td colspan="8" bgcolor="#E7E7DE" height="1"></td>
          </tr>
          <tr>
            <td class="tdhd s2">광고주</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=custname%></td>
            <td class="tdhd s2">사업부</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=custname2%></td>
            <td class="tdhd s2">총수량</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=formatnumber(totalqty,0)%></td>
            <td class="tdbd " colspan="2"><%if not isnull(medclass) then response.write "매체등급 " & medclass%> <% if not isnull(validclass) then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;효용성등급 " & validclass%></td>
          </tr>
          <tr>
            <td colspan="8" bgcolor="#E7E7DE" height="1"></td>
          </tr>
          <tr>
            <td class="tdhd s2">총광고료</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=formatnumber(totalprice, 0)%></td>
            <td class="tdhd s2">최초계약일</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=firstdate%></td>
            <td class="tdhd s2">시작일</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=startdate%></td>
            <td class="tdhd s2">종료일</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=enddate%></td>
          </tr>
          <tr>
            <td colspan="8" bgcolor="#E7E7DE" height="1"></td>
          </tr>
          <tr>
            <td class="tdhd s2">월광고료</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=formatnumber(monthprice,0)%></td>
            <td class="tdhd s2">월지급액</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=formatnumber(expense, 0)%></td>
            <td class="tdhd s2">내수액</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=formatnumber(income, 0)%></td>
            <td class="tdhd s2">내수율</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=formatnumber(incomeratio,2)%></td>
          </tr>
          <tr>
            <td colspan="8" bgcolor="#E7E7DE" height="1"></td>
          </tr>
          <tr>
            <td height="50" colspan="8" align="right" valign="bottom"><img src="/images/btn_side_reg.gif" width="78" height="18"  vspace="5" class="stylelink" onClick="pop_contact_medium_side_reg(<%=sidx%>);"><img src="/images/btn_account_mng.gif" width="88" height="18" vspace="5" class="stylelink"  onClick="pop_contact_account_edit(<%=idx%>);" hspace="10" ><img src="/images/btn_photo_reg.gif" width="78" height="18"  vspace="5" class="stylelink" onClick="pop_contact_photo_edit(<%=idx%>);"><img src="/images/btn_medium_edit.gif" width="78" height="18"  vspace="5" class="stylelink" onClick="pop_contact_medium_edit(<%=idx%>);" hspace="10"><img src="/images/btn_md_reg.gif" width="78" height="18" vspace="5"   class="stylelink" onClick="get_contact_medium_add(<%=contidx%>);"></td>
          </tr>
          <tr>
            <td></td>
          </tr>
		 </table>

		<%
' ********** 계약된 매체(면) 리스트를 가져온다

			dim tmpMdName
			dim tmpLocate
			dim tmpStandard
			dim tmpQuality
			dim tmpBrand
			dim tmpMedName
			sql = "select k.mdname, a.idx, m.sidx, d.side, a.qty, m.locate, d.standard, d.quality, j2.seqname, j.thema, a.monthprice, a.expense, c.custname as medname from dbo.wb_contact_md m inner join dbo.wb_contact_md_dtl d on m.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on d.idx = a.idx inner join dbo.vw_medium_category k on m.categoryidx = k.mdidx left outer  join dbo.wb_jobcust j on j.jobidx = a.jobidx left outer join dbo.sc_jobcust j2 on j.seqno = j2.seqno inner join dbo.sc_cust_temp c on m.medcode = c.custcode where m.contidx = " & contidx & " and a.cyear = '" & cyear & "' and a.cmonth = '" & cmonth &"' order by m.sidx, a.idx"

			call get_recordset(objrs, sql)
'
			if not objrs.eof then
				set idx2 = objrs("idx")
				set sidx = objrs("sidx")
				set mdname = objrs("mdname")
				set side = objrs("side")
				set qty = objrs("qty")
				set locate = objrs("locate")
				set standard = objrs("standard")
				set quality = objrs("quality")
				set seqname = objrs("seqname")
				set thema = objrs("thema")
				set monthprice2 = objrs("monthprice")
				set expense2 = objrs("expense")
				set medname = objrs("medname")
			end if

	%>
   <table width="1240" border="0" cellspacing="1" cellpadding="0" >
     <tr>
       <td class="hdbd" width="100">매체분류</td>
       <td class="hdbd" width="30">면</td>
       <td class="hdbd" width="70">수량</td>
       <td class="hdbd" width="200">설치위치</td>
       <td class="hdbd"  width="200">규격/재질</td>
       <td class="hdbd" width="100">브랜드</td>
       <td class="hdbd" width="100">소재</td>
       <td class="hdbd" width="100">월광고료</td>
       <td class="hdbd" width="100">월지급액</td>
<!--        <td class="hdbd" width="100">내수액</td>
       <td class="hdbd" width="70">내수율</td> -->
       <td class="hdbd" width="130">매체사 </td>
       <td class="hdbd"  width="15"><IMG SRC="/images/btn_comment-delete.gif" WIDTH="9" HEIGHT="9" BORDER="0" ALT="&nbsp;" > </td>
     </tr>
     <%
		do until objrs.eof
			'income = monthprice2-expense2
			'if income = 0 then incomeRatio = 0 else incomeRatio = income / monthprice2 * 100
	  %>
     <tr bgcolor="<%if  int(idx) = int(idx2) then response.write "#FFC1C1" else response.write "#FFFFFF"%>">
       <td class="tbd styleLink"  onClick="get_contact_medium_view(<%=idx2%>);" align="left"> <%if tmpMdName <> mdname then response.write mdname end if%></td>
       <td class="tbd styleLink"  onClick="get_contact_medium_view(<%=idx2%>);"><%=side%></td>
       <td class="tbd styleLink"  onClick="get_contact_medium_view(<%=idx2%>);"><%=qty%></td>
       <td class="tbd styleLink"  onClick="get_contact_medium_view(<%=idx2%>);" ><%if tmpLocate <> locate then response.write locate end if%></td>
       <td class="tbd styleLink"  onClick="get_contact_medium_view(<%=idx2%>);"> <%=standard%> <%if not isnull(quality) then response.write " / " & quality  %></td>
       <td class="tbd styleLink" onClick="get_contact_medium_view(<%=idx2%>);"><% if seqname <> tmpBrand then response.write seqname end if%>&nbsp;</td>
       <td class="tbd styleLink" onClick="get_contact_medium_view(<%=idx2%>);"><%=thema%>&nbsp;</td>
       <td class="tbd" align="right"><%=formatnumber(monthprice2,0)%></td>
       <td class="tbd" align="right"><%=formatnumber(expense2,0)%></td>
<!--        <td class="tbd" align="right"><%'=formatnumber(income,0)%></td>
       <td class="tbd" align="right"><%'=formatnumber(incomeRatio,2)%></td> -->
       <td class="tbd styleLink" onClick="get_contact_medium_view(<%=idx2%>);"><%=medname%></td>
       <td class="tbd" width="15" bgcolor="#FFFFFF"><IMG SRC="/images/btn_comment-delete.gif" WIDTH="9" HEIGHT="9" BORDER="0" ALT="선택한 매체정보를 삭제합니다." onclick="get_medium_delete(<%=idx2%>);" valign="middle" vspace="3" align="absmiddle" class="stylelink"></td>
     </tr>
	 <tr>
		<td height="1" colspan="13" align="right"  bgcolor="#ECECEC"></td>
	 </tr>
     <%
			tmpMdName = mdname
			tmpLocate = locate
			tmpStandard = standard
			tmpBrand = seqname
			tmpQuality = quality
			tmpMedName = medname
			objrs.movenext
		loop

		objrs.close
		set objrs = nothing
	  %>
 </table>
 <p>
        <table width="976" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td height="30" colspan="8" style="padding-top:10;"><table width="976" border="0" cellpadding="0" cellspacing="5" bgcolor="#EEEEEF">
                <tr>
                  <td   align="center" valign="top"><img src="<%if photo_1 <> "" then response.write "/pds/media/"&photo_1& """ class='stylelink' " else response.write "/images/noimage.gif"%>" width="230" height="152" border="0" onClick="pop_medium_photo('<%=photo_1%>');" ></td>
                  <td  align="center" valign="top"><img src="<%if photo_2 <> "" then response.write "/pds/media/"&photo_2& """   class='stylelink' "else response.write "/images/noimage.gif"%>" width="230" height="152" border="0" onClick="pop_medium_photo('<%=photo_2%>');"></td>
                  <td   align="center" valign="top"><img src="<%if photo_3 <> "" then response.write "/pds/media/"&photo_3 &""" class='stylelink' "else response.write "/images/noimage.gif"%>" width="230" height="152" border="0" onClick="pop_medium_photo('<%=photo_3%>');"></td>
                  <td  align="center" valign="top"><img src="<%if photo_4 <> "" then response.write "/pds/media/"&photo_4 &""" class='stylelink' "else response.write "/images/noimage.gif"%>" width="230" height="152" border="0" onClick="pop_medium_photo('<%=photo_4%>');"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="30" colspan="8">&nbsp;</td>
          </tr>
          <tr>
            <td height="30" colspan="8"><table width="976" border="0" cellspacing="0" cellpadding="0" align="center">
                <tr>
                  <td height="1" bgcolor="#E7E7DE"></td>
                  <td height="1" bgcolor="#E7E7DE"></td>
                  <td width="250" rowspan="7" align="center" valign="middle" bgcolor="#E7E7DE"><img src="<%if map <> "" then response.write "/map/"&map& """ class='stylelink' " else response.write "/images/noimage.gif"%>" width="250" height="165" border="0" onClick="pop_medium_map('<%=map%>');" ></td>
                </tr>
                <tr>
                  <td class="tdhd s2">&nbsp; 매체특성</td>
                  <td class="comment"><%if not isnull(mediummemo) then response.write replace(mediummemo, chr(13)&chr(10), "<br>") %>
                    &nbsp;</td>
                </tr>
                <tr>
                  <td height="1" bgcolor="#E7E7DE"></td>
                  <td height="1" bgcolor="#E7E7DE"></td>
                </tr>
                <tr>
                  <td class="tdhd s2">&nbsp; 지역특성</td>
                  <td class="comment"><%if not isnull(regionmemo) then response.write replace(regionmemo, chr(13)&chr(10), "<br>") %>
                    &nbsp;</td>
                </tr>
                <tr>
                  <td height="1" bgcolor="#E7E7DE"></td>
                  <td height="1" bgcolor="#E7E7DE"></td>
                </tr>
                <tr>
                  <td class="tdhd s2">&nbsp; 특이사항</td>
                  <td class="comment"><%if not isnull(comment) then response.write replace(comment, chr(13)&chr(10), "<br>") %>
                    &nbsp;</td>
                </tr>
                <tr>
                  <td height="1" bgcolor="#E7E7DE"></td>
                  <td height="1" bgcolor="#E7E7DE"></td>
                </tr>
            </table></td>
          </tr>
      </table></td>
  </tr>
  <tr>
    <td height="24">&nbsp;</td>
  </tr>
  <tr>
    <td height="24"><img src="/images/pop_bottom.gif" width="1240" height="71" align="absmiddle"></td>
  </tr>
</table>
</form>
</body>
</html>


<script language="javascript">
<!--

	//  ******************************************************************************  계약 매체 면 추가
	//  **************  sidx (계약 매체별 매체등록 일련번호)
	//  ******************************************************************************
	function pop_contact_medium_side_reg(sidx) {
		if ("<%=isPerform %>" == "True") {
			alert("광고비가 정산된 달은 매체를 추가할 수 없습니다.");
			return false ;
		}
		if (!sidx) {
			alert("면을 추가할 매체를 선택하세요");
			return false ;
		}
		var url = "/od/outdoor/pop_side_reg.asp?sidx="+sidx+"&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
		var name = "pop_side_reg";
		var opt = "width=540, height=480, resizable=no, top=100, left=660;"
		window.open(url, name, opt);
	}

	//  ******************************************************************************  계약 매체 등록
	//  **************  contidx (계약번호)
	//  ******************************************************************************
	function get_contact_medium_add(contidx) {
		if ("<%=isPerform %>" == "True") {
			alert("광고비가 정산된 달은 매체를 등록할 수 없습니다.");
			return false ;
		}
		var bln = "<%=cint(cancel)%>";
		if (bln == -1) {
			alert("해지된 계약은 매체를 등록할 수 없습니다.");
			return false ;
		}
		var url = "/od/outdoor/pop_contact_medium_reg.asp?contidx="+contidx+"&cyear=<%=cyear%>&cmonth=<%=cmonth%>" ;
		var name = "get_contact_medium_add";
		var opt = "width=540, height=668, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	//  ******************************************************************************  계약 매체 수정
	//  **************  idx 계약된 매체의 년,월별 계약정보 ()
	//  ******************************************************************************
	function pop_contact_medium_edit(idx) {
		if ("<%=isPerform %>" == "True") {
			alert("광고비가 정산된 달은 매체를 수정할 수 없습니다.");
			return false ;
		}
		if (idx == "")  {
			alert("수정할 매체를 먼저 선택하세요.");
			return false ;
		}
		var url = "pop_contact_medium_edit.asp?idx="+idx+"&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
		var name = "contact_medium_edit";
		var opt = "width=540, height=668, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	//  ******************************************************************************  매체별 사진 관리
	//  **************  idx 계약된 매체의 년,월별 계약정보 ()
	//  ******************************************************************************
	function pop_contact_photo_edit(idx) {
		if (idx == "")  {
			alert("매체를 먼저 등록하신 후 사진을 등록하세요");
			return false ;
		}
			var url = "pop_contact_photo_edit.asp?idx="+idx+"&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
			var name = "pop_contact_photo_edit";
			var opt = "width=540, height=558, resizable=no, scrollbars=no, status=yes, left=100, top=100";
			window.open(url, name, opt);
	}


	//  ******************************************************************************  계약 광고비 전체 관리
	//  **************  idx 계약된 매체의 년,월별 계약정보 ()
	//  ******************************************************************************
	function pop_contact_account_edit(idx) {
		if (idx == "")  {
			alert("광고비를 확인할 매체를 선택하세요.");
			return false ;
		}
		var url = "pop_contact_account_edit.asp?idx="+idx+"&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
		var name = "pop_contact_account_edit";
		var opt = "width=540, height=592, resizable=no, scrollbars=yes, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	//계약 매체 삭제하기
	function get_medium_delete(idx) {
		if ("<%=isPerform %>" == "True") {
			alert("광고비가 정산된 달은 매체를 삭제할 수 없습니다.");
			return false ;
		}
		if (confirm("계약정보에서 선택된 매체(면) 정보가 삭제됩니다.\n\n계약 매체(면)를 삭제하시겠습니까?")) {
			location.href = "contact_medium_delete_proc.asp?idx="+idx+"&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
		} else {
			return false ;
		}
	}

	// 계약 수정
	function pop_contact_edit() {
		var url = "pop_contact_edit.asp?contidx=<%=contidx%>";
		var name = "pop_contact_edit";
		var opt = "width=540, height=577, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	// 계약삭제
	function set_contact_delete() {
		if (confirm("계약에 해당하는 모든정보가 모두 삭제됩니다.\n\n삭제하시겠습니까?")) {
		location.href = "contact_delete_proc.asp?contidx=<%=contidx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>&searchstring=<%=searchstring%>&custcode=<%=custcode%>&custcode2=<%=custcode2%>";
		}
	}

	//계약연장
	function pop_contact_extention() {
		var url = "pop_contact_extention.asp?contidx=<%=contidx%>";
		var name = "pop_contact_edit";
		var opt = "width=540, height=577, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}
	// 계약 해지
	function set_contact_cancel() {
		if (confirm("계약을 해지하시겠습니까?")) {
			location.href = "contact_cancel.asp?contidx=<%=contidx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
		}
	}
	// 계약 검색
	function go_search() {
		var frm = document.forms[0];
		frm.action = "pop_contact_view.asp";
		frm.method = "post";
		frm.submit();
	}


	//  ******************************************************************************  매체 사진 보기
	//  **************  idx 해당년월 매체번호, photo 사진파일명
	//  ******************************************************************************
	function pop_medium_photo(photo) {
		if (photo != "") {
			var url = "pop_medium_photo.asp?photo=" + photo+"&idx=<%=idx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>" ;
			var name = "pop_medium_photo";
			var opt = "width=668, height=550, resizable=no, scrollbars=yes, status=yes, left=100, top=100";
			window.open(url, name, opt);
		}
	}

	function pop_medium_map(photo) {
		if (photo != "") {
			var url = "pop_medium_map.asp?photo=" + photo ;
			var name = "pop_medium_photo";
			var opt = "width=668, height=550, resizable=no, scrollbars=yes, status=yes, left=100, top=100";
			window.open(url, name, opt);
		}
	}

	function pop_validation_reg() {
			var url = "pop_validation_board.asp?photo=" + photo ;
			var name = "pop_medium_photo";
			var opt = "width=668, height=550, resizable=no, scrollbars=yes, status=yes, left=100, top=100";
			window.open(url, name, opt);
	}

	//  ******************************************************************************  매체 정보 확인 하기

	//  **************  contidx (계약번호)

	//  ******************************************************************************
	function get_contact_medium_view(idx) {
		location.href="/od/outdoor/pop_contact_view.asp?contidx=<%=contidx%>&idx="+idx+"&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
	}


	function pop_medium_validation(idx) {
		var url ;
		<% if  isnull(tidx) then %>
		url = "validation_led.asp?tidx=<%=tidx%>&contidx=<%=contidx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";						//LED
		<% else %>
		<% select case medIdx%>
		<% case "L" %>
		url = "pop_validation_led.asp?tidx=<%=tidx%>&contidx=<%=contidx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";						//LED
		<% case "B"%>
		url = "pop_validation_board.asp?tidx=<%=tidx%>&contidx=<%=contidx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";						//LED
		<% case "N" %>
		url = "pop_validation_neon.asp?tidx=<%=tidx%>&contidx=<%=contidx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";						//LED
		<% end select%>
		<% end if %>
		var name = "pop_medium_validation";
		var opt = "width=893, height=800, resizable=no, scrollbars=yes, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function set_close() {
		this.close();
	}

	window.onload = function () {
		self.focus();
	}


//-->
</script>