<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim idx					' ��ü(��) ����Ϸù�ȣ�� ������.
	dim contidx				' ���� ���õ� ����� ��� �Ϸù�ȣ
	dim cyear				' ���� ���õ� �⵵
	dim cmonth			' ���� ���õ� ��
	dim objrs				' ���ڵ���� �����ϱ� ���� ���ڵ�� ����
	dim sql					' ���� ������ �����ϴ� ����
	dim title					' ����
	dim firstdate			' ���ʰ����
	dim startdate			' ��� ������
	dim enddate			' ��� ������
	dim regionmemo		' ��� ���� Ư��
	dim mediummemo	' ��� ��ü Ư��
	dim comment			' ��� ���� ����
	dim canceldate		' ��� ���� ����
	dim cancel				' ��� ���� ���� IsCancel
	dim totalprice			' �ѱ����
	dim income				' ������
	dim incomeratio		' ������
	dim custname2		' ����μ�
	dim custname			' ������
	dim totalqty			' ���� ���õ� ��� ����� �ش� �ϴ� ��ü �Ѱ���
	dim idx2					' ��� ��Ͽ� ��Ÿ���� ��ü(��) �Ϸù�ȣ
	dim sidx					' ����ü �Ϸù�ȣ
	dim mdname
	dim side
	dim custcode
	dim custcode2
	dim qty, unit, locate, standard, quality, seqname, thema, monthprice2, expense2, medname, map, monthprice, expense, photo_1, photo_2, photo_3, photo_4
	dim isPerform, searchstring
	dim tidx					' ȿ�뼺 ��ȣ
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
		response.write "<script> alert('���Ⱓ�� ����� ����Դϴ�.'); history.back(); </script>"
		response.end
	end if


	title = objrs("title")									' ����
	firstdate = objrs("firstdate")						' ���� �������
	startdate = objrs("startdate")					' ��� ��������
	enddate = objrs("enddate")						' ��� ��������
	regionmemo = objrs("regionmemo")			' ��� ����Ư��
	mediummemo = objrs("mediummemo")		' ��� ��üƯ��
	comment = objrs("comment")					' ��� Ư�̻���(������� �̷�)
	canceldate = objrs("canceldate")				' ��� ��������
	cancel = objrs("cancel")							' ��� �������� -> isCancel�� �����ϴ°��� ����
	custname2 = objrs("custname2")				' ����μ�
	custname = objrs("custname")					' �����ָ�

	objrs.close

	' ********** ��� ��� �� ��ü�� ��ϵǱ� ���� ��� ������ Ȯ�� �ϱ� ���Ͽ� �������� �������� ���
	' ********** ��� ����Ʈ���� ���õǾ����� ���ʷ� ��ϵ� ��ü(��)�� �����ϵ��� �����Ѵ�.
	if idx = "" or isnull(idx) then

		sql = "select isnull(min(a.idx),0) from dbo.wb_contact_md_dtl d inner join dbo.wb_contact_md m on m.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on  d.idx = a.idx where contidx = "&contidx & " and a.cyear = '" & cyear & "' and a.cmonth = '" & cmonth & "' " ' �ش� ��࿡�� ���ʷ� ��ϵ� ��ü(��) ����

		call get_recordset(objrs, sql)

		idx = objrs(0)

		objrs.close

	End if

	' ********** ����� �ѱ���Ḧ ����Ѵ�.
	sql = "select isnull(sum(monthprice),0) from dbo.wb_contact_md_dtl d inner join dbo.wb_contact_md m on m.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on d.idx = a.idx where m.contidx = " & contidx

	call get_recordset(objrs, sql)

	totalPrice = objrs(0)

	objrs.close


	' ********** ���õ� ����� ��ü(��)�� �Ѱ����� ���Ѵ�.
	sql = "select isnull(sum(qty),0) from dbo.wb_contact_md_dtl d inner join dbo.wb_contact_md m on m.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on a.idx = d.idx where m.contidx = " & contidx & " and a.cyear = '" & cyear & "' and a.cmonth = '" & cmonth & "' "

	call get_recordset(objrs, sql)

	totalqty = objrs(0)

	objrs.close

	' ********** ���õ� ����� ��ü�� ������ �൵, ����, ���� �����, ���޾�, ������, �������� ���Ѵ�.
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

	' ��ü���, ȿ�뼺 ����� ��ȸ
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
<title>�Ƣ� SK M&C | Media Management System �Ƣ� </title>
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
		<img src="/images/btn_md_validation.gif" width="78" height="18"  hspace="10" border="0" class="stylelink" onClick="pop_medium_validation(<%=contidx%>, '<%=custcode%>');" hspace="10"><img src="/images/btn_contact_edit.gif" width="78" height="18" class="stylelink" onClick="pop_contact_edit();" alt="��� �Ⱓ(����, ����), ������, �����, �Է� ������ �����Ͻø� ��ư�� ��������. "  ><img src="/images/btn_contact_extension.gif" width="78" height="18" class="stylelink" onclick="pop_contact_extention();" alt="���� ��ϵ� ���� ����(�ݾ�, ��ü��)�� ����� �����Ͻ÷��� ��������. �ݾ� �Ǵ� ��ü�� ������ �ִ� ��쿡�� ���� ����� ����ϼ���"  hspace="10"><img src="/images/btn_contact_cancel.gif" width="78" height="18"  class="stylelink" onclick="set_contact_cancel();" alt="����� �Ⱓ �߰��� ������ ��쿡 ��࿡ ���� �����ʹ� �������� �ʽ��ϴ�."><img src="/images/btn_contact_delete.gif" width="78" height="18" class="stylelink" onclick="set_contact_delete();" alt="��������� �ý��ۿ��� ������ ���� �˴ϴ�."  hspace="10" ><%end if%><img src="/images/btn_close.gif" width="57" height="18"  style="cursor:hand" onClick="set_close();" > </td>
        <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
      </tr>
    </table>
        <br>
        <table width="976" border="0" cellspacing="0" cellpadding="0" align="center">
          <tr>
            <td colspan="8" bgcolor="#cacaca" height="1"></td>
          </tr>
          <tr>
            <td class="tdt">���(��ü)��</td>
            <td colspan="7" class="header3" style="padding-left:10px;"><%=title%> <%if cancel then response.write "������� (" & canceldate & ")" %></td>
          </tr>
          <tr>
            <td colspan="8" bgcolor="#E7E7DE" height="1"></td>
          </tr>
          <tr>
            <td class="tdhd s2">������</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=custname%></td>
            <td class="tdhd s2">�����</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=custname2%></td>
            <td class="tdhd s2">�Ѽ���</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=formatnumber(totalqty,0)%></td>
            <td class="tdbd " colspan="2"><%if not isnull(medclass) then response.write "��ü��� " & medclass%> <% if not isnull(validclass) then response.write "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;ȿ�뼺��� " & validclass%></td>
          </tr>
          <tr>
            <td colspan="8" bgcolor="#E7E7DE" height="1"></td>
          </tr>
          <tr>
            <td class="tdhd s2">�ѱ����</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=formatnumber(totalprice, 0)%></td>
            <td class="tdhd s2">���ʰ����</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=firstdate%></td>
            <td class="tdhd s2">������</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=startdate%></td>
            <td class="tdhd s2">������</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=enddate%></td>
          </tr>
          <tr>
            <td colspan="8" bgcolor="#E7E7DE" height="1"></td>
          </tr>
          <tr>
            <td class="tdhd s2">�������</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=formatnumber(monthprice,0)%></td>
            <td class="tdhd s2">�����޾�</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=formatnumber(expense, 0)%></td>
            <td class="tdhd s2">������</td>
            <td class="tdbd s3" style="padding-left:10px;"><%=formatnumber(income, 0)%></td>
            <td class="tdhd s2">������</td>
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
' ********** ���� ��ü(��) ����Ʈ�� �����´�

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
       <td class="hdbd" width="100">��ü�з�</td>
       <td class="hdbd" width="30">��</td>
       <td class="hdbd" width="70">����</td>
       <td class="hdbd" width="200">��ġ��ġ</td>
       <td class="hdbd"  width="200">�԰�/����</td>
       <td class="hdbd" width="100">�귣��</td>
       <td class="hdbd" width="100">����</td>
       <td class="hdbd" width="100">�������</td>
       <td class="hdbd" width="100">�����޾�</td>
<!--        <td class="hdbd" width="100">������</td>
       <td class="hdbd" width="70">������</td> -->
       <td class="hdbd" width="130">��ü�� </td>
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
       <td class="tbd" width="15" bgcolor="#FFFFFF"><IMG SRC="/images/btn_comment-delete.gif" WIDTH="9" HEIGHT="9" BORDER="0" ALT="������ ��ü������ �����մϴ�." onclick="get_medium_delete(<%=idx2%>);" valign="middle" vspace="3" align="absmiddle" class="stylelink"></td>
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
                  <td class="tdhd s2">&nbsp; ��üƯ��</td>
                  <td class="comment"><%if not isnull(mediummemo) then response.write replace(mediummemo, chr(13)&chr(10), "<br>") %>
                    &nbsp;</td>
                </tr>
                <tr>
                  <td height="1" bgcolor="#E7E7DE"></td>
                  <td height="1" bgcolor="#E7E7DE"></td>
                </tr>
                <tr>
                  <td class="tdhd s2">&nbsp; ����Ư��</td>
                  <td class="comment"><%if not isnull(regionmemo) then response.write replace(regionmemo, chr(13)&chr(10), "<br>") %>
                    &nbsp;</td>
                </tr>
                <tr>
                  <td height="1" bgcolor="#E7E7DE"></td>
                  <td height="1" bgcolor="#E7E7DE"></td>
                </tr>
                <tr>
                  <td class="tdhd s2">&nbsp; Ư�̻���</td>
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

	//  ******************************************************************************  ��� ��ü �� �߰�
	//  **************  sidx (��� ��ü�� ��ü��� �Ϸù�ȣ)
	//  ******************************************************************************
	function pop_contact_medium_side_reg(sidx) {
		if ("<%=isPerform %>" == "True") {
			alert("����� ����� ���� ��ü�� �߰��� �� �����ϴ�.");
			return false ;
		}
		if (!sidx) {
			alert("���� �߰��� ��ü�� �����ϼ���");
			return false ;
		}
		var url = "/od/outdoor/pop_side_reg.asp?sidx="+sidx+"&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
		var name = "pop_side_reg";
		var opt = "width=540, height=480, resizable=no, top=100, left=660;"
		window.open(url, name, opt);
	}

	//  ******************************************************************************  ��� ��ü ���
	//  **************  contidx (����ȣ)
	//  ******************************************************************************
	function get_contact_medium_add(contidx) {
		if ("<%=isPerform %>" == "True") {
			alert("����� ����� ���� ��ü�� ����� �� �����ϴ�.");
			return false ;
		}
		var bln = "<%=cint(cancel)%>";
		if (bln == -1) {
			alert("������ ����� ��ü�� ����� �� �����ϴ�.");
			return false ;
		}
		var url = "/od/outdoor/pop_contact_medium_reg.asp?contidx="+contidx+"&cyear=<%=cyear%>&cmonth=<%=cmonth%>" ;
		var name = "get_contact_medium_add";
		var opt = "width=540, height=668, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	//  ******************************************************************************  ��� ��ü ����
	//  **************  idx ���� ��ü�� ��,���� ������� ()
	//  ******************************************************************************
	function pop_contact_medium_edit(idx) {
		if ("<%=isPerform %>" == "True") {
			alert("����� ����� ���� ��ü�� ������ �� �����ϴ�.");
			return false ;
		}
		if (idx == "")  {
			alert("������ ��ü�� ���� �����ϼ���.");
			return false ;
		}
		var url = "pop_contact_medium_edit.asp?idx="+idx+"&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
		var name = "contact_medium_edit";
		var opt = "width=540, height=668, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	//  ******************************************************************************  ��ü�� ���� ����
	//  **************  idx ���� ��ü�� ��,���� ������� ()
	//  ******************************************************************************
	function pop_contact_photo_edit(idx) {
		if (idx == "")  {
			alert("��ü�� ���� ����Ͻ� �� ������ ����ϼ���");
			return false ;
		}
			var url = "pop_contact_photo_edit.asp?idx="+idx+"&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
			var name = "pop_contact_photo_edit";
			var opt = "width=540, height=558, resizable=no, scrollbars=no, status=yes, left=100, top=100";
			window.open(url, name, opt);
	}


	//  ******************************************************************************  ��� ����� ��ü ����
	//  **************  idx ���� ��ü�� ��,���� ������� ()
	//  ******************************************************************************
	function pop_contact_account_edit(idx) {
		if (idx == "")  {
			alert("����� Ȯ���� ��ü�� �����ϼ���.");
			return false ;
		}
		var url = "pop_contact_account_edit.asp?idx="+idx+"&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
		var name = "pop_contact_account_edit";
		var opt = "width=540, height=592, resizable=no, scrollbars=yes, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	//��� ��ü �����ϱ�
	function get_medium_delete(idx) {
		if ("<%=isPerform %>" == "True") {
			alert("����� ����� ���� ��ü�� ������ �� �����ϴ�.");
			return false ;
		}
		if (confirm("����������� ���õ� ��ü(��) ������ �����˴ϴ�.\n\n��� ��ü(��)�� �����Ͻðڽ��ϱ�?")) {
			location.href = "contact_medium_delete_proc.asp?idx="+idx+"&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
		} else {
			return false ;
		}
	}

	// ��� ����
	function pop_contact_edit() {
		var url = "pop_contact_edit.asp?contidx=<%=contidx%>";
		var name = "pop_contact_edit";
		var opt = "width=540, height=577, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	// ������
	function set_contact_delete() {
		if (confirm("��࿡ �ش��ϴ� ��������� ��� �����˴ϴ�.\n\n�����Ͻðڽ��ϱ�?")) {
		location.href = "contact_delete_proc.asp?contidx=<%=contidx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>&searchstring=<%=searchstring%>&custcode=<%=custcode%>&custcode2=<%=custcode2%>";
		}
	}

	//��࿬��
	function pop_contact_extention() {
		var url = "pop_contact_extention.asp?contidx=<%=contidx%>";
		var name = "pop_contact_edit";
		var opt = "width=540, height=577, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}
	// ��� ����
	function set_contact_cancel() {
		if (confirm("����� �����Ͻðڽ��ϱ�?")) {
			location.href = "contact_cancel.asp?contidx=<%=contidx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
		}
	}
	// ��� �˻�
	function go_search() {
		var frm = document.forms[0];
		frm.action = "pop_contact_view.asp";
		frm.method = "post";
		frm.submit();
	}


	//  ******************************************************************************  ��ü ���� ����
	//  **************  idx �ش��� ��ü��ȣ, photo �������ϸ�
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

	//  ******************************************************************************  ��ü ���� Ȯ�� �ϱ�

	//  **************  contidx (����ȣ)

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