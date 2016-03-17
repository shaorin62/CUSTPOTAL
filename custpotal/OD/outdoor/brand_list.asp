<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim gotopage : gotopage = request.querystring("gotopage")
	if gotopage = "" then gotopage = 1
	dim searchstring : searchstring = request("txtsearchstring")
	dim custcode : custcode = request("selcustcode")
	dim custcode2 : custcode2 = request("selcustcode2")
	dim seqno : seqno = request("seljobcust")
	dim thema : thema = request("selthema")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	if cyear = "" then cyear = year(date)
	if cmonth = "" then cmonth = month(date)
	if len(cmonth) = 1 then cmonth = "0"&cmonth
	dim sdate : sdate = Dateserial(cyear, cmonth, "01")

	dim old_thema : old_thema = thema
	dim objrs, sql
	sql = "select contidx, title, locate, side, qty, unit, m.idx, standard, quality, firstdate, startdate, enddate, isnull(totalprice,0) as totalprice, isnull(monthprice,0) as monthprice, isnull(expense,0) as expense, seqname, thema, custname, medname from ( select m.contidx, m.title, m2.locate, d.idx, d.side, a.qty, m2.unit, d.standard, d.quality, m.firstdate, m.startdate, m.enddate, a.monthprice, a.expense,j2.seqname, j.thema, c.custname, c2.custname as medname from dbo.wb_contact_mst m left outer join dbo.wb_contact_md m2 on m.contidx = m2.contidx left outer join dbo.wb_contact_md_dtl d on m2.sidx = d.sidx left outer join dbo.wb_contact_md_dtl_account a on d.idx = a.idx and a.cyear = '"&cyear&"' and a.cmonth = '"&cmonth&"' left outer join dbo.sc_cust_temp c on c.custcode = m.custcode left outer join dbo.sc_cust_temp c2 on c2.custcode = m2.medcode left outer join dbo.wb_jobcust j on a.jobidx = j.jobidx left outer join dbo.sc_jobcust j2 on j.seqno = j2.seqno where m.canceldate >= '" & sdate & "' and m.custcode like '" & custcode2 & "%' and c.highcustcode like '" & custcode & "%' and j.seqno like '" & seqno & "%') as m left outer join ( select d_.idx, sum(monthprice) as totalprice from dbo.wb_contact_mst m_ left outer join dbo.wb_contact_md m2_ on m_.contidx = m2_.contidx left outer join dbo.wb_contact_md_dtl d_ on m2_.sidx = d_.sidx left outer join dbo.wb_contact_md_dtl_account a_ on d_.idx = a_.idx group by d_.idx) as d on m.idx = d.idx  "
	if thema <> "" then sql = sql & " where thema = '" & old_thema & "' "
	sql = sql & " order by contidx desc "

	call get_recordset(objrs, sql)

	dim contidx, title, startdate, enddate, side, seqname, standard, quality, monthprice, custname, canceldate, contactcancel

	dim cnt : cnt = objrs.recordcount

	if not objrs.eof Then
		set contidx = objrs("contidx")
		set title = objrs("title")
		set startdate = objrs("startdate")
		set enddate = objrs("enddate")
		set side = objrs("side")
		set seqname = objrs("seqname")
		set standard = objrs("standard")
		set quality = objrs("quality")
		set monthprice = objrs("monthprice")
		set custname = objrs("custname")
		set thema = objrs("thema")
	end if
	if seqno = "" then seqno = "0"
%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  oncontextmenu="return false">
<form name="form1" method="post" action="">
<!--#include virtual="/od/top.asp" -->
  <table width="1240" height="652" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="Table_01">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/od/left_outdoor_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_outdoor.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td height="600" align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="19">&nbsp;</td>
          </tr>
          <tr>
            <td height="17"><TABLE  width="100%">
            <TR>
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"> 브랜드별 집행현황 </span></TD>
				<TD width="50%" align="right"><span class="navigator" > 옥외관리 &gt;  옥외광고현황 &gt; 브랜드별 집행현황  </span></TD>
            </TR>
            </TABLE>
			</td>
          </tr>
          <tr>
            <td height="15">&nbsp;</td>
          </tr>
          <tr>
            <td >			<table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
                  <td align="left" background="/images/bg_search.gif"> <%call get_year(cyear)%><%call get_month(cmonth)%>&nbsp; <%call get_custcode_mst(custcode, null, "contact_list.asp")%><span id="custcode2"><%call get_blank_select("사업부를 선택하세요", 207)%></span><span id='jobcust'><%call get_blank_select("브랜드를 선택하세요", 207)%></span><span id='thema'><%call get_blank_select("소재를 선택하세요", 207)%></span> <img src="/images/btn_search.gif" width="39" height="20" align="top" class="styleLink" onclick="go_search();"> </td>
                  <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="26" align="right"><img src="/images/btn_excel.gif" width="85" height="22" align="absmiddle" vspace="5" class="stylelink" onclick="get_excel_sheet();"></td>
          </tr>
          <tr>
            <td ><table width="1030"  border="3" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
                <tr>
                  <td>
				  <table width="1024" border="0" cellspacing="0" cellpadding="0" class="header">
                      <tr>
                        <td width="40" align="center" height="30">No</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="230" align="center" >매체명</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="75" align="center">시작일</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="75" align="center">종료일</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="30" align="center">면</td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="100" align="center">소재명</td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="75" align="center">브랜드</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="150" align="center">규격 / 재질</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="80" align="center">월광고료</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="70" align="center">운영팀</td>
                      </tr>
                  </table></td>
                </tr>
              </table>
              <table width="1024" height="31" border="0" cellpadding="0" cellspacing="0" style="padding-left:3px;" >
	     <%
			do until objrs.eof
		%>
                  <tr onClick="go_contact_view(<%=contidx%>)" class="styleLink" >
                    <td width="40" align="center"  height="30"><%=cnt%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="230" align="left"><%=title %></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="75" align="center"><%=startdate%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="75" align="center"><%=enddate%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="30" align="center"><%=side%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="100" align="center"><%=thema%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="75" align="center"><%=seqname%></td>
                    <td width="3"align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="150" align="center"><%if not isnull(standard) then response.write standard %> <%if not isnull(quality) then response.write " / " & quality %></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="80" align="right"><%if not isnull(monthprice) then response.write formatnumber(monthprice,0) else response.write "0"%>&nbsp;</td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="70" align="center"><%=custname%></td>
                  </tr>
                  <tr>
                    <td height="1" bgcolor="#E7E9E3" colspan="30"></td>
                  </tr>
				<%
						cnt = cnt -1
						objrs.movenext
					loop
					objrs.close
					set objrs = nothing
				%>
              </table></td>
          </tr>
      </table></td>
    </tr>
  </table>
</form>
<iframe id="ifrm" name="ifrm" width="0" height="0" frameborder="0" src="about:blank"></iframe>
</body>
</html>
<!--#include virtual="/bottom.asp" -->
<script language="JavaScript">
<!--
	function go_contact_view(idx) {
		var url = "pop_contact_view.asp?contidx=" + idx+"&cyear=<%=cyear%>&cmonth=<%=cmonth%>" ;
		var name = "pop_contact_view" ;
		var opt = "width=1258,resizable=yes, scrollbars=yes, status=yes, , menubar=no, top=100, right=0";
		window.open(url, name, opt);
	}

	function go_page(url) {
		var frm = document.getElementById("ifrm") ;
		var custcode = document.forms[0].selcustcode.options[document.forms[0].selcustcode.selectedIndex].value ;
		var custcode2 = "<%=custcode2%>" ;
		var seqno = "<%=seqno%>";
		frm.src="/inc/frm_code.asp?custcode="+custcode +"&custcode2="+custcode2+"&seqno="+seqno;
	}

	function go_search() {
		var frm = document.forms[0];
		frm.action = "brand_list.asp";
		frm.method = "post";
		frm.submit();
	}

	function get_excel_sheet() {
		location.href = "xls_brand_list.asp?selcustcode=<%=custcode%>&selcustcode2=<%=custcode2%>&seljobcust=<%=seqno%>&searchstring=<%=searchstring%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>&selthema=<%=old_thema%>";

	}

	window.onload = function () {
		go_page("");
	}
//-->
</script>

