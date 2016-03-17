<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim custcode : custcode = request.cookies("custcode")
	dim custcode2 : custcode2 = request("selcustcode2")
	dim searchstring : searchstring = request("txtsearchstring")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	if cyear = "" then cyear = year(date)
	if cmonth = "" then cmonth = month(date)
	dim canceldate : canceldate = dateserial(cyear, cmonth, "01")
	if len(cmonth) = 1 then cmonth = "0"&cmonth
	dim sdate : sdate = dateserial(cyear, cmonth, "01")
	dim edate : edate = dateadd("d", -1, dateadd("m", 1,  sdate))

	dim objrs, sql
	sql = "select m.contidx, title, firstdate, startdate, enddate, isnull(totalprice,0) as totalprice, monthprice, expense, custname, canceldate  from ( select m.contidx, m.title, m.firstdate, m.startdate, m.enddate, isnull(sum(a.monthprice),0) as monthprice, isnull(sum(a.expense),0) as expense, c.custname, m.canceldate from dbo.wb_contact_mst m inner join dbo.sc_cust_temp c on c.custcode = m.custcode left outer join dbo.wb_contact_md m2 on m.contidx = m2.contidx left outer join dbo.wb_contact_md_dtl d on m2.sidx = d.sidx left outer join dbo.wb_contact_md_dtl_account a on d.idx = a.idx and a.cyear = '" & cyear & "' and a.cmonth = '" & cmonth & "'  where m.custcode like '"&custcode2&"%' and c.highcustcode like '"&custcode&"%'   and m.title like '%"&searchstring&"%' and m.canceldate <= m.enddate group by m.contidx, m.title, m.firstdate, m.startdate, m.enddate, c.custname, m.canceldate ) as m left outer join (select m.contidx, sum(a.monthprice) as totalprice from dbo.wb_contact_md m inner join dbo.wb_contact_md_dtl d on m.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on d.idx = a.idx group by m.contidx ) as d on m.contidx = d.contidx where startdate <= '" & edate & "' and enddate >= '" & sdate & "' and m.canceldate >= '" & sdate & "'  order by m.contidx desc	"

	response.write sql
	response.end
	call get_recordset(objrs, sql)

	dim cnt, contidx, title, firstdate, startdate, enddate, totalprice, monthprice, expense, highcustcode, clientsubname, income, incomeRatio, gMonthPrice, gExpense, gIncome, gIncomeRatio

	cnt = objrs.recordcount

	if not objrs.eof Then
		set contidx = objrs("contidx")
		set title = objrs("title")
		set firstdate = objrs("firstdate")
		set startdate = objrs("startdate")
		set enddate = objrs("enddate")
		set totalprice = objrs("totalprice")
		set monthprice = objrs("monthprice")
		set clientsubname = objrs("custname")
	end if
%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  oncontextmenu="return false">
<form>
<INPUT TYPE="hidden" NAME="custcode" value="<%=custcode%>">
<!--#include virtual="/cust/top.asp" -->
  <table width="1240" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="Table_01">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/cust/left_outdoor_menu.asp"--></td>
      <td height="65" valign="top"><img src="/images/middle_navigater_outdoor.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td align="left" valign="top" >
	  <table width="1030" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="19">&nbsp;</td>
          </tr>
          <tr>
            <td height="17"><TABLE  width="100%">
            <TR>
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"> 옥외광고 집행현황 </span></TD>
				<TD width="50%" align="right"><span class="navigator" > 옥외관리 &gt;  옥외광고현황 &gt; 옥외광고 집행현황 </span></TD>
            </TR>
            </TABLE>
			</td>
          </tr>
          <tr>
            <td height="15">&nbsp;</td>
          </tr>
          <tr>
            <td >
			<table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13" valign="top" ><img src="/images/bg_search_left.gif" width="13" height="35" ></td>
                  <td background="/images/bg_search.gif"> <img src="/images/icon_dot_search.gif" width="4" height="3" align="absmiddle"> 계약검색  <%call get_year(cyear)%> <%call get_month(cmonth)%> &nbsp;   <span id="custcode2">  <%call get_blank_select("사업부를 선택하세요", 207)%></span>  <input type="text" name="txtsearchstring" size="30"  id="txtsearchstring" > <img src="/images/btn_search.gif" width="39" height="20" align="top" class="stylelink" onclick="go_search();"></td>
                  <td  align="right" background="/images/bg_search.gif" > </td>
                  <td width="13" ><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="26" align="right"><!-- <img src="/images/btn_excel.gif" width="85" height="22" align="absmiddle" vspace="5" class="stylelink" onclick="get_excel_sheet();"> --></td>
          </tr>
          <tr>
            <td valign="top" height="516"><table width="1030"  border="3" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
                <tr>
                  <td>
				  <table width="1024" border="0" cellspacing="0" cellpadding="0" class="header">
                      <tr>
                        <td width="40" align="center" height="30">No</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="400" align="center" >매체명</td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="100" align="center">최초<br>계약일자</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="100" align="center">시작일</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="100" align="center">종료일</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="100" align="center">총광고료</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="100" align="center">월광고료</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="150" align="center">사업부서</td>
                      </tr>
                  </table></td>
                </tr>
              </table>
              <table width="1024" height="31" border="0" cellpadding="0" cellspacing="0" >
	     <%
			do until objrs.eof
		%>
                  <tr  >
                    <td width="40" align="center"  height="30"><%=cnt%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="400" align="left" onClick="go_contact_view(<%=contidx%>,'<%=highcustcode%>')" class="styleLink"><%=title %><%if DateDiff("m", date, enddate) < 3 then %> <img src="/images/icon_notice.gif" width="10" height="13" hspace="5"><%end if%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="100" align="center"><%=firstdate%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="100" align="center"><%=startdate%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="100" align="center"><%=enddate%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="100" align="right"><%If Not IsNull(totalprice) Then response.write formatnumber(totalprice,0) Else response.write "0"%>&nbsp;</td>
                    <td width="3"align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="100" align="right"><%If Not IsNull(monthprice) or monthprice <> 0 Then response.write formatnumber(monthprice,0) Else response.write "0"%>&nbsp;</td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="150" align="center"><%=clientsubname%>&nbsp;</td>
                  </tr>
                  <tr>
                    <td height="1" bgcolor="#E7E9E3" colspan="15"></td>
                  </tr>
				<%
						gMonthPrice = gMonthPrice + monthprice
						cnt = cnt -1
						objrs.movenext
					loop
					objrs.close
					set objrs = nothing
				%>
                  <tr height="40" bgcolor="#ECECEC">
                    <td width="40" align="center"  height="30" colspan="10">  </td>
                    <td width="100" align="center">월광고료 합계</td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="100" align="right"><%If Not IsNull(gMonthPrice) Then response.write formatnumber(gMonthPrice,0) Else response.write "0"%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="150" align="center"></td>
                  </tr>
                    <td  align="center" colspan="15" class="header"></td>
                  </tr>
              </table></td>
          </tr>
      </table></td>
    </tr>
  </table>
<INPUT TYPE="hidden" NAME="menunum" value="<%=menunum%>">
</form>
<iframe id="ifrm" name="ifrm" width="0" height="0" frameborder="0" src="about:blank"></iframe>
</body>
</html>
<!--#include virtual="/bottom.asp" -->
<script language="JavaScript">
<!--
	function go_contact_view(idx) {
		var url = "pop_contact_view.asp?contidx=" + idx+"&cyear=<%=cyear%>&cmonth=<%=cmonth%>&searchstring=<%=searchstring%>&custcode=<%=custcode%>&custcode2=<%=custcode2%>" ;
		var name = "pop_contact_view" ;
		var opt = "width=1258,resizable=yes, scrollbars=yes, status=yes, , menubar=no, top=100, right=0";
		window.open(url, name, opt);
	}

	function pop_contact_reg() {
		var url = "pop_contact_reg.asp?custcode=<%=custcode%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
		var name = "pop_contact_reg";
		var opt = "width=540, height=577, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	function go_page(url) {
		var frm = document.getElementById("ifrm") ;
		var custcode = "<%=custcode%>" ;
		var custcode2 = "<%=custcode2%>" ;
		frm.src="/inc/frm_code.asp?custcode="+custcode +"&custcode2="+custcode2;
	}

	function go_search() {
		var frm = document.forms[0];
		frm.action = "contact_list.asp";
		frm.method = "post";
		frm.submit();
	}

	function get_excel_sheet() {
		var custname = "<%=request.cookies("custname")%>" ;
		location.href = "xls_contact_list.asp?custcode=<%=custcode%>&custcode2=<%=custcode2%>&searchstring=<%=searchstring%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>&custname="+custname;
	}

	window.onload = function () {
		go_page("");
	}
//-->
</script>

