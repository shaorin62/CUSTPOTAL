<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim gotopage : gotopage = request.querystring("gotopage")
	if gotopage = "" then gotopage = 1
	dim searchstring : searchstring = request("txtsearchstring")
	dim custcode : custcode = request("selcustcode")
	dim custcode3 : custcode3 = request("selcustcode3")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	if cyear = "" then cyear = year(date)
	if cmonth = "" then cmonth = Month(date)
	if len(cmonth) = 1 then cmonth = "0"&cmonth
	dim sdate : sdate = dateserial(cyear, cmonth, "01")
	dim edate : edate = dateadd("d", -1, dateadd("m", 1,  sdate))

	dim objrs, sql
	sql = "select m.contidx, title, firstdate, startdate, enddate, isnull(totalprice,0) as totalprice, monthprice, expense, custname, canceldate, IsPerform,IsClosing, medname from ( select m.contidx, m.title, m.firstdate, m.startdate, m.enddate, isnull(sum(a.monthprice),0) as monthprice, isnull(sum(a.expense),0) as expense, c.custname, m.canceldate, a.isperform, a.Isclosing, c2.custname as medname from dbo.wb_contact_mst m left outer  join dbo.sc_cust_temp c on c.custcode = m.custcode left outer join dbo.wb_contact_md m2 on m.contidx = m2.contidx left outer join dbo.wb_contact_md_dtl d on m2.sidx = d.sidx left outer join dbo.sc_cust_temp c2 on c2.custcode = m2.medcode left outer join dbo.wb_contact_md_dtl_account a on d.idx = a.idx and a.cyear = '" & cyear & "' and a.cmonth = '" & cmonth & "'  where m2.medcode like '"&custcode3&"%' and c.highcustcode like '"&custcode&"%'   and m.canceldate <= m.enddate group by m.contidx, m.title, m.firstdate, m.startdate, m.enddate, c.custname, m.canceldate, a.isPerform, a.Isclosing, c2.custname ) as m left outer join (select m.contidx, sum(a.monthprice) as totalprice from dbo.wb_contact_md m inner join dbo.wb_contact_md_dtl d on m.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on d.idx = a.idx group by m.contidx ) as d on m.contidx = d.contidx where startdate <= '" & edate & "' and enddate >= '" & sdate & "'  and m.canceldate >= '" & sdate & "'  order by m.contidx desc"

	call get_recordset(objrs, sql)

	dim cnt, contidx, title, firstdate, startdate, enddate, totalprice, monthprice, expense, clientsubname, income, incomeRatio, gMonthPrice, gExpense, gIncome, gIncomeRatio, isPerform, isClosing, canceldate, medname

	cnt = objrs.recordcount

	if not objrs.eof Then
		set contidx = objrs("contidx")
		set title = objrs("title")
		set firstdate = objrs("firstdate")
		set startdate = objrs("startdate")
		set enddate = objrs("enddate")
		set totalprice = objrs("totalprice")
		set monthprice = objrs("monthprice")
		set expense = objrs("expense")
		set clientsubname = objrs("custname")
		set medname = objrs("medname")
		set isPerform = objrs("isPerform")
		set isClosing = objrs("isClosing")
		set canceldate = objrs("canceldate")
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
<!--#include virtual="/od/top.asp" -->
  <table width="1240" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="Table_01">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/od/left_outdoor_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_outdoor.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td align="left" valign="top">
	  <table width="1030" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="19">&nbsp;</td>
          </tr>
          <tr>
            <td height="17"><TABLE  width="100%">
            <TR>
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"> 광고비용 집행현황 </span></TD>
				<TD width="50%" align="right"><span class="navigator" > 옥외관리 &gt; 옥외광고현황 &gt; 광고비용 집행현황   </span></TD>
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
                  <td background="/images/bg_search.gif"> <img src="/images/icon_dot_search.gif" width="4" height="3" align="absmiddle"> 검색년월
				  <%call get_year(cyear)%> <%call get_month(cmonth)%> &nbsp;   <%call get_custcode_mst(custcode, null, "excution_list.asp")%><span id="custcode3"><%call get_custcode_custcode3(custcode3, null)%></span> <img src="/images/btn_search.gif" width="39" height="20" align="top" class="stylelink" onclick="go_search();">
				 </td>
                  <td  align="right" background="/images/bg_search.gif" ><img src="/images/btn_execution.gif" width="78" height="18" align="absmiddle" border="0" onclick="check_execution_reg();" class="styleLink"> <!-- <img src="/images/btn_execution_cancel.gif" width="78" height="18" align="absmiddle" border="0" onclick="check_execution_cancel();" class="styleLink"> --></td>
                  <td width="13" ><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="26">&nbsp;</td>
          </tr>
          <tr>
            <td><table width="1030"  border="3" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
                <tr>
                  <td>
				  <table width="1024" border="0" cellspacing="0" cellpadding="0" class="header">
                      <tr>
                        <td width="40" align="center" height="30"><INPUT TYPE="checkbox" NAME="chkAll" onclick="check_All();" ID="chkAll"></td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="200" align="center" >매체명</td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="75" align="center">시작일</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="75" align="center">종료일</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="80" align="center">총광고료</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="80" align="center">월광고료</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="80" align="center">월지급액</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="80" align="center">내수액</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="50" align="center">내수율</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="80" align="center">사업부서</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="100" align="center">매체사</td>
                      </tr>
                  </table></td>
                </tr>
              </table>
              <table width="1024" height="31" border="0" cellpadding="0" cellspacing="0" >
	     <%
			dim total_monthprice, total_expense, total_income, total_incomeratio
			do until objrs.eof
		%>
                  <tr >
                    <td width="43" align="left"  height="30">&nbsp;<input TYPE="checkbox" NAME="chkitem" ID="chkitem" value="<%=contidx%>" <%if IsPerForm then response.write " checked "%> <%if IsClosing then response.write "disabled"%>><%if not IsClosing then %><%if IsPerForm then%><IMG SRC="/images/btn_comment-delete.gif" WIDTH="9" HEIGHT="9" BORDER="0" ALT="정산을 취소하시려면 클릭하세요"  onclick="check_execution_cancel('<%=contidx%>')"><%end if%><% end if%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="200" align="left"><span class="stylelink" onclick="go_contact_view(<%=contidx%>)"> <%= title %></span> </td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="75" align="center"><%=startdate%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="75" align="center"><%=enddate%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="80" align="right"><%=formatnumber(totalprice,0)%></td>
                    <td width="3"align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="80" align="right"><%=formatnumber(monthprice,0)%>&nbsp;</td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="80" align="right"><%=formatnumber(expense,0)%>&nbsp;</td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="80" align="right"><%If expense <> 0  Then response.write formatnumber(monthprice-expense,0) Else response.write "0"%>&nbsp;</td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="50" align="right"><%If monthprice <> 0 Then response.write formatnumber((monthprice-expense)/(monthprice)*100, 2) Else response.write "0.00"%>&nbsp;</td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="80" align="center"><%=clientsubname%>&nbsp;</td>
                    <td width="100" align="right"><%=medname%>&nbsp;</td>
                  </tr>
                  <tr>
                    <td height="1" bgcolor="#E7E9E3" colspan="30"></td>
                  </tr>
				<%
						total_monthprice = total_monthprice + monthprice
						total_expense = total_expense + expense
						objrs.movenext
					loop
					objrs.close
					set objrs = nothing
				%>
                  <tr height="40" bgcolor="#ECECEC">
                    <td  align="center" colspan="10" class="header">총합계 </td>
                    <td width="80" align="right" class="header"><%If Not IsNull(total_monthprice) Then response.write formatnumber(total_monthprice,0) Else response.write "0"%>&nbsp;</td>
                    <td width="3" align="center"></td>
                    <td width="80" align="right" class="header"><%If Not IsNull(total_expense) Then response.write formatnumber(total_expense,0) Else response.write "0"%>&nbsp;</td>
                    <td width="3" align="center"></td>
                    <td width="80" align="right" class="header"><%If total_monthprice <> 0  Then response.write formatnumber(total_monthprice-total_expense,0) Else response.write "0"%>&nbsp;</td>
                    <td width="3" align="center"></td>
                    <td width="50" align="right" class="header"><%If total_monthprice <> 0 Then response.write formatnumber((total_monthprice-total_expense)/total_monthprice*100, 2) Else response.write "0.00"%>&nbsp;</td>
                    <td width="3" align="center"></td>
                    <td width="100" align="center" colspan="2">&nbsp;</td>
                  </tr>
              </table></td>
          </tr>
          <tr>
           <td height="40" colspan="30"></td>
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
//
//	function go_contact_view(idx) {
//		var url = "pop_execution_edit.asp?contidx=" + idx +"&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
//		var name = "pop_execution_edit" ;
//		var opt = "width=1018, height=300, resizable=no,  scrollbars=yes, top=100, left=100";
//		window.open(url, name, opt);
//	}

	function check_execution_reg() {
		if (confirm("선택하신 계약을 정산확인 처리하시겠습니까?")) {
			var frm = document.forms[0];
			var flag = true ;
			if (frm.chkitem.length == undefined) {
				if (frm.chkitem.checked)  flag = false;
			} else {
				for (var i = 0 ; i < frm.chkitem.length; i++) {
					if (frm.chkitem[i].checked)  	flag = false ;
				}
				if (flag) {
					alert("정산 확인할 계약을 선택하세요");
					return false;
				}
			}

			frm.action = "execution_reg_proc.asp";
			frm.method = "post";
			frm.submit();
		}
	}

	function check_execution_cancel(contidx) {
		var frm = document.forms[0];
		if (confirm("선택하신 계약을 정산취소 처리하시겠습니까?")) {
			location.href = "execution_cancel_proc.asp?contidx="+contidx+"&cyear=<%=cyear%>&cmonth=<%=cmonth%>&custcode=<%=custcodE%>";
		}
	}

	function go_page(url) {
		return false ;
		var frm = document.getElementById("ifrm") ;
		var custcode = document.forms[0].selcustcode.options[document.forms[0].selcustcode.selectedIndex].value ;
		var custcode3 = "<%=custcode3%>" ;
		frm.src="/inc/frm_code.asp?custcode="+custcode +"&custcode3="+custcode3;
	}

	function go_search() {
		var frm = document.forms[0];
		frm.action = "execution_list.asp";
		frm.method = "post";
		frm.submit();
	}

	function check_All() {
		var frm = document.forms[0];
		var flag = frm.chkAll.checked ;

		if (isNaN(Number(frm.chkitem.length))) {
			frm.chkitem.checked = flag ;
		} else {
			for (var i = 0;  i < frm.chkitem.length ; i++) {
				frm.chkitem[i].checked = flag ;
			}
		}
	}

	window.onload = function () {
		go_page("");
	}
//-->
</script>

