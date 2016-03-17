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
	if cmonth = "" then cmonth = month(date)
	dim c_date : c_date = Dateserial(cyear, cmonth, "01")

	dim objrs, sql
	sql = "select m.contidx, m.title, m.firstdate, m.startdate, m.enddate, m.canceldate, isnull(t.monthprice,0) as totalprice, c3.custname as custname3 , isnull(sum(d.monthprice),0) as monthprice, isnull(sum(d.expense),0) as expense, d.perform, d.closing from dbo.wb_contact_mst m left outer  join dbo.wb_contact_md m2 on m.contidx = m2.contidx left outer  join dbo.wb_contact_md_dtl d on m2.contidx = d.contidx and m2.sidx = d.sidx inner join dbo.sc_cust_temp c on m.custcode = c.custcode inner join dbo.sc_cust_temp c2 on c.highcustcode = c2.custcode left outer join dbo.wb_medium_mst m3 on m2.mdidx = m3.mdidx inner join dbo.sc_cust_temp c3 on m3.custcode = c3.custcode left outer join dbo.vw_contact_totalprice t on m.contidx = t.contidx where c.highcustcode like '"&custcode&"%' and d.cyear = '"&cyear&"' and d.cmonth = '"&cmonth&"'  and c3.custcode like '"&custcode3&"%' and m.canceldate is null group by m.contidx, m.title, m.firstdate, m.startdate, m.enddate, m.canceldate, t.monthprice , c3.custname, d.perform, d.closing order by m.title "
'	response.write sql

	call get_recordset(objrs, sql)

	dim  contidx, sidx, title, startdate, enddate, totalprice, monthprice, expense, custname3, period, perform, closing, canceldate

	if not objrs.eof Then
		set contidx = objrs("contidx")
		set title = objrs("title")
		set startdate = objrs("startdate")
		set enddate = objrs("enddate")
		set totalprice = objrs("totalprice")
		set monthprice = objrs("monthprice")
		set expense = objrs("expense")
		set custname3 = objrs("custname3")
		set perform = objrs("perform")
		set closing = objrs("closing")
		set canceldate = objrs("canceldate")
	end if
%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒</title>
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
                        <td width="220" align="center" >매체명</td>
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
                        <td width="120" align="center">매체사</td>
                      </tr>
                  </table></td>
                </tr>
              </table>
              <table width="1024" height="31" border="0" cellpadding="0" cellspacing="0" >
	     <%
			dim total_monthprice, total_expense, total_income, total_incomeratio
			do until objrs.eof
			if day(startdate) = "1" then
				period = datediff("m", startdate, enddate)+1
			else
				period = datediff("m", startdate, enddate)
			end if
		%>
                  <tr >
                    <td width="43" align="left"  height="30">&nbsp;<input TYPE="checkbox" NAME="chkitem" ID="chkitem" value="<%=contidx%>" <%if perform then response.write " checked "%> <%if closing then response.write "disabled"%>><%if not closing then %><%if perform then%><IMG SRC="/images/btn_comment-delete.gif" WIDTH="9" HEIGHT="9" BORDER="0" ALT="정산을 취소하시려면 클릭하세요"  onclick="check_execution_cancel('<%=contidx%>')"><%end if%><% end if%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="220" align="left"><span class="stylelink" onclick="go_contact_view(<%=contidx%>)"> <% if not isnull(canceldate) then response.write "<del>"&title&"</del>" else response.write title %></span> </td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="75" align="center"><%=startdate%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="75" align="center"><%=enddate%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="80" align="right"><%If Not IsNull(totalprice) Then response.write formatnumber(totalprice,0) Else response.write "0"%></td>
                    <td width="3"align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="80" align="right"><%If Not IsNull(monthprice) Then response.write formatnumber(monthprice,0) Else response.write "0"%>&nbsp;</td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="80" align="right"><%If Not IsNull(expense) Then response.write formatnumber(expense,0) Else response.write "0"%>&nbsp;</td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="80" align="right"><%If Not IsNull(expense)  Then response.write formatnumber(monthprice-expense,0) Else response.write "0"%>&nbsp;</td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="50" align="right"><%If monthprice <> 0 Then response.write formatnumber((monthprice-expense)/(monthprice)*100, 2) Else response.write "0.00"%>&nbsp;</td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="120" align="right"><%=custname3%>&nbsp;</td>
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
                    <td width="100" align="center">&nbsp;</td>
                  </tr>
              </table></td>
          </tr>
          <tr>
           <td height="40" colspan="30"></td>
          </tr>
      </table></td>
    </tr>
  </table>
<!--#include virtual="/bottom.asp" -->
</form>
<iframe id="ifrm" name="ifrm" width="0" height="0" frameborder="0" src="about:blank"></iframe>
</body>
</html>
<script language="JavaScript">
<!--
	function go_contact_view(idx) {
		var url = "pop_execution_edit.asp?contidx=" + idx +"&cyear=<%=cyear%>&cmonth=<%=cmonth%>";
		var name = "pop_execution_edit" ;
		var opt = "width=1018, height=300, resizable=no,  scrollbars=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function check_execution_reg() {
		if (confirm("선택하신 계약을 정산확인 처리하시겠습니까?")) {
			var frm = document.forms[0];
			var flag = true ;
			for (var i = 0 ; i < frm.chkitem.length; i++) {
				if (frm.chkitem[i].checked)  flag = false ;
			}
			if (flag) {
				alert("정산 확인할 계약을 선택하세요");
				return false;
			}

			frm.action = "execution_reg_proc.asp";
			frm.method = "post";
			frm.submit();
		}
	}

	function check_execution_cancel(contidx) {
		var frm = document.forms[0];
		if (confirm("선택하신 계약을 정산취소 처리하시겠습니까?")) {
			frm.action = "execution_cancel_proc.asp?contidx="+contidx;
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

