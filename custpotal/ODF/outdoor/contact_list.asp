<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim gotopage : gotopage = request.querystring("gotopage")
	if gotopage = "" then gotopage = 1
	dim searchstring : searchstring = request("txtsearchstring")
	if searchstring = "�˻��� ��ü���� �Է��ϼ���" then searchstring = ""
	dim custcode : custcode = request("selcustcode")
	dim custcode2 : custcode2 = request("selcustcode2")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	if cyear = "" then cyear = year(date)
	if cmonth = "" then cmonth = month(date)
	dim s_date : s_date = Dateserial(cyear, cmonth, "01")
	dim e_date : e_date = DateAdd("d", -1, DateAdd("m", 1, s_date))
	dim total_monthprice, total_expense, total_income, total_incomeratio
	Dim chkAll : chkAll = request("chkAll")

	dim objrs, sql
	sql = "select m.contidx, m.title, m.firstdate, m.startdate, m.enddate, m.canceldate, isnull(t.monthprice,0) as totalprice, isnull(sum(d.monthprice),0) as monthprice, isnull(sum(d.expense),0) as expense, c.custname as custname2 from dbo.wb_contact_mst m left outer join dbo.vw_contact_totalprice t on m.contidx = t.contidx left outer join dbo.wb_contact_md_dtl d on m.contidx = d.contidx  left outer join dbo.sc_cust_temp c on m.custcode = c.custcode left outer join dbo.sc_cust_temp c2 on c.highcustcode = c2.custcode where m.title like '%" & searchstring & "%' and m.custcode like '"&custcode2&"%' and c2.custcode like '"&custcode&"%'  and d.cyear =  '"&cyear&"' and d.cmonth = '"&cmonth&"' "
	'response.write sql
	If chkAll = "" Then sql = sql & " and m.canceldate is null "
	sql = sql & " group by m.contidx, m.title, m.firstdate, m.startdate, m.enddate, m.canceldate, t.monthprice, c.custname order by m.title"
	call get_recordset(objrs, sql)

	dim cnt, contidx, title, firstdate, startdate, enddate, period, monthprice, expense, income, incomeratio, custname2, totalprice,canceldate

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
		set canceldate = objrs("canceldate")
		set custname2 = objrs("custname2")
	end if
%>
<html>
<head>
<title>�Ƣ� SK M&C | Media Management System �Ƣ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  oncontextmenu="return false">
<form>
<!--#include virtual="/od/top.asp" -->
  <table width="1240" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="Table_01">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/od/left_outdoor_menu.asp"--></td>
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
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"> ���ܱ��� ������Ȳ </span></TD>
				<TD width="50%" align="right"><span class="navigator" > ���ܰ��� &gt;  ���ܱ�����Ȳ &gt; ���ܱ��� ������Ȳ </span></TD>
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
                  <td background="/images/bg_search.gif"> <img src="/images/icon_dot_search.gif" width="4" height="3" align="absmiddle"> �˻����
				  <%call get_year(cyear)%> <%call get_month(cmonth)%> &nbsp;    <%call get_custcode_mst(custcode, null, "contact_list.asp")%><span id="custcode2"><%call get_blank_select("����θ� �����ϼ���", 207)%></span><input type="text" name="txtsearchstring" size="30"  id="txtsearchstring" > <img src="/images/btn_search.gif" width="39" height="20" align="top" class="stylelink" onclick="go_search();"> <INPUT TYPE="checkbox" NAME="chkAll"> �����ü </td>
                  <td  align="right" background="/images/bg_search.gif" ><img src="/images/btn_contact_reg.gif" width="78" height="18" align="absmiddle" border="0" onclick="pop_contact_reg();" class="styleLink"> </td>
                  <td width="13" ><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="26" align="right"><img src="/images/btn_excel.gif" width="85" height="22" align="absmiddle" vspace="5" class="stylelink" onclick="get_excel_sheet();"></td>
          </tr>
          <tr>
            <td><table width="1030"  border="3" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
                <tr>
                  <td>
				  <table width="1024" border="0" cellspacing="0" cellpadding="0" class="header">
                      <tr>
                        <td width="40" align="center" height="30">No</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="220" align="center" >��ü��</td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="75" align="center">����<br>�������</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="75" align="center">������</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="75" align="center">������</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="80" align="center">�ѱ����</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="80" align="center">�������</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="80" align="center">�����޾�</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="80" align="center">������</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="50" align="center">������</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="100" align="center">����μ�</td>
                      </tr>
                  </table></td>
                </tr>
              </table>
              <table width="1024" height="31" border="0" cellpadding="0" cellspacing="0" >
	     <%
			do until objrs.eof
			if day(startdate) = "1" then
				period = datediff("m", startdate, enddate)+1
			else
				period = datediff("m", startdate, enddate)
			end if

		%>
                  <tr onClick="go_contact_view(<%=contidx%>)" class="styleLink" >
                    <td width="43" align="center"  height="30"><%=cnt%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="220" align="left"><% if not isnull(canceldate) then response.write "<del>"&title&"</del>" else response.write title %><%if DateDiff("m", date, enddate) < 3 then %> <img src="/images/icon_notice.gif" width="10" height="13" hspace="5"><%end if%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="75" align="center"><%=firstdate%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="75" align="center"><%=startdate%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="75" align="center"><%=enddate%></td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="80" align="right"><%If Not IsNull(totalprice) Then response.write formatnumber(totalprice,0) Else response.write "0"%>&nbsp;</td>
                    <td width="3"align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="80" align="right"><%If Not IsNull(monthprice) or monthprice <> 0 Then response.write formatnumber(monthprice,0) Else response.write "0"%>&nbsp;</td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="80" align="right"><%If Not IsNull(expense) Then response.write formatnumber(expense,0) Else response.write "0"%>&nbsp;</td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="80" align="right"><%If Not IsNull(expense)  Then response.write formatnumber(monthprice-expense,0) Else response.write "0"%>&nbsp;</td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="50" align="right"><%If monthprice <> 0 Then response.write formatnumber((monthprice-expense)/monthprice*100, 2) Else response.write "0.00"%>&nbsp;</td>
                    <td width="3" align="center"><img src="/images/dot_bg.gif" width="3" height="5"></td>
                    <td width="100" align="center"><%=custname2%>&nbsp;</td>
                  </tr>
                  <tr>
                    <td height="1" bgcolor="#E7E9E3" colspan="30"></td>
                  </tr>
				<%
						total_monthprice = total_monthprice + monthprice
						total_expense = total_expense + expense
						cnt = cnt -1
						objrs.movenext
					loop
					objrs.close
					set objrs = nothing

					total_income = total_monthprice - total_expense
					if total_income = 0 then
						total_incomeratio = "0.00"
					else
						total_incomeratio = total_monthprice - total_expense / total_monthprice * 100
					end if
				%>
                  <tr height="40" bgcolor="#ECECEC">
                    <td  align="center" colspan="12" class="header">���հ� </td>
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
		var url = "pop_contact_view.asp?contidx=" + idx+"&cyear=<%=cyear%>&cmonth=<%=cmonth%>" ;
		var name = "pop_contact_view" ;
		var opt = "width=1258,resizable=yes, scrollbars=yes, status=yes, , menubar=no, top=100, right=0";
		window.open(url, name, opt);
	}

	function pop_contact_reg() {
		var url = "pop_contact_reg.asp"
		var name = "pop_contact_reg";
		var opt = "width=540, height=577, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);
	}

	function go_page(url) {
		var frm = document.getElementById("ifrm") ;
		var custcode = document.forms[0].selcustcode.options[document.forms[0].selcustcode.selectedIndex].value ;
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
		var custname = document.forms[0].selcustcode.options[document.forms[0].selcustcode.selectedIndex].text ;
		location.href = "xls_contact_list.asp?custcode=<%=custcode%>&custcode2=<%=custcode2%>&searchstring=<%=searchstring%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>&custname="+custname;
	}

	window.onload = function () {
		go_page("");
	}
//-->
</script>

