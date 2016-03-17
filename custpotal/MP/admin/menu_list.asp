<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
	dim mnu_menu : mnu_menu = request("menunum")
	response.cookies("menunum") = mnu_menu
	dim custcode : custcode = request("selcustcode2")
	if custcode = "" then custcode = mnu_menu

	dim objrs, sql, acc_menu
	sql = "select c.custname, c2.custname from dbo.sc_cust_temp c inner join  dbo.sc_cust_temp c2 on c.custcode = c2.highcustcode where c2.custcode = '" & custcode & "' "
	call get_recordset(objrs, sql)
	dim custname, custname2
	custname = objrs(0).value
	custname2 = objrs(1).value
	objrs.close

	sql = "select m.midx, m.title, c.custname, m.lvl, m.isfile, iscomment, isemail from dbo.wb_menu_mst m inner join dbo.sc_cust_temp c on m.custcode = c.custcode where m.custcode = '"&custcode&"' or m.custcode = (select highcustcode from dbo.sc_cust_temp where custcode = '"&custcode&"') order by  m.ref, m.lvl"
	call get_recordset(objrs, sql)

	dim midx, title, isfile, isemail, iscomment, lvl
	if not objrs.eof then
		set midx = objrs("midx")
		set title = objrs("title")
		set isfile = objrs("isfile")
		set isemail = objrs("isemail")
		set iscomment = objrs("iscomment")
		set lvl = objrs("lvl")
	end if
%>
<html>
<head>
<title>▒ SK MARKETING & COMPANY ▒</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  oncontextmenu="return false">

<form >
<!--#include virtual="/hq/top.asp" -->
  <table id="Table_01" width="1240" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/hq/left_admin_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_admin.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td height="600" align="left" valign="top"><table width="1030" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="19">&nbsp;</td>
          </tr>
          <tr>
            <td height="17"><TABLE  width="100%">
            <TR>
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"> <%=custname2%> 메뉴현황 </span></TD>
				<TD width="50%" align="right"><span class="navigator" > 관리모드 &gt; 메뉴관리 &gt; <%=custname%> <%if custname <> custname2 then response.write "&gt; " & custname2 %> </span></TD>
            </TR>
            </TABLE>
			</td>
          </tr>
          <tr>
            <td height="15">&nbsp;</td>
          </tr>
          <tr>
            <td ><table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
                  <td width="80%" align="left" background="/images/bg_search.gif"> <%call get_custcode_custcode2(mnu_menu, custcode)%> <img src="/images/btn_search.gif" width="39" height="20" align="top" class="styleLink" onClick="search_cust_dept()"></td>
                  <td width="20%" align="right" background="/images/bg_search.gif"><img src="/images/btn_menu_reg.gif" width="78" height="18" alt="" border="0" onclick="pop_menu_reg();" class="stylelink"></td>
                  <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="26" >&nbsp;</td>
          </tr>
          <tr>
            <td ><table width="1030" height="31" border="3" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
                <tr>
                  <td><table width="1024" border="0" cellspacing="0" cellpadding="0" class="header">
                      <tr>
                        <td width="724" align="center" >메뉴명</td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="100" align="center" >첨부파일</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="100" align="center" >메일발송</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="100" align="center" >댓글기능</td>
                      </tr>
                  </table></td>
                </tr>
              </table>
                <table width="1024"  border="0" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
				<% do until objrs.eof %>
                  <tr onClick="go_menu_view('<%=midx%>')" class="styleLink" height="31">
                    <td width="724" align="left"  class="styleLink header"  style="padding-left:<%=lvl*15%>px;"><%=title%>&nbsp;</td>
                    <td width="3" align="center">&nbsp;</td>
                    <td width="100" align="center"><%if isfile then response.write "사용"%>&nbsp;</td>
                    <td width="3">&nbsp;</td>
                    <td width="100" align="center"><%if isemail then response.write "사용"%>&nbsp;</td>
                    <td width="3">&nbsp;</td>
                    <td width="100" align="center"><%if iscomment then response.write "사용"%>&nbsp;</td>
                  </tr>
                  <tr>
                    <td height="1" bgcolor="#E7E9E3" colspan="13"></td>
                  </tr>
				<%
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
<!--#include virtual="/bottom.asp" -->
</body>
</html>
<script language="JavaScript">
<!--
	function pop_menu_reg() {
		var url = "pop_menu_reg.asp?custcode=<%=custcode%>";
		var name ="pop_menu_reg" ;
		var opt = "width=540, height=302, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}



	function go_menu_view(uid) {
		var url = "pop_menu_view.asp?midx=" + uid;
		var name ="pop_menu_view" ;
		var opt = "width=540, height=358, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}


	function search_cust_dept(str) {
		var frm = document.forms[0];
//		if (str !="") {
//			if (str.indexOf("--") != -1) {
//				alert("사용할 수 없는 문자를 입력하셨습니다.");
//				frm.txtsearchstring.value = "";
//				frm.txtsearchstring.focus();
//				return false;
//			}
//		}
//		frm.txtsearchstring.value = str;
		frm.action = "menu_list.asp";
		frm.method = "post";
		frm.submit();
	}
//-->
</script>