<%@ language="vbscript" codepage="65001" %>
<!--#include virtual="/inc/func.asp" -->

<%
	dim custcode : custcode = request("custcode2")
	if custcode = "" then custcode = null

	dim objrs, sql
	if not isnull(custcode) then
		sql = "select m.midx, m.title, c.custname, m.lvl, m.isfile, iscomment, isemail from dbo.wb_menu_mst m inner join dbo.sc_cust_temp c on m.custcode = c.custcode where m.custcode = '"&custcode&"' order by  m.ref  , m.lvl"
	else
		sql = "select m.midx, m.title, c.custname, m.lvl, m.isfile, iscomment, isemail from dbo.wb_menu_mst m left outer join dbo.sc_cust_temp c on m.custcode = c.custcode where m.custcode is null order by  m.ref , m.lvl"
	end if
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
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<table width="1030" height="31" border="3" cellpadding="0" cellspacing="0" bordercolor="#8D652B" align="center">
                <tr>
                  <td><table width="1024" border="0" cellspacing="0" cellpadding="0" class="header" align="center">
                      <tr>
                        <td width="624" align="center" >메뉴명</td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="100" align="center" >첨부파일</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="100" align="center" >메일발송</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="100" align="center" >댓글기능</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="110" align="center" >하위메뉴</td>
                      </tr>
                  </table></td>
                </tr>
              </table>
                <table width="1024"  border="0" cellpadding="0" cellspacing="0" bordercolor="#8D652B" style="margin-left:3px;">
				<% do until objrs.eof %>
                  <tr class="styleLink" height="31">
                    <td width="624" align="left"  class="styleLink" style="padding-left:20px;" onClick="go_menu_view('<%=midx%>')" ><%if lvl = 2 then %><img src="/images/tree-branch.gif" width="19" height="14" border="0" alt="" hspace="5"> <%end if%><%=title%>&nbsp;</td>
                    <td width="3" align="center">&nbsp;</td>
                    <td width="100" align="center"><%if isfile then response.write "사용"%>&nbsp;</td>
                    <td width="3">&nbsp;</td>
                    <td width="100" align="center"><%if isemail then response.write "사용"%>&nbsp;</td>
                    <td width="3">&nbsp;</td>
                    <td width="100" align="center"><%if iscomment then response.write "사용"%>&nbsp;</td>
                    <td width="3">&nbsp;</td>
                    <td width="110" align="center" onClick="go_submenu_reg('<%=midx%>')" ><%if lvl = 1 then %><img src="/images/btn_submeun_reg.gif" width="100" height="18" border="0" alt=""><%end if%></td>
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
            </table>

<script language="JavaScript">
<!--

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