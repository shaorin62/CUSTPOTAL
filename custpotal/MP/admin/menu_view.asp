<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
	dim gotopage : gotopage = request.QueryString("gotopage")
	if gotopage = "" then gotopage = 1
	dim midx : midx = request("midx")

	dim objrs, sql
	sql = "select m2.title , m.title as subtitle, m.custcode as custcode2, m.isfile, m.iscomment, m.isemail,  c2.custcode, c.custname as custname2 , c2.custname, m.isuse, m.lvl  from dbo.wb_menu_mst m inner join dbo.sc_cust_temp c on m.custcode = c.custcode inner join dbo.sc_cust_temp c2 on c.highcustcode = c2.custcode left outer join dbo.wb_menu_mst m2 on m2.midx = m.ref where m.midx = "&midx

	call get_recordset(objrs, sql)

	dim title, custcode, custcode2, file, comment, email, custname, custname2, isuse, lvl, subtitle
	if not objrs.eof then
		title = objrs("title")
		subtitle = objrs("subtitle")
		custcode2 = objrs("custcode2")
		custcode = objrs("custcode")
		custname = objrs("custname")
		custname2 = objrs("custname2")
		file = objrs("isfile")
		comment = objrs("iscomment")
		email = objrs("isemail")
		isuse = objrs("isuse")
		lvl = objrs("lvl")
	else
		response.write "<script type='text/javascript'> alert('삭제된 계정이거나 잘못된 계정아이디 입니다.'); location.href='/main.asp'</script>"
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
<!--#include virtual="/hq/top.asp" -->
  <table id="Table_01" width="1240" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/hq/left_admin_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_admin.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td height="500" align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="24">&nbsp;</td>
          </tr>
          <tr>
            <td height="19" valign="top" class="navigator" >관리모드 &gt; 메뉴관리 <%=custname%> <%if custname <> custname2 then response.write " &gt; " & custname2 %> &gt; <%if title <> subtitle then response.write  title & " &gt; "%> <%=subtitle%>  </td>
          </tr>
          <tr>
            <td height="17"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"><%=custname%> <%if custname <> custname2 then response.write " &gt; " & custname2 %> &gt; <%if title <> subtitle then response.write  title & " &gt; "%> <%=subtitle%> </span></td>
          </tr>
          <tr>
            <td height="27">&nbsp;</td>
          </tr>
          <tr>
            <td class="bdpdd"><table border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
				</tr>
              <tr>
                <td class="tdhd">메뉴명</td>
                <td class="tdbd"><%=subtitle%>&nbsp; </td>
              </tr>
				<tr>
					<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
				</tr>
				<% if lvl = 1 then %>
              <tr>
                <td class="tdhd">광고주</td>
                <td class="tdbd"><%=custname%>&nbsp;</td>
              </tr>
				<tr>
					<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
				</tr>
              <tr>
                <td class="tdhd">사업부</td>
                <td class="tdbd"><%if custname <> custname2 then response.write custname2%>&nbsp;</td>
              </tr>
				<tr>
					<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
				</tr>
				<% end if%>
              <tr>
                <td class="tdhd">파일첨부</td>
                <td class="tdbd"><% if file then response.write "파일 첨부 가능" %>&nbsp; </td>
              </tr>
				<tr>
					<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
				</tr>
              <tr>
                <td class="tdhd">메일발송</td>
                <td class="tdbd"><% if email then response.write "메일 발송 가능" %>&nbsp; </td>
              </tr>
				<tr>
					<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
				</tr>
              <tr>
                <td class="tdhd">댓글기능</td>
                <td class="tdbd"><% if comment then response.write "댓글 작성 가능" %>&nbsp; </td>
              </tr>
				<tr>
					<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
				</tr>
              <tr>
                <td class="tdhd">사용여부</td>
                <td class="tdbd"><%if ucase(isuse) = "Y" then response.write "사용" else response.write "중지"%></td>
              </tr>
				<tr>
					<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
				</tr>
                <tr>
                  <td  height="50" valign="bottom"><a href="/hq/admin/menu_list.asp?selcustcode2=<%=custcode2%>&menunum=<%=request.cookies("menunum")%>"><img src="/images/btn_list.gif" width="59" height="20" border="0"></a></td>
                  <td  align="right" valign="bottom"><%if lvl = 1 then %> <img src="/images/btn_sub_menu_reg.gif" width="78" height="18" border="0" vspace="5" class="stylelink" onclick="pop_sub_menu_reg();"><%end if%> <img src="/images/btn_edit.gif" width="59" height="18" hspace="10" vspace="5" border="0" class="stylelink" onclick="pop_menu_edit();"><img src="/images/btn_delete.gif" width="59" height="18" vspace="5" border="0" class="stylelink" onClick="pop_menu_delete();"></td>
                </tr>
              </table></td>
          </tr>
      </table></td>
    </tr>
  </table>
<!--#include virtual="/bottom.asp" -->
  </form>
</body>
</html>
<SCRIPT LANGUAGE="JavaScript">
<!--
	function pop_menu_edit() {
		if (confirm("메뉴를 수정하시겠습니까?")) {
			<% if lvl = 1 then %>
			var url = "pop_menu_edit.asp?midx=<%=midx%>";
			var name = "pop_menu_edit";
			var opt = "width=540, height=302, resziable=no, scrollbars = no, status=yes, top=100, left=100";
			<% else %>
			var url = "pop_menu_sub_edit.asp?midx=<%=midx%>";
			var name = "pop_menu_sub_edit";
			var opt = "width=540, height=207, resziable=no, scrollbars = no, status=yes, top=100, left=100";
			<%end if%>
			window.open (url, name, opt);
		}
	}

	function pop_sub_menu_reg() {
			var url = "pop_menu_sub_reg.asp?midx=<%=midx%>&custcode=<%=custcode2%>&title=<%=title%>";
			var name = "pop_menu_sub_reg";
			var opt = "width=540, height=207, resziable=no, scrollbars = no, status=yes, top=100, left=100";
			window.open (url, name, opt);
	}

	function pop_menu_delete() {
		if (confirm("선택한 메뉴에 등록된 레포트도 모두 삭제됩니다.\n\n메뉴를 삭제하시겠습니까?")) {
		var url = "menu_delete_proc.asp?midx=<%=midx%>";
		var name = "";
		var opt = "";
		location.href = url ;
		}
	}
//-->
</SCRIPT>