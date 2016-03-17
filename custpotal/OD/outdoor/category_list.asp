<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim searchstring : searchstring = request("txtsearchstring")
	dim objrs, sql
	sql = "select ggroupidx, ggroupname, mgroupidx, mgroupname, sgroupidx, sgroupname, dgroupidx, dgroupname, mdidx , mdname from dbo.vw_medium_category  where ggroupname like '%" & searchstring &"%' or mgroupname like '%" & searchstring &"%' or sgroupname like '%" & searchstring &"%' or mdname like '%"&searchstring&"%' "

	call get_recordset(objrs, sql)

	dim ggroupidx, ggroupname, mgroupidx, mgroupname, sgroupidx, sgroupname, dgroupidx, dgroupname, mdidx , mdname
	if not objrs.eof then
		set ggroupname = objrs("ggroupname")
		set ggroupidx = objrs("ggroupidx")
		set mgroupname = objrs("mgroupname")
		set mgroupidx = objrs("mgroupidx")
		set sgroupname = objrs("sgroupname")
		set sgroupidx = objrs("sgroupidx")
		set dgroupname = objrs("dgroupname")
		set dgroupidx = objrs("dgroupidx")
		set mdname = objrs("mdname")
		set mdidx = objrs("mdidx")
	end if
	objrs.sort = "ggroupidx, mgroupidx, sgroupidx"
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
  <table id="Table_01" width="1240"  border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/od/left_outdoor_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_outdoor.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td  height="600" align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="19">&nbsp;</td>
          </tr>
          <tr>
            <td height="17"><TABLE  width="100%">
            <TR>
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"> 매체분류 </span></TD>
				<TD width="50%" align="right"><span class="navigator" > 옥외관리 &gt; 매체관리 &gt; 매체분류  </span></TD>
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
                  <td align="left" background="/images/bg_search.gif"><input type="text" name="txtsearchstring" size="25"> <img src="/images/btn_search.gif" width="39" height="20" align="absmiddle" onClick="get_serch();" class="styleLink" ></td>
                  <td align="right" background="/images/bg_search.gif"><!-- <img src="/images/btn_category_reg.gif" width="78" height="18" align="absmiddle" border="0" class="stylelink" onclick="pop_category_reg();"> <img src="/images/btn_category_edit.gif" width="78" height="18" align="absmiddle" border="0" class="stylelink" onclick="pop_category_edit();"> <img src="/images/btn_category_delete.gif" width="78" height="18" align="absmiddle" border="0" class="stylelink" onclick="pop_category_delete();"> --></td>
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
                        <td width="256" align="center" class="header">대분류</td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="256" align="center" >중분류</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="256" align="center">소분류</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="256" align="center" >세분류 </td>
                      </tr>
                  </table></td>
                </tr>
              </table>
                <table width="1024" height="31" border="0" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
				<% do until objrs.eof %>
                  <tr>
                    <td width="256" height="31"  style="padding-left:10px;cursor:hand;" onclick="pop_category_edit('<%=ggroupidx%>', null)" ><%=ggroupname%></td>
                    <td width="3"></td>
                    <td width="256" style="padding-left:10px;cursor:hand;" onclick="pop_category_edit('<%=mgroupidx%>', 1)"> &nbsp; <%=mgroupname%></td>
                    <td width="3"></td>
                    <td width="256" style="padding-left:10px;cursor:hand;" onclick="pop_category_edit('<%=sgroupidx%>', 2)"> &nbsp; <%=sgroupname%></td>
                    <td width="3"></td>
                    <td width="256" style="padding-left:10px;cursor:hand;" onclick="pop_category_edit('<%=dgroupidx%>', 3)"> &nbsp; <%=dgroupname%></td>
                  </tr>
                  <tr>
                    <td height="1" bgcolor="#E7E9E3" colspan="11"></td>
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
<!--#include virtual="bottom.asp" -->
</form>
</body>
</html>
<script language="JavaScript">
<!--

	function get_serch() {
		var frm = document.forms[0];
		frm.action = "category_list.asp";
		frm.method = "post";
		frm.submit() ;
	}

	function pop_category_reg() {
		var url = "pop_category_reg.asp";
		var name = "pop_category_reg";
		var opt = "width=540, height=266, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open(url, name, opt);

	}

	function pop_category_edit(idx, lvl) {
		var url = "pop_category_edit.asp?categoryidx="+idx+"&lvl="+lvl;
		var name = "pop_category_edit";
		var opt = "width=540, height=201, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open(url, name, opt);

	}

	function pop_category_delete() {
		var url = "pop_category_delete.asp";
		var name = "pop_category_delete";
		var opt = "width=540, height=296, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open(url, name, opt);

	}

	window.onload = function () {
	}
//-->
</script>