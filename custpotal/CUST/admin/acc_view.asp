<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
	Dim gotopage : gotopage = request("gotopage")
	dim userid : userid = request.QueryString("userid")

	Dim objrs
	dim sql : sql = "select a.userid, a.password, a.class, a.custcode, c.custname, isuse from dbo.wb_account a inner join dbo.sc_cust_temp c on a.custcode = c.custcode where a.userid='" & userid & "'"
	Call get_recordset(objrs, sql)

	Dim password, class_, custcode, custname, isuse
	if not objrs.eof Then
		userid = objrs("userid")
		password = objrs("password")
		class_ = objrs("class")
		custcode = objrs("custcode")
		custname = objrs("custname")
		isuse = objrs("isuse")
	else
		response.write "<script type='text/javascript'> alert('삭제된 계정이거나 잘못된 계정아이디 입니다.'); location.href='/main.asp';</script>"
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
<!--#include virtual="/cust/top.asp" -->
  <table id="Table_01" width="1240" height="652" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/cust/left_admin_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_admin.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td height="600" align="left" valign="top"><table width="1016" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="24">&nbsp;</td>
          </tr>
          <tr>
            <td height="19" valign="top" class="navigator" >관리모드 &gt; 계정관리 &gt; 계정정보</td>
          </tr>
          <tr>
            <td height="17"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"> <%=userid%> 계정 정보</span></td>
          </tr>
          <tr>
            <td height="27">&nbsp;</td>
          </tr>
          <tr>
            <td class="bdpdd"><table  border="0" cellpadding="0" cellspacing="0">
			<tr>
				<td colspan="2" bgcolor="#cacaca" height="1"></td>
			</tr>
              <tr>
                <td class="hw">아이디</td>
                <td class="bw"> <%=userid%></td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
              <tr>
                <td class="hw">비밀번호</td>
                <td class="bw"> <%=password%></td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
              <tr>
                <td class="hw">접속권한</td>
                <td class="bw"> <%If class_ = "C" Then response.write "일반 사용자" Else response.write "관리자"%></td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
              <tr>
                <td class="hw">사업부</td>
                <td class="bw">  <%=custname%> &nbsp; ( <%=custname%> ) </td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
              <tr>
                <td class="hw">사용여부</td>
                <td class="bw"> <%if ucase(isuse) = "Y" then response.write "사용" else response.write "중지"%></td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
                <tr>
                  <td  height="50" valign="bottom"><a href="/cust/admin/acc_list.asp"><img src="/images/btn_list.gif" width="57" height="20" border="0"></a></td>
                  <td  align="right" valign="bottom"> <img src="/images/btn_edit.gif" width="57" height="18" hspace="10" vspace="5" border="0" class="stylelink" onClick="pop_account_edit();"><img src="/images/btn_delete.gif" width="57" height="18" vspace="5" border="0" class="stylelink" onClick="pop_account_delete();"></td>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td class="bdpdd">&nbsp;</td>
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
	function pop_account_edit() {
		var url = "pop_account_edit.asp?userid=<%=userid%>";
		var name = "pop_account_edit";
		var opt = "width=540, height=328, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open (url, name, opt) ;
	}
//-->
</SCRIPT>