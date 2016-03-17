<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
	dim menuidx : menuidx = request.querystring("menuidx")
	dim gotopage : gotopage = request.QueryString("gotopage")
	if gotopage = "" then gotopage = 1

	dim sql : sql = " SELECT M.MENUIDX, M.MENUNAME, CUSTNAME, M.ISFILE, M.ISEMAIL, M.ISCOMMENT, M.ISUSE, M2.MENUNAME AS HIGHMENUNAME " &_
						" FROM DBO.WEB_BOARD_MENU M LEFT OUTER JOIN DBO.SC_CUST_TEMP C ON M.CUSTCODE = C.CUSTCODE " &_
						" LEFT OUTER JOIN DBO.WEB_BOARD_MENU M2 ON M.HIGHMENUIDX = M2.MENUIDX " &_
						" WHERE M.MENUIDX = " & menuidx

	dim objrs : set objrs = server.createobject("adodb.recordset")
	objrs.activeconnection = application("connectionstring")
	objrs.cursorlocation = aduseclient
	objrs.cursortype = adopenforwardonly
	objrs.locktype = adlockreadonly
	objrs.source = sql
	objrs.open

	objrs.find = "MENUIDX = '" & menuidx &"'"

	dim menuname, custname, isfile, isemail, iscomment, isuse, highmenuname
	if not objrs.eof then
		menuname = objrs("MENUNAME")
		custname = objrs("CUSTNAME")
		isfile = objrs("ISFILE")
		isemail = objrs("ISEMAIL")
		iscomment = objrs("ISCOMMENT")
		isuse = objrs("ISUSE")
		highmenuname = objrs("HIGHMENUNAME")
	else
		response.write "<script type='text/javascript'> alert('삭제된 계정이거나 잘못된 계정아이디 입니다.'); location.href='acc_list.asp?gotopage="&gotopage&";</script>"
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
            <td height="19" valign="top" class="navigator" >관리모드 &gt; 메뉴관리 &gt; 메뉴등록 </td>
          </tr>
          <tr>
            <td height="17"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle">메뉴등록</span></td>
          </tr>
          <tr>
            <td height="27">&nbsp;</td>
          </tr>
          <tr>
            <td class="bdpdd"><table width="800" border="1" cellpadding="0">
              <tr>
                <td width="150" height="30">메뉴명</td>
                <td colspan="3"><%=menuname%>&nbsp; </td>
              </tr>
              <tr>
                <td height="30">사업부</td>
                <td ><%=custname%>&nbsp;</td>
              </tr>
              <tr>
                <td height="30">상위메뉴</td>
                <td colspan="3"><%=highmenuname%>&nbsp;</td>
              </tr>
              <tr>
                <td  rowspan=3>메뉴기능</td>
                <td colspan="3" height="30"><% if isfile then response.write "첨부파일 기능" %>&nbsp; </td>
              </tr>
              <tr>
                <td colspan="3" height="30"><% if isemail then response.write "메일발송 기능" %></td>
              </tr>
              <tr>
                <td colspan="3" height="30"><% if iscomment then response.write  "댓글작성 기능" %></td>
              </tr>
              <tr>
                <td height="30">사용여부</td>
                <td colspan="3"><%if ucase(isuse) = "Y" then response.write "사용" else response.write "중지"%></td>
              </tr>
            </table>
			<table width="800" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="50%" height="50" valign="bottom"><a href="/hq/admin/mnu_list.asp"><img src="/images/btn_list.gif" width="59" height="20" border="0"></a></td>
                  <td width="50%" align="right" valign="bottom"> <a href="/admin/mnu_reg.asp"><img src="/images/btn_reg.gif" width="58" height="20" alt="" border="0" vspace="5"></a> <img src="/images/btn_edit.gif" width="59" height="20" hspace="10" vspace="5" border="0" class="stylelink" onClick="checkForEdit()"><img src="/images/btn_delete.gif" width="59" height="20" vspace="5" border="0" class="stylelink" onClick="checkForDelete();"></td>
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
<script type="text/javascript" src="/js/admin.js"></script>
<script language="JavaScript">
<!--
	function checkForEdit() {
		if (confirm("메뉴 정보를 수정하시겠습니까?"))	location.href="mnu_edit.asp?menuidx=<%=menuidx%>";
		return false ;
	}

	function checkForDelete() {
		if (confirm("시스템에서 사용된 메뉴정보가 모두 삭제됩니다.\n\메뉴를 삭제하시겠습니까?"))	location.href="mnu_delete_proc.asp?menuidx=<%=menuidx%>";
		return false;
	}
//-->
</script>