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
		response.write "<script type='text/javascript'> alert('������ �����̰ų� �߸��� �������̵� �Դϴ�.'); location.href='acc_list.asp?gotopage="&gotopage&";</script>"
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
            <td height="19" valign="top" class="navigator" >������� &gt; �޴����� &gt; �޴���� </td>
          </tr>
          <tr>
            <td height="17"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle">�޴����</span></td>
          </tr>
          <tr>
            <td height="27">&nbsp;</td>
          </tr>
          <tr>
            <td class="bdpdd"><table width="800" border="1" cellpadding="0">
              <tr>
                <td width="150" height="30">�޴���</td>
                <td colspan="3"><%=menuname%>&nbsp; </td>
              </tr>
              <tr>
                <td height="30">�����</td>
                <td ><%=custname%>&nbsp;</td>
              </tr>
              <tr>
                <td height="30">�����޴�</td>
                <td colspan="3"><%=highmenuname%>&nbsp;</td>
              </tr>
              <tr>
                <td  rowspan=3>�޴����</td>
                <td colspan="3" height="30"><% if isfile then response.write "÷������ ���" %>&nbsp; </td>
              </tr>
              <tr>
                <td colspan="3" height="30"><% if isemail then response.write "���Ϲ߼� ���" %></td>
              </tr>
              <tr>
                <td colspan="3" height="30"><% if iscomment then response.write  "����ۼ� ���" %></td>
              </tr>
              <tr>
                <td height="30">��뿩��</td>
                <td colspan="3"><%if ucase(isuse) = "Y" then response.write "���" else response.write "����"%></td>
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
		if (confirm("�޴� ������ �����Ͻðڽ��ϱ�?"))	location.href="mnu_edit.asp?menuidx=<%=menuidx%>";
		return false ;
	}

	function checkForDelete() {
		if (confirm("�ý��ۿ��� ���� �޴������� ��� �����˴ϴ�.\n\�޴��� �����Ͻðڽ��ϱ�?"))	location.href="mnu_delete_proc.asp?menuidx=<%=menuidx%>";
		return false;
	}
//-->
</script>