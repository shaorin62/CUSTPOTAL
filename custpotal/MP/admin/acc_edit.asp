<!--#include virtual="/inc/getdbcon.asp" -->
<%
	dim gotopage : gotopage = request.QueryString("gotopage")
	if gotopage = "" then gotopage = 1
	dim uid : uid = request.QueryString("uid")

	dim sql : sql = "DBO.VW_ACCOUNT"

	dim objrs : set objrs = server.createobject("adodb.recordset")
	objrs.activeconnection = application("connectionstring")
	objrs.cursorlocation = aduseclient
	objrs.cursortype = adopenforwardonly
	objrs.locktype = adlockreadonly
	objrs.source = sql
	objrs.open

	if not objrs.eof then
		objrs.find = "USERID = '" & uid &"'"
		dim userid : userid = objrs("USERID")
		dim password : password = objrs("PASSWORD")
		dim classcode : classcode = objrs("CLASS")
		dim authority : authority = objrs("CLASSNAME")
		dim deptcode : deptcode = objrs("DEPTCODE")
		dim deptname : deptname = objrs("DEPTNAME")
		dim custname : custname = objrs("CUSTNAME")
		dim custcode : custcode = objrs("CUSTCODE")
		dim isuse : isuse = objrs("ISUSE")
	else
		response.write "<script type='text/javascript'> alert('������ �����̰ų� �߸��� �������̵� �Դϴ�.'); location.href='acc_list.asp?gotopage="&gotopage&";</script>"
	end if
%>
<html>
<head>
<title>�Ƣ� SK M&C | Media Management System �Ƣ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="../style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  oncontextmenu="return false">
<form>
<!--#include virtual="/hq/top.asp" -->
  <table id="Table_01" width="1240" height="652" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/hq/left_admin_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_admin.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td height="600" align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="24">&nbsp;</td>
          </tr>
          <tr>
            <td height="19" valign="top" class="navigator" >������� &gt; �������� &gt; �������� </td>
          </tr>
          <tr>
            <td height="17"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle">��������</span></td>
          </tr>
          <tr>
            <td height="27">&nbsp;</td>
          </tr>
          <tr>
            <td class="bdpdd"><table width="800" border="1" cellpadding="0">
              <tr>
                <td width="150" height="30">���̵�</td>
                <td colspan="3"><%=userid%><input type="hidden" name="txtaccount" value="<%=userid%>"></td>
              </tr>
              <tr>
                <td height="30">��й�ȣ</td>
                <td width="250"><input name="txtpassword" type="text" class="kor" id="txtpassword" value="" maxlength="12"></td>
                <td width="150">��й�ȣȮ��</td>
                <td width="250"><input name="txtrepassword" type="text" class="kor" id="txtrepassword" value="" maxlength="12"></td>
              </tr>
              <tr>
                <td height="30">���ӱ���</td>
                <td colspan="3"><input name="rdoauthority" type="radio" value="A" onclick="checkForCustomer(this.value)" <%If classcode="A" Then response.write "checked"%> >
                  ������
                    <input name="rdoauthority" type="radio" value="C"  onclick="checkForCustomer(this.value)" <%If classcode="C" Then response.write "checked"%> >
                  �Ϲݻ����</td>
              </tr>
              <tr>
                <td height="30">�����</td>
                <td colspan="3"><input name="txtdeptcode" type="hidden" id="txtdeptcode" value="<%=deptcode%>" readonly>
                  <input name="txtdeptname" type="text" class="kor" id="txtdeptname"  size="30" value="<%=deptname%>" readonly> <img src="/images/btn_search.gif" width="39" height="20" border="0" alt="" align="absmiddle" class="stylelink" onclick="checkForCustomer()">  </td>
              </tr>
              <tr>
                <td height="30">��뿩��</td>
                <td colspan="3"><input name="rdoisuse" type="radio" value="Y"  <%if ucase(isuse) = "Y" then response.write "checked" %>>
                  ���
                  <input name="rdoisuse" type="radio" value="N" <%if ucase(isuse) = "N" then response.write "checked" %>>
                  ����</td>
              </tr>
            </table>
              <table width="800" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="50%" height="50" align="left" valign="bottom"><a href="/admin/acc_list.asp"><img src="/images/btn_list.gif" width="59" height="20" border="0"></a></td>
                  <td width="50%" align="right" valign="bottom"><img src="/images/btn_save.gif" width="59" height="20" hspace="10" vspace="5" style="cursor:hand" onClick="checkForSubmit();"><img src="/images/btn_init.gif" width="67" height="20" vspace="5" style="cursor:hand" onClick="checkForReset();"></td>
                </tr>
              </table></td>
          </tr>
          <tr>
            <td class="bdpdd">&nbsp;</td>
          </tr>

      </table></td>
    </tr>
  </table>
  <input type="hidden" name="gotopage" value="<%=gotopage%>">
<!--#include virtual="/bottom.asp" -->
  </form>
</body>
</html>
<script language="JavaScript">
<!--
	function checkForSubmit() {
		var frm = document.forms[0];

		if (frm.txtpassword.value != frm.txtrepassword.value) {
			alert("��й�ȣ�� �߸��Է��ϼ̽��ϴ�.\n\n��й�ȣ�� ��Ȯ�ϰ� �Է��ϼž� �մϴ�.");
			frm.txtrepassword.value = "";
			frm.txtrepassword.focus();
			return false;
		}

		frm.method = "POST";
		frm.action = "acc_edit_proc.asp";
		frm.submit();
	}

	function checkForReset() {
		var frm = document.forms[0];
		frm.reset();
		frm.txtaccount.focus();
		return false;
	}

//-->
</script>
<script type="text/JavaScript" src="/js/admin.js"></script>
<script type="text/JavaScript" src="/js/menu.js"></script>