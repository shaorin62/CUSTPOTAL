<!--#include virtual="/inc/getdbcon.asp" -->
<%
%>
<html>
<head>
<title>�Ƣ� SK M&C | Media Management System �Ƣ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="../style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  oncontextmenu="return false">
<form>
<!--#include virtual="/cust/top.asp" -->
  <table id="Table_01" width="1240" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/cust/left_admin_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_admin.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td height="500" align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="24">&nbsp;</td>
          </tr>
          <tr>
            <td height="19" valign="top" class="navigator" >������� &gt; �������� &gt; ������� </td>
          </tr>
          <tr>
            <td height="17"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle">�������</span></td>
          </tr>
          <tr>
            <td height="27">&nbsp;</td>
          </tr>
          <tr>
            <td class="bdpdd"><table width="800" border="1" cellpadding="0">
              <tr>
                <td width="150" height="30">���̵�</td>
                <td colspan="3"><input name="txtaccount" type="text" id="txtaccount" maxlength="12" onkeyup="checkForKey()"> <img src="/images/btn_overlap.gif" width="57" height="20" border="0" alt="" align="absmiddle" class="stylelink" onclick="checkForAccount()"></span><INPUT TYPE="hidden" NAME="checkAccount" value="N"></td>
              </tr>
              <tr>
                <td height="30">��й�ȣ</td>
                <td width="250"><input name="txtpassword" type="password"  id="txtpassword" value="" maxlength="12"></td>
                <td width="150">��й�ȣȮ��</td>
                <td width="250"><input name="txtrepassword" type="password" id="txtrepassword" value="" maxlength="12"></td>
              </tr>
              <tr>
                <td height="30">���ӱ���</td>
                <td colspan="3"><input name="rdoauthority" type="radio" value="A" >
                  ������
                    <input name="rdoauthority" type="radio" value="C"  >
                  �Ϲݻ����</td>
              </tr>
              <tr>
                <td height="30">����μ�</td>
                <td colspan="3"><input name="txtdeptcode" type="hidden" id="txtdeptcode" value="" readonly>
                  <input name="txtdeptname" type="text" class="kor" id="txtdeptname" value="" size="30" readonly> <img src="/images/btn_search.gif" width="39" height="20" border="0" alt="" align="absmiddle" class="stylelink" onclick="checkForCustomer()"> </td>
              </tr>
              <tr>
                <td height="30">��뿩��</td>
                <td colspan="3"><input name="rdoisuse" type="radio" value="Y" checked>
                  ���
                  <input name="rdoisuse" type="radio" value="N">
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
	window.onload = function init() {
	}

	function checkForSubmit() {
		var frm = document.forms[0];
		var flag = true ;

		if (frm.txtaccount.value == "") {
			alert("������ �Է��ϼž� �մϴ�.");
			frm.txtaccount.focus();
			return false;
		}
		if (frm.checkAccount.value == "N") {
			alert("���̵� �ߺ��˻縦 ���ϼ̽��ϴ�.");
			checkForAccount();
			return false;
		}
		if (frm.txtpassword.value == "" || frm.txtpassword.value.length < 4) {
			alert("��й�ȣ�� ����, ������������ 4~12�ڸ��� �Է��ϼž� �մϴ�.");
			frm.txtpassword.focus();
			return false;
		}
		if (frm.txtpassword.value != frm.txtrepassword.value) {
			alert("��й�ȣ�� �߸��Է��ϼ̽��ϴ�.\n\n��й�ȣ�� ��Ȯ�ϰ� �Է��ϼž� �մϴ�.");
			frm.txtrepassword.value = "";
			frm.txtrepassword.focus();
			return false;
		}

		for (var i = 0 ; i < frm.rdoauthority.length; i++ ) {
			if (frm.rdoauthority[i].checked) flag = false ;
		}

		if (flag) {
			alert("������ ������ �����ϼž� �մϴ�.");
			frm.rdoauthority[0].focus();
			return false;
		}
		frm.method = "POST";
		frm.action = "acc_reg_proc.asp";
		frm.submit();
	}

	function checkForReset() {
		var frm = document.forms[0];
		frm.reset();
		frm.txtaccount.focus();
		return false;
	}

	function checkForAccount() {
		var userid = document.forms[0].txtaccount;
		if (userid.value == "") {
			alert("���̵� ���� �Է��ϼž� �մϴ�.");
			userid.focus();
			return false;
		}
		window.open("checkAccount.asp?txtaccount="+userid.value,"CheckLog", "width=400, height=300, resizable=no, top =100, left=100");
	}
//-->
</script>