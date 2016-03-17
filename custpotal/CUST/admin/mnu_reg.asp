<!--#include virtual="/inc/getdbcon.asp" -->

<%
%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒</title>
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
                <td colspan="3"><input name="txtmenuname" type="text" id="txtaccount" maxlength="12" size="30"> </td>
              </tr>
              <tr>
                <td height="30">사업부</td>
                <td ><input name="txtdeptcode" type="hidden" id="txtdeptcode" value="" readonly>
                  <input name="txtdeptname" type="text" class="kor" id="txtdeptname" value="" size="30" readonly> <img src="/images/btn_find.gif" width="39" height="20" border="0" alt="" align="absmiddle" class="stylelink" onclick="checkForCustomer()"></td>
              </tr>
              <tr>
                <td height="30">상위메뉴</td>
                <td colspan="3"><input type="hidden" name="txthighmenuidx"> <input type="text" name="txthighmenuname" size="30" readonly> <img src="/images/btn_find.gif" width="39" height="20" border="0" alt="" align="absmiddle" class="stylelink" onclick="checkForHighMenu()"></td>
              </tr>
              <tr>
                <td height="30" rowspan=3>메뉴기능</td>
                <td colspan="3"><input type="checkbox" name="chkfile""> 첨부파일 기능</td>
              </tr>
              <tr>
                <td colspan="3"><input type="checkbox" name="chkmail" > 메일발송 기능</td>
              </tr>
              <tr>
                <td colspan="3"><input type="checkbox" name="chkcomment" > 댓글작성 기능</td>
              </tr>
              <tr>
                <td height="30">사용여부</td>
                <td colspan="3"><input name="rdoisuse" type="radio" value="Y" checked>
                  사용
                  <input name="rdoisuse" type="radio" value="N">
                  중지</td>
              </tr>
            </table>
              <table width="800" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="50%" height="50" align="left" valign="bottom"><a href="/admin/mnu_list.asp"><img src="/images/btn_list.gif" width="59" height="20" border="0"></a></td>
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

		if (frm.txtmenuname.value == "") {
			alert("메뉴명을 입력하셔야 합니다.");
			frm.txtmenuname.focus();
			return false;
		}

		frm.method = "POST";
		frm.action = "mnu_reg_proc.asp";
		frm.submit();
	}

	function checkForReset() {
		var frm = document.forms[0];
		frm.reset();
		frm.txtmenuname.focus();
		return false;
	}

	function checkForCustomer() {
		window.open("customer_List2.asp", "Authority", "width=500, height=500, resizable=no, scrollbars=yes, top=100, left=100");
	}

	function checkForHighMenu() {
		var frm = document.forms[0];
		if (frm.txtdeptcode.value == "") {
			alert("사업부를 먼저 선택하세요");
			checkForCustomer();
			return false;
		}
		var url = "highmenu_list.asp?deptcode="+frm.txtdeptcode.value;
		var name = "WinHighMenu";
		var opt = "width=500, height=500, resizable=no, scrollbars=yes, top=100, left=100";
		window.open(url, name, opt);
	}
//-->
</script>