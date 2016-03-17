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
		response.write "<script type='text/javascript'> alert('삭제된 계정이거나 잘못된 계정아이디 입니다.'); location.href='acc_list.asp?gotopage="&gotopage&";</script>"
	end if
%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒</title>
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
            <td height="19" valign="top" class="navigator" >관리모드 &gt; 계정관리 &gt; 계정변경 </td>
          </tr>
          <tr>
            <td height="17"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle">계정변경</span></td>
          </tr>
          <tr>
            <td height="27">&nbsp;</td>
          </tr>
          <tr>
            <td class="bdpdd"><table width="800" border="1" cellpadding="0">
              <tr>
                <td width="150" height="30">아이디</td>
                <td colspan="3"><%=userid%><input type="hidden" name="txtaccount" value="<%=userid%>"></td>
              </tr>
              <tr>
                <td height="30">비밀번호</td>
                <td width="250"><input name="txtpassword" type="text" class="kor" id="txtpassword" value="" maxlength="12"></td>
                <td width="150">비밀번호확인</td>
                <td width="250"><input name="txtrepassword" type="text" class="kor" id="txtrepassword" value="" maxlength="12"></td>
              </tr>
              <tr>
                <td height="30">접속권한</td>
                <td colspan="3"><input name="rdoauthority" type="radio" value="A" onclick="checkForCustomer(this.value)" <%If classcode="A" Then response.write "checked"%> >
                  관리자
                    <input name="rdoauthority" type="radio" value="C"  onclick="checkForCustomer(this.value)" <%If classcode="C" Then response.write "checked"%> >
                  일반사용자</td>
              </tr>
              <tr>
                <td height="30">사업부</td>
                <td colspan="3"><input name="txtdeptcode" type="hidden" id="txtdeptcode" value="<%=deptcode%>" readonly>
                  <input name="txtdeptname" type="text" class="kor" id="txtdeptname"  size="30" value="<%=deptname%>" readonly> <img src="/images/btn_search.gif" width="39" height="20" border="0" alt="" align="absmiddle" class="stylelink" onclick="checkForCustomer()">  </td>
              </tr>
              <tr>
                <td height="30">사용여부</td>
                <td colspan="3"><input name="rdoisuse" type="radio" value="Y"  <%if ucase(isuse) = "Y" then response.write "checked" %>>
                  사용
                  <input name="rdoisuse" type="radio" value="N" <%if ucase(isuse) = "N" then response.write "checked" %>>
                  중지</td>
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
			alert("비밀번호를 잘못입력하셨습니다.\n\n비밀번호를 정확하게 입력하셔야 합니다.");
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