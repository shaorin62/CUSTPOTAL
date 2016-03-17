<!--#include virtual="/inc/getdbcon.asp" -->

<%
	dim gotopage : gotopage = request.QueryString("gotopage")
	if gotopage = "" then gotopage = 1
	dim midx : midx = request("menuidx")
	dim custcode : custcode = request("selcustcode")
	dim deptcode : deptcode = request("seldeptcode")
%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="../style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  oncontextmenu="return false">
<form enctype="multipart/form-data">
<!--#include virtual="/cust/top.asp" -->
  <table id="Table_01" width="1240" height="652" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"></td>
      <td height="65"><img src="/images/default_03.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td height="600" align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="24">&nbsp;</td>
          </tr>
          <tr>
            <td height="19" valign="top" class="navigator" >매체별 리포트 &gt; 리포트 작성</td>
          </tr>
          <tr>
            <td height="17"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle">리포트 작성</span></td>
          </tr>
          <tr>
            <td height="27">&nbsp;</td>
          </tr>
          <tr>
            <td class="bdpdd"><table width="800" border="1" cellpadding="0">
              <tr>
                <td width="150" height="30">리포트 제목</td>
                <td width="650"><input name="txtsubject" type="text" class="kor" id="txtsubject" size="50" maxlength="50">
				  </td>
              </tr>
              <tr>
                <td height="30">리포트 내용</td>
                <td><textarea name="txtcontents" cols="70" rows="10" class="kor" id="txtcontents"></textarea></td>
              </tr>
              <tr>
                <td height="30">&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td height="30">첨부파일</td>
                <td><input name="txtfile" type="file" id="txtfile" size="50"></td>
              </tr>
              <tr>
                <td height="30">받는사람(Email)</td>
                <td><input name="txtemail" type="text" id="txtemail" size="50" maxlength="200">
                  (한명 이상인 경우 ,로 구분)</td>
              </tr>
              <tr>
                <td height="30">&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
            </table>
              <table width="800" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="50%" height="50" align="left" valign="bottom"><a href="/board/list.asp"><img src="/images/btn_list.gif" width="59" height="20" border="0"></a></td>
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
<!--#include virtual="bottom.asp" -->
  </form>
</body>
</html>
<script language="JavaScript">
<!--
	function checkForSubmit() {
		var frm = document.forms[0];

		if (frm.txtsubject.value == "") {
			if (!confirm("리포트 제목이 입력되지 않았습니다.\n\n제목없이 저장하시겠습니까?")) {
				frm.txtsubject.focus();
				return false;
			}
		}

		frm.method = "POST";
		frm.action = "reg_proc.asp";
		frm.submit();
	}

	function checkForReset() {
		var frm = document.forms[0];
		frm.reset();
		frm.txtsubject.focus();
		return false;
	}
//-->
</script>