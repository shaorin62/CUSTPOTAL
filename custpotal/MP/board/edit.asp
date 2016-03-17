<!--#include virtual="/inc/getdbcon.asp" -->

<%
	dim gotopage : gotopage = request.form("gotopage")
	if gotopage = "" then gotopage = 1
	dim menuidx : menuidx = request("menuidx")
	dim idx : idx = request.form("idx")
	dim custcode : custcode = request("selcustcode")
	dim deptcode : deptcode = request("seldeptcode")


	if idx = "" then Response.write "<script>alert('이미 삭제되거나 존재하지 않는 리포트입니다.'); location.href='list.asp?gotopage="&gotopage&"';</script>"
	dim objrs : set objrs = server.CreateObject("adodb.recordset")
	objrs.activeconnection = application("connectionstring")
	objrs.cursorlocation = aduseclient
	objrs.cursortype = adopenforwardonly
	objrs.locktype = adlockreadonly
	'objrs.source = "SELECT IDX, SUBJECT, CONTENTS, FILENAME, EMAIL FROM dbo.WEB_BOARD WHERE IDX = " & idx
	objrs.source = "dbo.WEB_BOARD"
	objrs.open

	objrs.Find= "IDX=" & idx

%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="../style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  oncontextmenu="return false">
<form enctype="multipart/form-data">
<!--#include virtual="/hq/top.asp" -->
  <table id="Table_01" width="1240" height="652" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/hq/left_report_menu.asp" --></td>
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
                <td width="650"><input name="txtsubject" type="text" class="kor" id="txtsubject" size="50" maxlength="50" value="<%=objrs("SUBJECT")%>"></td>
              </tr>
              <tr>
                <td height="30">리포트 내용</td>
                <td><textarea name="txtcontents" cols="70" rows="10" class="kor" id="txtcontents"><%=objrs("CONTENTS")%></textarea></td>
              </tr>
              <tr>
                <td height="30">&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td height="30">첨부파일</td>
                <td><input name="txtfile" type="file" id="txtfile" size="50"> (등록된 파일 : <%=objrs("FILENAME")%>)</td>
              </tr>
              <tr>
                <td height="30">받는사람(Email)</td>
                <td><input name="txtemail" type="text" id="txtemail" size="50" maxlength="200" value="<%=objrs("EMAIL")%>">
                  (한명 이상인 경우 ;로 구분)</td>
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
  <input type="hidden" name="idx" value="<%=idx%>">
  <input type="hidden" name="menuidx" value="<%=menuidx%>">
  <input type="hidden" name="gotopage" value="<%=gotopage%>">
  <input type="hidden" name="txtattach" value="<%=objrs("FILENAME")%>">
</form>
<%
	objrs.close
	set objrs = nothing
%>
</body>
</html>
<script language="JavaScript">
<!--
	function checkForSubmit() {
		var frm = document.forms[0];

		frm.method = "POST";
		frm.action = "edit_proc.asp";
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