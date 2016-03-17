<!--#include virtual="/inc/getdbcon.asp" -->

<%
	dim gotopage : gotopage = request.QueryString("gotopage")
	if gotopage = "" then gotopage = 1
	dim menuidx : menuidx = request("menuidx")
	dim idx : idx = request.querystring("idx")
	dim custcode : custcode = request("selcustcode")
	dim deptcode : deptcode = request("seldeptcode")

	if idx = "" then Resposne.write "<script>alert('이미 삭제되거나 존재하지 않는 리포트입니다.'); history.back(); </script>"
	dim objrs : set objrs = server.CreateObject("adodb.recordset")
	objrs.activeconnection = application("connectionstring")
	objrs.cursorlocation = aduseclient
	objrs.cursortype = adopenforwardonly
	objrs.locktype = adlockreadonly
	objrs.source = "SELECT IDX, SUBJECT, CONTENTS, FILENAME, EMAIL FROM dbo.WEB_BOARD WHERE IDX = " & idx
	objrs.open

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
                <td width="650"><%=objrs("subject")%>&nbsp;</td>
              </tr>
              <tr>
                <td height="30">리포트 내용</td>
                <td><span class="textline"><%=replace(objrs("contents"), chr(13)&chr(10), "<BR>")%>&nbsp;</span></td>
              </tr>
              <tr>
                <td height="30">&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td height="30">첨부파일</td>
                <td><span class="styleLink" onClick="checkForDownload('<%=objrs("FILENAME")%>');"><%if Not IsNull(objrs("filename"))then response.write "<img src='/images/icon_file.gif' width='15' height='15' align='absmiddle' hspace='5' >" & objrs("filename") end if%></span>&nbsp;</td>
              </tr>
              <tr>
                <td height="30">받는사람(Email)</td>
                <td><%=objrs("email")%>&nbsp;</td>
              </tr>
              <tr>
                <td height="30">&nbsp;</td>
                <td>&nbsp;</td>
              </tr>
              <tr>
                <td height="30" colspan="2" align="right"> <img src="/images/btn_comment_reg.gif" width="78" height="18" border="0" alt="" class="stylelink" onclick="checkForComment();"></td>
              </tr>
              <tr>
                <td height="30" colspan="2" align="right">
				<%
					objrs.close
					objrs.source = "SELECT IDX,COMMENTS, FILENAME, CUSER, CDATE FROM dbo.WEB_BOARD_COMMENT WHERE BOARDIDX=" & idx
					objrs.open

					dim commentsidx, comments, filename, c_user, c_date
					if not objrs.eof then
						set commentsidx = objrs("IDX")
						set comments  = objrs("COMMENTS")
						set filename  = objrs("FILENAME")
						set c_user  = objrs("CUSER")
						set c_date  = objrs("CDATE")
					end if
				%>
				<table border="1">
				<% do until objrs.eof %>
                <tr>
					<td width="100" height="40" align="center"><%=c_user%><input type="hidden" name="txtcommentsidx" value="<%=commentsidx%>"></td>
					<td width="500"><%=replace(comments, chr(13)&chr(10), "<BR>")%> <br><span class="sub"><%=c_date%> &nbsp;<%=formatdatetime(c_date,4)%></span><img src='/images/reply_view_lineone_close.gif' width='11' height='11' align='absmiddle' hspace='5' class="stylelink" onclick="checkForDeleteComment('<%=commentsidx%>','<%=objrs("FILENAME")%>');"><br></td>
					<td width="200"><span class="styleLink" onClick="checkForDownload('<%=objrs("FILENAME")%>');"><img src='/images/ico_attach.gif' width='7' height='12' align='absmiddle' hspace='5' ><%=filename%></span>&nbsp;</td>
                </tr>
				<%
					objrs.movenext
					Loop
				%>
                </table>

				</td>
              </tr>
            </table>
			<table width="800" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="50%" height="50" valign="bottom"><a href="/board/list.asp"><img src="/images/btn_list.gif" width="59" height="20" border="0"></a></td>
                  <td width="50%" align="right" valign="bottom"><img src="/images/btn_edit.gif" width="59" height="20" hspace="10" vspace="5" border="0" class="stylelink" onClick="checkForEdit()"><img src="/images/btn_delete.gif" width="59" height="20" vspace="5" border="0" class="stylelink" onClick="checkForDelete();"></td>
                </tr>
              </table>
			 </td>
          </tr>
          <tr>
            <td class="bdpdd">&nbsp;</td>
          </tr>

      </table></td>
    </tr>
  </table>
  <input type="hidden" name="idx" value="<%=idx%>">
  <input type="hidden" name="menuidx" value="<%=menuidx%>">
  <input type="hidden" name="gotopage" value="<%=gotopage%>">
<!--#include virtual="bottom.asp" -->
  </form>
<%
	objrs.close
	set objrs = nothing
%>
</body>
</html>
<script language="JavaScript">
<!--
	function checkForEdit() {
		if (confirm("리포트 내용을 수정하시겠습니까?")) {
			var frm = document.forms[0];
			frm.action = "/board/edit.asp";
			frm.method = "POST";
			frm.submit();
		}
	}

	function checkForDelete() {
		if (confirm("시스템에서 리포트가 삭제됩니다.\n첨부된 파일도 삭제됩니다.\n\n리포트를 삭제하시겠습니까?")) {
			var frm = document.forms[0];
			frm.action = "/board/delete_proc.asp";
			frm.method = "POST";
			frm.submit();
		}
	}

	function checkForDownload(name) {
		location.href="download.asp?filename="+name;
	}

	function checkForComment() {
		var url = "reg_comments.asp?boardidx=<%=idx%>";
		var name = "WinComment";
		var opt = "width=510, height=200, resizable=no, top=100, left=100";
		window.open(url, name, opt);
	}

	function checkForDeleteComment(commentsidx, filename) {

		if (confirm("선택한 댓글을 삭제하시겠습니까?")) {
			location.href="comments_delete_proc.asp?commentsidx=" + commentsidx+"&idx=<%=idx%>&menuidx=<%=menuidx%>&gotopage=<%=gotopage%>&filename="+filename;
//			var frm = document.forms[0];
//			frm.action = "comments_delete_proc.asp";
//			frm.method = "post";
//			frm.submit();
		}
	}
//-->
</script>