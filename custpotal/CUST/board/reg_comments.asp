

<%
	dim boardidx : boardidx = request.querystring("boardidx")
%>
<html>
 <head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
  <meta name="Generator" content="EditPlus">
  <meta name="Author" content="">
  <meta name="Keywords" content="">
  <meta name="Description" content="">
 </head>

 <body  oncontextmenu="return false">
<form enctype="multipart/form-data">
  <table width="500" border="1">
  <tr>
	<td width="100">댓글내용</td>
	<td width="400"><textarea name="txtcomments" rows="5" cols="53"></textarea></td>
  </tr>
  <tr>
	<td>첨부파일</td>
	<td><input type="file" name="txtfile" size="40"></td>
  </tr>
  <tr>
	<td height="40" valign="bottom" align="right" colspan="2"><img src="/images/btn_save.gif" width="59" height="20" hspace="10" vspace="5" style="cursor:hand" onClick="checkForSubmit();"><img src="/images/btn_init.gif" width="67" height="20" vspace="5" style="cursor:hand" onClick="checkForReset();"></td>
  </tr>
  </table>
<input type="hidden" name="txtboardidx" value="<%=boardidx%>">
</form>
 </body>
</html>

<script language="JavaScript">
<!--
	function checkForSubmit() {
		var frm = document.forms[0];

		if (frm.txtcomments.value == "") {
				alert("댓글 내용을 입력하세요");
				frm.txtcomments.focus();
				return false;
		}

		frm.method = "POST";
		frm.action = "reg_comments_proc.asp";
		frm.submit();
	}

	function checkForReset() {
		var frm = document.forms[0];
		frm.reset();
		frm.txtcomments.focus();
		return false;
	}
//-->
</script>