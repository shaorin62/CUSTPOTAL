<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
 <head>
  <title> New Document </title>
  <meta name="Generator" content="EditPlus">
  <meta name="Author" content="">
  <meta name="Keywords" content="">
  <meta name="Description" content="">
 </head>

 <body  oncontextmenu="return false">
  <form enctype="multipart/form-data">
	<input type="text" name="txtfile">
<img src="/images/btn_save.gif" width="59" height="18"  vspace="5" style="cursor:hand" onClick="check_submit();">
  </form>

 </body>
</html>

<script language="JavaScript">
<!--
	function check_submit() {
		var frm = document.forms[0];
		if (frm.txtfile.value == "") {
			alert("첨부할 파일을 선택하세요");
			return false;
		}
		frm.action = "reg_proc.asp";
		frm.mehtod="post";
		frm.submit();
	}
//-->
</script>