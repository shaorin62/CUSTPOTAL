<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
%>��ü�� ����Ʈ ��� _ ��Ȳ

<HTML>
<BODY>
<form name="write_form" enctype="multipart/form-data">
	
	File Upload Size Limit<br><br>
	Select the file to upload  :  
	
	<input type="file" name="file"><br><br>

	1. ���ε� ���� �ӽ� ���� ��δ� C:\TEMP�Դϴ�.<br>
	2.'ã�ƺ���' ��ư�� �������� �ʴ� �������� ��� �ֽ� ������ �������� ������Ʈ �Ͻñ� �ٶ��ϴ�.<br>
	3. �� ���������� ��ü ���� ũ��� �ƴ� ������ ������ 1MB ���Ϸ� �����մϴ�.<br>
	4. ���� ������Ʈ�� Ư�� ������ ������ ��� ���ε� �� �Ŀ� ũ�� ������ �� �� �ֽ��ϴ�.<br>
	&nbsp&nbsp&nbsp ���ε� ������ ������ ũ�⸦ üũ�ϸ� ���� �޽����� ���������� ����� �� �����ϴ�.<br>
	5. ��ü ũ�⸦ �����ϴ� TotalLen �Ӽ��� 50MB�� �����Ǿ� �ֽ��ϴ�.<br>
    &nbsp&nbsp&nbsp ���ε� ������ ũ�⸦ üũ������ �������� ���� �޽����� ������� �ʽ��ϴ�.<br>
    &nbsp&nbsp&nbsp TotalLen�� ũ�⸦ �ʰ��ϴ� �����͸� ���ε� �ϸ� �������� Ư���� �ٽ� �����ؾ� �մϴ�.<br><br>
	
	<img src="/images/btn_save.gif" width="59" height="18"  vspace="5" style="cursor:hand" onClick="check_submit();">
</form>
</BODY>
</HTML>

<script type="text/javascript" src="/js/calendar.js"></script>
<script type="text/javascript" src="/js/script.js"></script>
<script language="javascript">
<!--
	function check_submit() {
		var frm = document.forms[0];
		
		frm.method = "POST";
		frm.action = "report_list_proc.asp";
		frm.submit();
	}

	window.onload = function () {

	}
//-->
</script>