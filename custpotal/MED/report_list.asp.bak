<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
%>매체사 리포트 등록 _ 현황

<HTML>
<BODY>
<form name="write_form" enctype="multipart/form-data">
	
	File Upload Size Limit<br><br>
	Select the file to upload  :  
	
	<input type="file" name="file"><br><br>

	1. 업로드 파일 임시 저장 경로는 C:\TEMP입니다.<br>
	2.'찾아보기' 버튼을 지원하지 않는 브라우저인 경우 최신 버전의 브라우저로 업데이트 하시기 바랍니다.<br>
	3. 이 예제에서는 전체 파일 크기기 아닌 각각의 파일을 1MB 이하로 제한합니다.<br>
	4. 서버 컴포넌트의 특성 때문에 파일을 모두 업로드 한 후에 크기 제한을 할 수 있습니다.<br>
	&nbsp&nbsp&nbsp 업로드 이전에 파일의 크기를 체크하면 에러 메시지를 정상적으로 출력할 수 없습니다.<br>
	5. 전체 크기를 제한하는 TotalLen 속성은 50MB로 설정되어 있습니다.<br>
    &nbsp&nbsp&nbsp 업로드 이전에 크기를 체크하지만 정상적인 에러 메시지를 출력하지 않습니다.<br>
    &nbsp&nbsp&nbsp TotalLen의 크기를 초과하는 데이터를 업로드 하면 브라우저의 특성상 다시 실행해야 합니다.<br><br>
	
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