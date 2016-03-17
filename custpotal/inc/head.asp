<%
'페이지에서 에러가 발생하여도 페이지 오류를 외부로 출력하지 않기위해 사용
On Error Resume Next
'On Error GoTo 0도 가능하나 2003에서는 실행되지 않음
if err.number <> 0 then
	'Response.Write err.description & "<BR>" & err.source & "<BR>"
	err.clear
End if
%>
