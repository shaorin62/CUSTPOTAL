
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%
'Option Explicit
Dim inet
Dim url, str, worldPop
Dim iStart, iEnd

url="http://10.110.10.86:6666/pds/test.pptx"

'① inet 컨트롤의 인스턴스를 생성한다.
set inet = CreateObject("InetCtls.Inet")
response.write IsObject(inet)
response.end

'② Internet Transfer Control을 통해서 지정된 URL의 소스를 가져온다
inet.RequestTimeOut = 20 
inet.Url = url
str =inet.OpenURL

'③ <H1>..</H1>사이의 값을 뽑아내기
'iStart =instr(str, "<H1>") '<H1>태그가 시작하는 위치의 값 
'iEnd = instr(str, "</H1>") '</H1>태그가 시작하는 위치의 값 
Response.write str
response.end
'worldPop = Mid(str,iStart,iEnd-iStart)
'%>
<!-- <HTML>
<HEAD>
<BODY>
<center>현재 세계 총 인구 (<%'=Now%>) : <p><%'=worldPop %> 명</P>
</BODY>
</HTML>  -->
<%
'Session.Codepage=65001
'Response.ContentType="text/HTML"
'Response.Charset="utf-8"
'url= "http://cafe.daum.net/ljo3527922/xiY/739?docid=Qx8D|xiY|739|20090610192459&q=%BD%C3%B1%B9&srchid=CCBQx8D|xiY|739|20090610192459"
'
'Set xml = server.CreateObject("Microsoft.XMLHTTP") '개체생성
'
'xml.open "POST", "" & url & "", false '원하는 url불러오기
'xml.send "" '실행
'
'strStatus = xml.Status '실행상태 받아오기
'str = xml.responseText '실제 받고자하는 데이터
'response.write InStr(str, "조회수")
'
'Set xml = Nothing '개체 소멸

'Response.Write str '결과 출력
%>
