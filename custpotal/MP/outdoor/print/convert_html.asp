<!--#include virtual="/MP/outdoor/inc/Function.asp" -->
<%
	Dim filename : filename = request("filename")
	If request("flag") = "B" Then
		Call MakeUrlToFile("http://mms.raed.co.kr/MP/outdoor/print/prt_b_contact.asp?contidx="&request("contidx")&"&cyear="&request("cyear")&"&cmonth="&request("cmonth"), filename,request("contidx"),request("cyear"),request("cmonth"))
	Else
		Call MakeUrlToFile("http://mms.raed.co.kr/MP/outdoor/print/prt_s_contact.asp?contidx="&request("contidx")&"&cyear="&request("cyear")&"&cmonth="&request("cmonth"), filename,request("contidx"),request("cyear"),request("cmonth"))
	End If
%>
<script type="text/javascript">
<!--
	if (confirm('현황 파일을 생성하였습니다.\n생성한 현황 파일을 다운로드하시겠습니까?')) {
		location.href='/MP/outdoor/process/download.asp?filename=<%=filename%>';
	} else {
		location.href = "/MP/inc/blank.htm";
	}
//-->
</script>


