<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<%
	Dim filename : filename = request("filename")
	If request("flag") = "B" Then
		Call MakeUrlToFile("http://mms.raed.co.kr/cust/outdoor/print/prt_b_contact.asp?contidx="&request("contidx")&"&cyear="&request("cyear")&"&cmonth="&request("cmonth"), filename,request("contidx"),request("cyear"),request("cmonth"))
	Else
		Call MakeUrlToFile("http://mms.raed.co.kr/cust/outdoor/print/prt_s_contact.asp?contidx="&request("contidx")&"&cyear="&request("cyear")&"&cmonth="&request("cmonth"), filename,request("contidx"),request("cyear"),request("cmonth"))
	End If
%>
<script type="text/javascript">
<!--
	if (confirm('��Ȳ ������ �����Ͽ����ϴ�.\n������ ��Ȳ ������ �ٿ�ε��Ͻðڽ��ϱ�?')) {
		location.href='/cust/outdoor/process/download.asp?filename=<%=filename%>';
	} else {
		location.href = "/cust/inc/blank.htm";
	}
//-->
</script>


