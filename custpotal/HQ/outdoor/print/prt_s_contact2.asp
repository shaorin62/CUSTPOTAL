<!--#include virtual="/hq/outdoor/inc/Function.asp" -->
<object id="factory" style="display:none" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/hq/outdoor/inc/ScriptX.cab#Version=6,1,431,2">
</object>
<%	
	Dim pcontidx : pcontidx = request("contidx")
	Dim pcyear : pcyear = request("cyear")
	Dim pcmonth : pcmonth = request("cmonth")


'	response.write pcontidx
	' ���� ��� ���� ���� 
	Dim sql : sql = "select c.title, c.comment, c.mediummemo, c.regionmemo,  t.highcustcode, c.startdate, c.enddate  "
	sql = sql & " from wb_contact_mst c "
	sql = sql & "  left outer  join sc_cust_dtl t on c.custcode = t.custcode "
	sql = sql & "  left outer  join vw_contact_exe_monthly e on c.contidx = e.contidx and e.cyear = '" & pcyear & "' and e.cmonth = '" & pcmonth & "' "
	sql = sql & " where c.contidx = " & pcontidx 
'	response.write sql
	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType =adCmdText
	Dim rs : Set rs = cmd.execute
	Set cmd = Nothing 
	If Not rs.eof Then 
		Dim title : title = rs("title")
		Dim comment : comment = rs("comment")
		Dim mediummemo : mediummemo = rs("mediummemo")
		Dim regionmemo : regionmemo = rs("regionmemo")
		Dim startdate : startdate = rs("startdate")
		Dim enddate : enddate = rs("enddate")
		If Not IsNull(comment) Then comment = Replace(comment, Chr(13)&Chr(10), "<br>")
		If Not IsNull(mediummemo) Then  mediummemo= Replace(mediummemo, Chr(13)&Chr(10), "<br>")
		If Not IsNull(regionmemo) Then  regionmemo= Replace(regionmemo, Chr(13)&Chr(10), "<br>")
	End If 

%>
<html>
<head>
<title>�Ƣ� SK M&C | Media Management System �Ƣ� </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="http://10.110.10.86:6666/hq/outdoor/style.css" rel="stylesheet" type="text/css">
<script defer>
function PrintTest() {
  var factory = document.getElementById("factory");	
  factory.printing.header = "";   // Header�� �� ����
  factory.printing.footer = "";   // Footer�� �� ����
  factory.printing.portrait = false   // true �� �����μ�, false �� ���� �μ�
  factory.printing.leftMargin = 1.0   // ���� ���� ������
  factory.printing.topMargin = 1.0   // �� ���� ������
  factory.printing.rightMargin = 1.0  // ������ ���� ������
  factory.printing.bottomMargin = 0.1  // �Ʒ� ���� ������
//  factory.printing.SetMarginMeasure(2); // �׵θ� ���� ������ ������ ��ġ�� �����մϴ�.
 // factory.printing.printer = "HP DeskJet 870C";  // ����Ʈ �� ������ �̸�
//  factory.printing.paperSize = "A4";   // ���� ������
  factory.printing.paperSource = "Manual feed";   // ���� Feed ���
//  factory.printing.collate = true;   //  ������� ����ϱ�
//  factory.printing.copies = 2;   // �μ��� �ż�
//  factory.printing.SetPageRange(false, 1, 3); // True�� �����ϰ� 1, 3�̸� 1���������� 3���������� ���
  factory.printing.Print(true) // ����ϱ�
}
</script>
<script type="text/javascript">
<!--
	window.onload = function () {
		PrintTest();
		//self.focus();
		//this.print();
		//this.close();
	}
//-->
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  >
<!-- ��� ��� �̹��� -->
<table width="1240"  class="title" align="center">
	<tr>
		<td><img src="/images/pop_top.gif" width="1240" height="60" align="absmiddle"></td>
	</tr>
</table>
<!-- // ��� ��� �̹��� -->
<table width="1024"   align="center" style="margin-top:30px;">
	<tr>
		<td class="title"><img src="http://10.110.10.86:6666/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><%=title%> </td>
	</tr>
</table>
<% server.execute("/hq/outdoor/print/prt_contactsummary_s.asp") %>
<% server.execute("/hq/outdoor/print/prt_contactdetail_s.asp") %>
<% server.execute("/hq/outdoor/print/prt_reportphoto.asp") %>
<table width="1024" align="center" style="margin-top:10px;">
	<tr>
	  <th class="title" width='100' >��üƯ��</td>
	  <td width='684'  class="context" style="font-family:���� ���;font-size:9px"><%=mediummemo%></td>
	</tr>
	<tr>
	  <th class="title" >����Ư��</td>
	  <td  class="context" style="font-family:���� ���;font-size:9px"><%=regionmemo %></td>
	</tr>
	<tr>
	  <th class="title" >Ư�̻���</td>
	  <td  class="context" style="font-family:���� ���;font-size:9px"><%=comment%></td>
	</tr>
</table>
</body>
</html>
