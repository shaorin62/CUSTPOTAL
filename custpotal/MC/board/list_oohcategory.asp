<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%

%>
<html>
<head>
<title>�Ƣ� SK M&C | Media Management System �Ƣ� </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/hq/outdoor/style.css" rel="stylesheet" type="text/css">
<script type='text/javascript' src='/js/ajax.js'></script>
<script type='text/javascript' src='/js/script.js'></script>
<script type="text/javascript">
<!--
		var rows = 0;

	

-->
</script>
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  >
<form action="list_contact.asp" method='post'>
<INPUT TYPE="hidden" NAME="menunum" value="11">
<!--#include virtual="/mc/top.asp" -->
  <table width="1240" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="Table_01">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/mc/left_report_menu.asp"--></td>
      <td height="65" valign="top"><img src="/images/middle_navigater_outdoor.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td align="left" valign="top" colspan='2'>
	  <table width="1030" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="19" colspan='2'>&nbsp;</td>
          </tr>
          <tr>
            <td height="17" colspan='2'><TABLE  width="100%">
            <TR>
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="boldtitle"> ���ܱ��� ������Ȳ </span></TD>
				<TD width="50%" align="right"><span class="navigator" > ���ܱ��� &gt;  ���ܱ�����Ȳ &gt; ���ܱ��� ������Ȳ </span></TD>
            </TR>
            </TABLE>
			</td>
          </tr>
          <tr>
            <td height="15" colspan='2'>&nbsp;</td>
          </tr>
          <tr>
            <td colspan='2'>
			<table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13" valign="top" ><img src="/images/bg_search_left.gif" width="13" height="35" ></td>
                  <td background="/images/bg_search.gif"> <img src="/images/icon_dot_search.gif" width="4" height="3" align="absmiddle"> �˻����
				   &nbsp;  <span id="custcode">������ �˻�</span> <span id="teamcode">��� �˻�</span> ��ü��:<input type="text" name="medname" width="100"> <input type="image" src="/images/btn_search.gif" width="39" height="20" align="absmiddle" border=0></td>
                  <td  align="right" background="/images/bg_search.gif" ></td>
                  <td width="13" ><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="35" > <img src='/images/m_reload.gif' width='16' height='16' border=0 alt="����" align='absmiddle' >  ����  <img src='/images/m_edit.gif' width='16' height='15' alt="����"  align='absmiddle'>  ����  <img src='/images/m_delete.gif' width='16' height='15' alt="����"  align='absmiddle'>  ���� </td>
			<td align='right'> <a href="#" onclick="getcontactview(0, 'c'); return false;"><img src='/images/m_add.gif' width='14' height='15'  align='bottom'> �ű� </a>  <a href="#" onclick="getexcel(); return false;"><img src='/images/icon_xls.gif' width='17' height='16'  align='bottom'> ���� </a>  </td>
          </tr>
          <tr>
            <td valign="top" height="516" colspan='2'>

				  <table border="0"width="1030" cellpadding="0" cellspacing="0" bordercolor="#8D652B" id="contact">
				  <thead>
					<tr height='30' align='center'>
						<th class="hd left" width="20">No</th>
						<th class="hd center" width="50">����</th>
						<th class="hd center" width="242">��ü��</th>
						<th class="hd center" width="70">���ʰ��</th>
						<th class="hd center" width="70">��������</th>
						<th class="hd center" width="70">��������</th>
						<th class="hd center" width="70">�ѱ����</th>
						<th class="hd center" width="70">�������</th>
						<th class="hd center" width="70">�����޾�</th>
						<th class="hd center" width="67">������</th>
						<th class="hd center" width="47">������</th>
						<th class="hd center" width="75">������</th>
						<th class="hd right" width="100">���</th>
					</tr>
					</thead>
					<tbody id='tbody'>
					<%
						Do Until rs.eof
							income = monthly-expense
							If monthly = 0 Then incomerate = "0.00" Else 	incomerate = income/monthly*100
					%>
					<tr height='32'>
						<td  class="hd none" style='padding-left:3px; text-align:left;'><span name="totalrecord"><%=totalrecord%></span></td>
						<td  class="hd none" style=';text-align:center;'><a href="#" onclick="getcontactview(<%=contidx%>, 'e'); return false;" ><img src='/images/m_reload.gif' width='16' height='16' border=0 alt="���� ���" hspace=1></a><a href="#" onclick="getcontactview(<%=contidx%>, 'u'); return false;"><img src='/images/m_edit.gif' width='16' height='15' alt="��� ���� ����"  ></a><a href="#" onclick="getcontactdelete(<%=contidx%>); return false;"><img src='/images/m_delete.gif' width='16' height='15' alt="��� ���� ����" hspace=1></a></td>
						<td  class="hd none" style="padding-left:5px;"><a href="#" onclick="getcontact(<%=contidx%>,'<%=flag%>'); return false;" title="<%=title%>" class='subject'><%=cutTitle(title, 40)%></a></td>
						<td  class="hd none" style=' text-align:center;'><%=firstdate%></td>
						<td  class="hd none" style=' text-align:center;'><%=startdate%></td>
						<td  class="hd none" style=' text-align:center;'><%=enddate%></td>
						<td  class="hd none" style='padding-right:3px; text-align:right;'><%=FormatNumber(totalprice, 0)%></td>
						<td  class="hd none" style='padding-right:3px; text-align:right;'><%=FormatNumber(monthly, 0)%></td>
						<td  class="hd none" style='padding-right:3px; text-align:right;'><%=formatnumber(expense,0)%></td>
						<td  class="hd none" style='padding-right:3px; text-align:right;'><%=formatnumber(income,0)%></td>
						<td  class="hd none" style='padding-right:10px; text-align:right;'><%=formatnumber(incomerate,2)%></td>
						<td  class="hd none" style='padding-left:3px;'><%=custname%></td>
						<td  class="hd none" style='padding-left:3px;'><%=teamname%></td>
					</tr>
				  <%
							totalrecord = totalrecord - 1
							grandmonthly = CDbl(grandmonthly) + CDbl(monthly)
							grandexpense = CDbl(grandexpense) + CDbl(expense)
							grandtotalprice = CDbl(grandtotalprice) + CDbl(totalprice)
							rs.movenext
						Loop

						grandincome = CDbl(grandmonthly) - CDbl(grandexpense)
						if grandincome = 0 Then grandincomerate = "0.00" else	grandincomerate = grandincome/grandmonthly *100
				  %>
				  </tbody>
				  <tfoot>
                  <tr height="30">
                    <td class="hd left"  colspan='7' style="text-align:center;"><strong>���հ�</strong> </td>
<!--                     <td class="hd center" style=' text-align:right; font-weight:bold;font-size:11px;'><%=formatnumber(grandtotalprice/1000,0)%>&nbsp;</td> -->
                    <td class="hd center" style=' text-align:right; font-weight:bold'><%=formatnumber(grandmonthly,0)%>&nbsp;</td>
                    <td class="hd center" style=' text-align:right; font-weight:bold'><%=formatnumber(grandexpense,0)%>&nbsp;</td>
                    <td class="hd center" style=' text-align:right; font-weight:bold'><%=formatnumber(grandincome,0)%>&nbsp;</td>
                    <td class="hd center" style=' text-align:right; font-weight:bold'><%=formatnumber(grandincomerate,2)%>&nbsp;</td>
                    <td class="hd right" colspan='2'>&nbsp;</td>
                  </tr>
				  </tfoot>
              </table>
			  <div id="debugConsole"> &nbsp;</div>
			  </td>
          </tr>
      </table></td>
    </tr>
  </table>
</form>
<iframe src='about:blank' name='processFrame' frameborder=0 width='0' height='0'></iframe>
</body>
</html>
<!--#include virtual="/bottom.asp" -->