<!--#include virtual="/cust/admin/inc/func.asp" -->
<%
	dim userid : userid = request.QueryString("userid")

	Dim objrs
	dim cmd
	dim sql : sql = "select a.userid, a.password, a.class, a.custcode, c.custname, isuse from dbo.wb_account a left outer join dbo.sc_cust_dtl c on a.custcode = c.custcode where a.userid=?"
	set cmd = getCommand(cmd, sql)
	cmd.parameters.append cmd.createparameter("userid", adVarchar, adParamInput, 120
	cmd.parameters("userid").value = userid


	Call get_recordset(objrs, sql)

	Dim password, class_, custcode, custname, isuse
	if not objrs.eof Then
		userid = objrs("userid")
		password = objrs("password")
		class_ = objrs("class")
		custcode = objrs("custcode")
		custname = objrs("custname")
		isuse = objrs("isuse")
	else
		response.write "<script type='text/javascript'> alert('������ �����̰ų� �߸��� �������̵� �Դϴ�.'); location.href='/main.asp';</script>"
	end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<title>�Ƣ� SK M&C | Media Management System �Ƣ� </title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>

<body background="/images/pop_bg.gif" >
<form>

<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> <%=userid%> ���� ����</td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="540" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="22"><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td>
<!--  --><table  border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
              <tr>
                <td class="hw">���̵�</td>
                <td class="bw bbd"> <%=userid%></td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
              <tr>
                <td class="hw">��й�ȣ</td>
                <td class="bw bbd"> ************** </td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
              <tr>
                <td class="hw">���ӱ���</td>
                <td class="bw bbd">
				<%
					select case class_
					case  "A" : response.write "Administrator"
					case  "N" : response.write "Admin(Non-SKT)"
					case  "C" : response.write "������"
					case  "D" : response.write "����μ�"
					case  "O" : response.write "���� ������"
					case  "F" : response.write "���� ����͸�"
					case  "M" : response.write "��ü��"
					end select
				%></td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
              <tr>
                <td class="hw">�����Ҽ�</td>
                <td class="bw bbd">  <%if isnull(custname) then response.write "���� ����͸� ��ü" else response.write custname%></td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
              <tr>
                <td class="hw">��뿩��</td>
                <td class="bw bbd"> <%if ucase(isuse) = "Y" then response.write "�����" else response.write "�������"%></td>
              </tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
                <tr>
                  <td  height="50" valign="bottom"><img src="/images/space.gif" width="57" height="20" border="0"></td>
                  <td  align="right" valign="bottom"> <img src="/images/btn_init.gif" width="67" height="18" vspace="5" style="cursor:hand" onClick="set_reset_password();" hspace="5"><img src="/images/btn_edit.gif" width="57" height="18" vspace="5" border="0" class="stylelink" onClick="pop_account_edit();"><% if ucase(isuse) = "Y" then %><img src="/images/btn_stop.gif" width="78" height="18" vspace="5" style="cursor:hand" onClick="set_account_stop();" hspace="5"><% else %><img src="/images/btn_use.gif" width="78" height="18" vspace="5" style="cursor:hand" onClick="set_account_use();" hspace="5"><% end if %><img src="/images/btn_close.gif" width="57" height="18" vspace="5" style="cursor:hand" onClick="set_close();" ></td>
                </tr>
              </table>
<!--  -->
	</td>
    <td background="/images/pop_right_middle_bg.gif">&nbsp;</td>
  </tr>
  <tr>
    <td><img src="/images/pop_left_bottom_bg.gif" width="22" height="25"></td>
    <td background="/images/pop_center_bottom_bg.gif">&nbsp;</td>
    <td><img src="/images/pop_right_bottom_bg.gif" width="23" height="25"></td>
  </tr>
</table>
</form>
<iframe name="scriptFrame" id="scriptFrame" width="0" height="0" frameborder="0" src="/cust/space.asp"></iframe>
</body>
</html>
<SCRIPT LANGUAGE="JavaScript">
<!--
	function pop_account_edit() {
		var url = "pop_account_edit.asp?userid=<%=userid%>";
		var name = "pop_account_edit";
		var opt = "width=540, height=318, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		//window.open (url, name, opt) ;
		location.href=url ;
	}

	function set_reset_password() {
		if (confirm("������ ��й�ȣ�� ���� ��й�ȣ�� �ʱ�ȭ�մϴ�.\n\n��й�ȣ�� �ʱ�ȭ �Ͻðڽ��ϱ�?")) {
			scriptFrame.location.href = "account_password_init.asp?userid=<%=userid%>";
		}
	}

	function set_account_stop() {
		if (confirm("������ ������� ��Ű�ڽ��ϱ�?")) {
			scriptFrame.location.href = "account_stop.asp?userid=<%=userid%>";
		}
	}

	function set_account_use() {
		if (confirm("���������� �����Ͻðڽ��ϱ�?")) {
			scriptFrame.location.href = "account_use.asp?userid=<%=userid%>";
		}
	}

	function set_close() {
		this.close() ;
	}
//-->
</SCRIPT>