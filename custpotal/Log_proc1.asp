<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim userid : userid = Lcase(request.form("txtuserid"))
	dim password : password = Lcase(request.form("txtpassword"))
	userid = replace(userid, "--", "")
	password = replace(password, "--", "")


	Dim objrs, sql
	sql = "select c.custcode, c.custname, c.highcustcode, a.password, a.isuse, a.ispwdchange, a.lastchangedate from dbo.wb_account a inner join dbo.sc_cust_temp c on c.custcode = a.custcode where userid = '" & userid &"'"
	Call set_recordset(objrs, sql)

	dim custcode2 , custname2, custcode, pwd, isuse,  ispwdchange, lastchangedate
	custcode = objrs("highcustcode")
	custcode2 = objrs("custcode")
	custname2 = objrs("custname")
	pwd = objrs("password")
	isuse = objrs("isuse")
	ispwdchange = objrs("ispwdchange")
	lastchangedate = objrs("lastchangedate")

	if isuse = "N" then  '������� �ʴ� ������ ���
		response.write "<script type='text/javascript'> alert('����� ������ ���̵��Դϴ�..\n\n��� �����ڿ��� ���ǹٶ��ϴ�.'); parent.location.href = '/index.htm'</script>"
		resposne.end
	end if

	if objrs.eof then  '������ �������� �ʴ� ���
		response.write "<script type='text/javascript'> alert('��й�ȣ �̵���� �Դϴ�.\n\n��� �����ڿ��� ���ǹٶ��ϴ�.'); parent.location.href = '/index.htm'</script>"
		resposne.end
	end if

	if clipinglevel = 5 then ' ��й�ȣ ���� 5ȸ �̻��
		response.wrie "<script type='text/javascript'> alert('��й�ȣ ����Ƚ�� �ʰ��Դϴ�. \n\n���� ����ڿ��� ���ǹٶ��ϴ�.'); parent.location.href = '/index.htm'; </script>"
		resposne.end
	end if

	if pwd <> password then  '��й�ȣ�� �ٸ� ��� �ش� ���̵��� clipinglevel ����
		objrs.fields("clipinglevel").value = clipinglevel + 1
		objrs.update
		response.write "<script type='text/javascript'> alert('��й�ȣ �Է¿��� �Դϴ�.\n\n�Է¿��� "& clipinglevel + 1 & "ȸ �Դϴ�.'); parent.location.href = '/index.htm'</script>"
		resposne.end
	end

	if not ispwdchange then ' ��й�ȣ ���� �̷��� ���� ���
		response.write "<script type='text/javascript'>password_change();</script>"
		resposne.end
	end if

	if datediff("m", lastchangedate, date ) > 6 then  '���̵� ������ �б⺰�� ��й�ȣ ����
		response.write "<script type='text/javascript'>password_change();</script>"
		resposne.end
	end if

	response.cookies("userid") = userid
	response.cookies("custcode") = objrs("custcode")
	response.cookies("custcode2") =
	response.cookies("custname") =
	select case objrs("class")
		case "A"
			response.write "<script type='text/javascript'>parent.location.href='/hq/main.asp';</script>"
		case "C"
			response.write "<script type='text/javascript'>parent.location.href='/cust/main.asp';</script>"
		case "D"
			response.write "<script type='text/javascript'>parent.location.href='/cust/main.asp';</script>"
		case "R"
			response.write "<script type='text/javascript'>parent.location.href='/cust/main.asp';</script>"
		case "M"
			response.write "<script type='text/javascript'>parent.location.href='/cust/main.asp';</script>"
		case "O"
			response.write "<script type='text/javascript'>parent.location.href='/cust/main.asp';</script>"
	end select

	objrs.close
	set objrs = nothing
%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	function password_change() {
		var url = "/password_change.asp?userid=<%=userid%>&password=<%=password%>";
		var name = "password_check";
		var opt = "width=540, height=204, resizable=yes, left=100; top=100";
		window.open(url, name, opt);
	}
//-->
</SCRIPT>