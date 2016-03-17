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

	if isuse = "N" then  '사용하지 않는 계정인 경우
		response.write "<script type='text/javascript'> alert('사용이 중지된 아이디입니다..\n\n담당 관리자에게 문의바랍니다.'); parent.location.href = '/index.htm'</script>"
		resposne.end
	end if

	if objrs.eof then  '계정이 존재하지 않는 경우
		response.write "<script type='text/javascript'> alert('비밀번호 미등록자 입니다.\n\n담당 관리자에게 문의바랍니다.'); parent.location.href = '/index.htm'</script>"
		resposne.end
	end if

	if clipinglevel = 5 then ' 비밀번호 오류 5회 이상시
		response.wrie "<script type='text/javascript'> alert('비밀번호 오류횟수 초과입니다. \n\n관리 담당자에게 문의바랍니다.'); parent.location.href = '/index.htm'; </script>"
		resposne.end
	end if

	if pwd <> password then  '비밀번호가 다른 경우 해당 아이디의 clipinglevel 증가
		objrs.fields("clipinglevel").value = clipinglevel + 1
		objrs.update
		response.write "<script type='text/javascript'> alert('비밀번호 입력오류 입니다.\n\n입력오류 "& clipinglevel + 1 & "회 입니다.'); parent.location.href = '/index.htm'</script>"
		resposne.end
	end

	if not ispwdchange then ' 비밀번호 변경 이력이 없는 경우
		response.write "<script type='text/javascript'>password_change();</script>"
		resposne.end
	end if

	if datediff("m", lastchangedate, date ) > 6 then  '아이딩 생성후 분기별로 비밀번호 강제
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