<!--#include virtual="/inc/getdbcon_first.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim userid : userid = Lcase(request.form("txtuserid"))
	dim password : password = Lcase(request.form("txtpassword"))
	userid = replace(userid, "--", "")
	userid = replace(userid, "'", "")
	password = replace(password, "--", "")
	password = replace(password, "'", "")
%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<SCRIPT LANGUAGE="JavaScript">
<!--
	function password_change() {
		var url = "/password_change.asp?userid=<%=userid%>&password=<%=password%>";
		var name = "password_check";
		var opt = "width=540, height=204, resizable=yes, left=100; top=100";
		window.open(url, name, opt);
	}

	function password_change2() {
		var url = "/password_change2.asp?userid=<%=userid%>&password=<%=password%>";
		var name = "password_check2";
		var opt = "width=540, height=204, resizable=yes, left=100; top=100";
		window.open(url, name, opt);
	}
//-->
</SCRIPT>
<%

	Dim objrs, sql
	sql = "select c.custcode, c.custname, c.highcustcode, d.custname highcustname,a.password, a.isuse, a.class, a.clipinglevel, a.ispwdchange, a.lastchangedate, a.uuser, a.udate from dbo.wb_account a left outer  join dbo.sc_cust_dtl c on c.custcode = a.custcode  left outer join dbo.sc_cust_hdr d on c.highcustcode = d.highcustcode where userid =  '" & userid &"'  and class in ('A', 'N', 'C', 'D', 'G', 'H', 'O', 'F')"
	Call set_recordset(objrs, sql)

	dim custcode2 , custname, custcode, pwd, isuse,  ispwdchange, lastchangedate, clipinglevel, userClass, highcustname
	if objrs.eof then  '계정이 존재하지 않는 경우
		sql = "select empid, medcode, empname, useflag, ispwdchange, lastchangedate, emppwd, clipinglevel , custname from wb_med_employee a inner join sc_cust_hdr b on a.medcode=b.highcustcode where empid=?"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adcmdtext
		cmd.parameters.append	cmd.createparameter("empid", adChar, adParamInput, 9)
		cmd.parameters("empid").value = Left(userid,9)
		dim rs2 : set rs2 = cmd.execute
'		Dim rs2 : Set rs2 = server.CreateObject("adodb.recordset")
'		rs2.cursorlocation = aduseclient
'		rs2.cursortype = adopenStatic
'		rs2.locktype = adLockOptimistic
'		rs2.open cmd.execute
'		response.write rs2.eof
		If Not rs2.eof Then

			if rs2(6) <> password then  '비밀번호가 다른 경우 해당 아이디의 clipinglevel 증가
				clipinglevel = rs2(7) + 1
				Call set_clipingLevel(userid, clipinglevel)
				response.write "<script type='text/javascript'> alert('비밀번호 입력오류 입니다.\n\n입력오류 "&clipinglevel & "회 입니다.'); parent.location.href = '/';</script>"
				response.end
			end if

			if  clipinglevel = 5 then ' 비밀번호 오류 5회 이상시
				response.write "<script type='text/javascript'> alert('비밀번호 오류횟수 초과입니다. \n\n관리 담당자에게 문의바랍니다.'); parent.location.href = '/'; </script>"
				response.end
			end if

			If rs2("useflag")="0" Then response.write "<script> alert('사용이 중지된 아이디입니다.\n\n사용을 원하실 경우 SKMNC 옥외 담당자에게 문의하십시오'); parent.location.href ='/';</script>"
			userClass = "M"
			userid = rs2(0)
			ispwdchange = rs2(4)
			lastchangedate = rs2(5)
			highcustname = rs2(8)
			clipinglevel = rs2(7)
			custname = null

			If rs2(4) = 0 Then
				response.write "<script type='text/javascript'>password_change2();</script>"
				response.end
			End If

			if datediff("m", rs2(5), date ) > 6 then  '아이딩 생성후 분기별로 비밀번호 강제
				response.write "<script type='text/javascript'>password_change2();</script>"
				response.end
			end if
		Else
			response.write "<script type='text/javascript'> alert('비밀번호 미등록자 입니다.\n\n담당 관리자에게 문의바랍니다.'); parent.location.href = '/';</script>"
			response.End
		End If
			Call set_clipingLevel(userid, 0)
	Else	' 해당 아이디로 로그인 했을때
		custcode = objrs("highcustcode")
		custname = objrs("custname")
		pwd = objrs("password")
		isuse = objrs("isuse")
		ispwdchange = objrs("ispwdchange")
		lastchangedate = objrs("lastchangedate")
		clipinglevel = objrs("clipinglevel")
		userClass = objrs("class")
		highcustname = objrs("Highcustname")

		if isuse = "N" then  '사용하지 않는 계정인 경우
			response.write "<script type='text/javascript'> alert('사용이 중지된 아이디입니다..\n\n담당 관리자에게 문의바랍니다.'); parent.location.href = '/';</script>"
			response.end
		end if

		if  clipinglevel = 5 then ' 비밀번호 오류 5회 이상시
			response.write "<script type='text/javascript'> alert('비밀번호 오류횟수 초과입니다. \n\n관리 담당자에게 문의바랍니다.'); parent.location.href = '/'; </script>"
			response.end
		end if

		if pwd <> password then  '비밀번호가 다른 경우 해당 아이디의 clipinglevel 증가
			clipinglevel =clipinglevel + 1
			Call set_clipingLevel(userid, clipinglevel)
			response.write "<script type='text/javascript'> alert('비밀번호 입력오류 입니다.\n\n입력오류 "& clipinglevel + 1 & "회 입니다.'); parent.location.href = '/';</script>"
			response.end
		end if
			Call set_clipingLevel(userid, 0)

	end if

		session("userid") = userid
		session("class") = userClass
		session("LogTime") = Now
		if isnull(custcode) then
			response.cookies("custname") = "옥외 모니터링"
			session("custname") = "옥외 모니터링"
		else
			if userClass = "C" or userClass = "G" then
				response.cookies("custname") = highcustname
				session("custname") =highcustname
			else
				response.write "custname : " & custname
				response.cookies("custname") = custname
				session("custname") = custname
			end if
		end if
		response.cookies("userid") = userid
		response.cookies("class") = userClass
		response.cookies("LogTime") = Now


		if ispwdchange = 0 then ' 비밀번호 변경 이력이 없는 경우
			response.write "<script type='text/javascript'>password_change();</script>"
			response.end
		end if

		if datediff("m", lastchangedate, date ) > 6 then  '아이딩 생성후 분기별로 비밀번호 강제
			response.write "<script type='text/javascript'>password_change();</script>"
			response.end
		end if

	select case userClass
		case "A"
			response.cookies("custcode").path = "/hq"
			response.cookies("custcode2").path = "/hq"
			response.cookies("custname").path = "/hq"
			response.cookies("class").path ="/hq"
			response.cookies("logtime").path ="/hq"
			response.write "<script type='text/javascript'>location.href='/hq/';</script>"
		case "N"
			response.cookies("custcode").path = "/hq"
			response.cookies("custcode2").path = "/hq"
			response.cookies("custname").path = "/hq"
			response.cookies("class").path ="/hq"
			response.cookies("logtime").path ="/hq"
			response.write "<script type='text/javascript'>location.href='/hq/';</script>"
		case "C"
			response.cookies("custcode").path = "/cust"
			response.cookies("custcode2").path = "/cust"
			response.cookies("custname").path = "/cust"
			response.cookies("class").path ="/cust"
			response.cookies("logtime").path ="/cust"
			response.write "<script type='text/javascript'>location.href='/cust/';</script>"
		case "G"
			response.cookies("custcode").path = "/cust"
			response.cookies("custcode2").path = "/cust"
			response.cookies("custname").path = "/cust"
			response.cookies("class").path ="/cust"
			response.cookies("logtime").path ="/cust"
			response.write "<script type='text/javascript'>location.href='/cust/';</script>"
		case "D"
			response.cookies("custcode").path = "/dept"
			response.cookies("custcode2").path = "/dept"
			response.cookies("custname").path = "/dept"
			response.cookies("class").path ="/dept"
			response.cookies("LogTlogtimeime").path ="/dept"
			response.write "<script type='text/javascript'>location.href='/dept/';</script>"
		case "H"
			response.cookies("custcode").path = "/dept"
			response.cookies("custcode2").path = "/dept"
			response.cookies("custname").path = "/dept"
			response.cookies("class").path ="/dept"
			response.cookies("LogTlogtimeime").path ="/dept"
			response.write "<script type='text/javascript'>location.href='/dept/';</script>"
		case "F"	'옥외 모니터링
			response.cookies("custcode").path = "/ODF"
			response.cookies("custcode2").path = "/ODF"
			response.cookies("custname").path = "/ODF"
			response.cookies("class").path ="/ODF"
			response.cookies("logtime").path ="/ODF"
			response.write "<script type='text/javascript'>location.href='/ODF/';</script>"
		case "M"
			response.cookies("custcode").path = "/med"
			response.cookies("custcode2").path = "/med"
			response.cookies("custname").path = "/med"
			response.cookies("class").path ="/med"
			response.cookies("logtime").path ="/med"
			response.write "<script type='text/javascript'>location.href='/med/';</script>"
		case "O"
			response.cookies("custcode").path = "/od"
			response.cookies("custcode2").path = "/od"
			response.cookies("custname").path = "/od"
			response.cookies("class").path ="/od"
			response.cookies("logtime").path ="/od"
			response.write "<script type='text/javascript'>location.href='/od/outdoor/';</script>"
	end select

	objrs.close
	set objrs = nothing
%>