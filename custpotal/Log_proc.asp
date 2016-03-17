<!--#include virtual="/inc/func.asp" -->
<%
	dim userid : userid = request.form("txtuserid")
	dim password : password = request.form("txtpassword")
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

	function clearCookie(){
		document.cookie = "cookiemidx= "
		document.cookie = "midx= "
		document.cookie = "cookieittlename= "
		document.cookie = "title= "
	}


//-->
</SCRIPT>
<%

	Dim rs, sql, usertype
	'sql = "select c.custcode, c.custname, c.highcustcode, d.custname highcustname,a.password, a.isuse, a.class, a.clipinglevel, a.ispwdchange, a.lastchangedate, a.uuser, a.udate from dbo.wb_account a left outer  join dbo.sc_cust_dtl c on c.custcode = a.custcode  left outer join dbo.sc_cust_hdr d on c.highcustcode = d.highcustcode where userid = ?"

	sql = "select userid, username, password, isuse,  class,  clipinglevel,  ispwdchange,  lastchangedate,  uuser,  udate, cnt  from dbo.wb_account where userid = ? "
	dim cmd : set cmd = server.createobject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandtype = adCmdText
	cmd.commandText = sql
	cmd.parameters.append	cmd.createparameter("userid", adVarchar, adParamINput, 20)
	cmd.parameters("userid").value = userid
	set rs = cmd.execute
	usertype = true



	if rs.eof then '' 일반 계정 테이블이 존재하지 않으면 매체사 계정 테이블은 검사한다.
		sql = "select empid, empname,  emppwd, isuse,class, clipinglevel, ispwdchange, lastchangedate, '' uuser, '' udate  from wb_med_employee  where empid=?"
		cmd.commandText = sql
		set rs = cmd.execute
		usertype =False



		if rs.eof then
			' 존재하지 않는 아이디
			response.write "<script type='text/javascript'> alert('비밀번호 미등록자 입니다.\n\n02)6390-3981로 문의바랍니다.'); parent.location.href = '/';</script>"
			response.End
		end if
	end if



	dim clipinglevel

	if rs(3) = "N" then
		response.write "<script> alert('사용이 중지된 아이디입니다.\n\n사용을 원하실 경우 \n\n02)6390-3981로 문의바랍니다.'); parent.location.href ='/';</script>"
		response.end
	end if

	if rs(2) <> password then
		clipinglevel = rs(5) + 1
		if ucase(rs(4)) <> "M" then
			Call set_clipingLevel(userid, clipinglevel) ' MC _  비밀번호 오류시 클리핑 레벨 증가
		else
			Call set_clipingLevel2(userid, clipinglevel) ' 매체사 _  비밀번호 오류시 클리핑 레벨 증가
		end if

		if  clipinglevel >= 5 then ' 비밀번호 오류 5회 이상시
			if ucase(rs(4)) <> "M" then
				Call set_isuse(userid)  'MC _  오류가 5회일시에는 레벨 0 , 사용여부를 N으로 만든다.
			else
				Call set_isuse2(userid) ' 매체사 _ 오류가 5회일시에는 레벨 0 , 사용여부를 N으로 만든다.
			end if
			response.write "<script type='text/javascript'> alert('입력오류 횟수 초과로 사용중지 되었습니다. \n\n02)6390-3944로 문의바랍니다.'); parent.location.href = '/'; </script>"
			response.end
		else
			response.write "<script type='text/javascript'> alert('비밀번호 입력오류 입니다.\n\n입력오류 "&clipinglevel & "회 입니다.'); parent.location.href = '/';</script>"
		end if

		response.end
	end if






'  옥외팀 요청사항 - 3개월단위로 비밀번호 변경은 의미가 없음.
'	if not rs(8) or datediff("m", rs(9), date ) > 6 then
'		if ucase(rs(6)) <> "M" then
'			response.write "<script type='text/javascript'>password_change();</script>"
'			response.end
'		else
'			response.write "<script type='text/javascript'>password_change2();</script>"
'			response.end
'		end if
'	end if


	if ucase(rs(4)) <> "M" then
		Call set_initclipingLevel(userid) 'MC 로그인이 잘되었으면 클리핑레벨 0으로 초기화
	else
		Call set_initclipingLevel2(userid) '매체사 로그인이 잘되었으면 클리핑레벨 0으로 초기화
	end if

	if ucase(rs(4)) = "F"  then
		response.cookies("classname") = "옥외 모니터링"
	ElseIf UCase(rs(4)) = "A" Then
		response.cookies("classname") = "admin 관리자"
	ElseIf UCase(rs(4)) = "G" Then
		response.cookies("classname") = "내부 관리자"
	ElseIf UCase(rs(4)) = "C" Then
		response.cookies("classname") = "광고주"
	ElseIf UCase(rs(4)) = "M" Then
		response.cookies("classname") = "매체사 담당자"
	ElseIf UCase(rs(4)) = "Z" Then
		response.cookies("classname") = "내부 공용"
	end If

	if rs(4) = "M" then
		response.cookies("username") = rs(1)
	end if

	response.cookies("userid") = userid
	response.cookies("class") = rs(4)
	response.cookies("LogTime") = Now
	response.cookies("pagename") = ""



	dim objrs


	If UCase(rs(4)) = "M" Then

	else
		sql = "select userid, cnt  from dbo.wb_account where userid ='"&userid&"'"

		call set_recordset(objrs, sql)

		Dim strcnt
		strcnt = CLng(rs(10)) + 1

		objrs.fields("cnt").value =strcnt

		objrs.update
		objrs.close

	end If



	response.write "<script type='text/javascript'>clearCookie();</script>"

	select case rs(4)
		case "A" 'SKMNC 에서 관리자
			response.cookies("class").path ="/hq"
			response.cookies("logtime").path ="/hq"
			response.write "<script type='text/javascript'>location.href='/hq/';</script>"
		case "G" ' SKMNC 일반 사용자
			response.cookies("class").path ="/mp"
			response.cookies("logtime").path ="/mp"
			response.write "<script type='text/javascript'>location.href='/mp/';</script>"
		case "C" ' 일반 광고주
			response.cookies("class").path ="/cust"
			response.cookies("logtime").path ="/cust"
			response.write "<script type='text/javascript'>location.href='/cust/';</script>"
		case "F"	'옥외 모니터링
			response.cookies("custcode").path = "/ODF"
			response.cookies("custcode2").path = "/ODF"
			response.cookies("custname").path = "/ODF"
			response.cookies("class").path ="/ODF"
			response.cookies("logtime").path ="/ODF"
			response.write "<script type='text/javascript'>location.href='/ODF/';</script>"
		case "Z" ' 내부 공용아이디
			Dim stripchk
			stripchk = Mid(Request.ServerVariables("REMOTE_ADDR"),1,9)

			If Request.ServerVariables("REMOTE_ADDR") = "10.110.10.86" Or Request.ServerVariables("REMOTE_ADDR") = "203.235.202.73" Then
				response.cookies("class").path ="/mc"
				response.cookies("logtime").path ="/mc"
				response.write "<script type='text/javascript'>location.href='/mc/';</script>"
			Else
				If   stripchk = "10.110.21" Or stripchk = "10.110.31" Or stripchk = "10.110.32" Or stripchk = "10.110.37" Or stripchk = "10.110.51" Or stripchk = "10.110.52" Or stripchk = "10.110.53" Or stripchk = "10.110.57" Or stripchk = "10.110.61" Or stripchk = "10.110.62" Or stripchk = "10.110.63" Or stripchk = "10.110.67"  Or stripchk = "10.110.86"  Then
					response.cookies("class").path ="/mc"
					response.cookies("logtime").path ="/mc"
					response.write "<script type='text/javascript'>location.href='/mc/';</script>"
				Else
					response.write "<script type='text/javascript'> alert('회사외부에서는 이용하실 수 없습니다.'); parent.location.href = '/'; </script>"
					response.end
				End If
			End If		
		case "M"
			response.cookies("custcode").path = "/med"
			response.cookies("custcode2").path = "/med"
			response.cookies("custname").path = "/med"
			response.cookies("class").path ="/med"
			response.cookies("logtime").path ="/med"
			response.write "<script type='text/javascript'>location.href='/med/';</script>"
	end select

	rs.close
	set rs = nothing
%>

