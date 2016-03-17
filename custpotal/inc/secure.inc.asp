<%
'공격 위험성이 존재하는 문자들을 필터링
'문자열 입력값을 검증
'숫자형은 데이터 타입을 별도로 체크하도록 한다.
Function sqlFilter(search)
	Dim strSearch(5), strReplace(5), cnt, data

	'SQL Injection 특수문자 필터링
	'필수 필터링 문자 리스트
	strSearch(0)="'"
	strSearch(1)=""""
	strSearch(2)="\"
	strSearch(3)=null
	strSearch(4)="#"
	strSearch(5)="--"
	strSearch(6)=";"

	'변환될 필터 문자
	strReplace(0)="''"
	strReplace(1)=""""""
	strReplace(2)="\\"
	strReplace(3)="\"&null
	strReplace(4)="\#"
	strReplace(5)="\--"
	strReplace(6)="\;"

	data = search
	For cnt = 0 to 6 '필터링 인덱스를 배열 크기와 맞춰준다.
		data = replace(data, LCASE(strSearch(cnt)), strReplace(cnt))
	Next

	sqlFilter = data
End Function

'XSS 출력 필터 함수
'XSS 필터 함수
'$str - 필터링할 출력값
'$avatag - 허용할 태그 리스트 예)  $avatag = "p,br"
Function clearXSS(strString, avatag)
	'XSS 필터링
	strString = replace(strString, "<", "&lt;")
	strString = replace(strString, ">", "&gt;")
	strString = replace(strString, "(", "&#40;")
	strString = replace(strString, ")", "&#41;")
	'strString = replace(strString, "#", "&#35;")
	'strString = replace(strString, "&", "&#38;")
	strString = replace(strString, "\0", "")

	'허용할 태그 변환
	avatag = replace(avatag, " ", "")		'공백 제거
	If (avatag <> "") Then
		taglist = split(avatag, ",")

		for each p in taglist
			strString = replace(strString, "&lt;"&p&" ", "<"&p&" ", 1, -1, 1)
			strString = replace(strString, "&lt;"&p&">", "<"&p&">", 1, -1, 1)
			strString = replace(strString, "&lt;/"&p&" ", "</"&p&" ", 1, -1, 1)
		next
	End If

	clearXSS = strString
End Function

'확장자 검사
'$filename: 파일명
'$avaext: 허용할 확장자 예) $avaext = "jpg,gif,pdf"
'리턴값: true-"ok", false-"error"
Function Check_Ext(filename,avaext)
	Dim bad_file, FileStartName, FileEndName
	Dim p
	Dim ok_file

	Check_Ext = "error"
'	If instr(filename, "\0") Then
'		Response.Write "허용하지 않는 입력값"
'		Response.End
'	End If

	'업로드 금지 확장자 체크
	bad_file = "ASP,HTML,HTM,ASA,HTA,JS,ASP,PHP,EXE,JSP,CGI,PERL,PL"


	filename = Replace(filename, " ", "")
	filename = Replace(filename, "%", "")

	FileStartName = Left(filename,InstrRev(filename,".")-1)
	FileEndName = Mid(filename, InstrRev(filename, ".")+1)

	bad_file = split(bad_file, ",")

	for each p in bad_file
		if instr(UCase(FileEndName) , p)>0 then
			Check_Ext = "error"
			Exit Function
		end If
	Next

	'허용할 확장자 체크
	if avaext <> "" Then
		ok_file = split(avaext, ",")

		for each p in ok_file
			if instr(UCase(FileEndName), p)>0 then
				Check_Ext = "ok"
				Exit Function
			End If
		next
	End If

	Check_Ext = "error"
End Function

Function Check_SpecialKey(strKEY)
	Dim bad_key
	Dim FileEndName
	Dim p
	Dim ok_file

	Check_SpecialKey = "ok"
'	If instr(filename, "\0") Then
'		Response.Write "허용하지 않는 입력값"
'		Response.End
'	End If

	'입력 금지 문자 체크
	bad_key = ";, ,:,--"

	FileEndName = Mid(strKEY, InstrRev(strKEY, ".")+1)

	bad_key = split(bad_key, ",")

	for each p in bad_key
		if instr(UCase(FileEndName) , p)>0 then
			Check_SpecialKey = "error"
			Exit Function
		end If
	Next

	Check_SpecialKey = "ok"
End Function

'다운로드 경로 체크 함수
'$dn_dir - 다운로드 디렉토리 경로(path)
'$fname - 다운로드 파일명
'리턴 - true:파운로드 파일 경로, false: "error"
Function Check_Path(dn_dir, fname)
	'디렉토리 구분자를 하나로 통일
	dn_dir = Replace(dn_dir, "/", "\")
	fname = Replace(fname, "/", "\")

	strFile = Server.MapPath(dn_dir) & "\" & fname '서버 절대경로

	strFname = Mid(fname,InstrRev(fname,"\")+1) '파일 이름 추출, ..\ 등의 하위 경로 탐색은 제거 됨
	Response.Write strFname

	strFPath = Server.MapPath(dn_dir) & "\" & strFname '웹서버의 파일 다운로드 절대 경로

	If strFPath = strFile Then
		Check_Path = strFile '정상일 경우 파일 경로 리턴
	Else
		Check_Path = "error"
	End If
End Function

'IP 체크 함수
Function Check_IP(IP_Addr)
	If Request.Servervariables("REMOTE_ADDR") = IP_Addr Then
		Check_IP = "TRUE"
	Else
		Check_IP = "FALSE"
	End If
End Function
%>
