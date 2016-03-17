<%
'���� ���輺�� �����ϴ� ���ڵ��� ���͸�
'���ڿ� �Է°��� ����
'�������� ������ Ÿ���� ������ üũ�ϵ��� �Ѵ�.
Function sqlFilter(search)
	Dim strSearch(5), strReplace(5), cnt, data

	'SQL Injection Ư������ ���͸�
	'�ʼ� ���͸� ���� ����Ʈ
	strSearch(0)="'"
	strSearch(1)=""""
	strSearch(2)="\"
	strSearch(3)=null
	strSearch(4)="#"
	strSearch(5)="--"
	strSearch(6)=";"

	'��ȯ�� ���� ����
	strReplace(0)="''"
	strReplace(1)=""""""
	strReplace(2)="\\"
	strReplace(3)="\"&null
	strReplace(4)="\#"
	strReplace(5)="\--"
	strReplace(6)="\;"

	data = search
	For cnt = 0 to 6 '���͸� �ε����� �迭 ũ��� �����ش�.
		data = replace(data, LCASE(strSearch(cnt)), strReplace(cnt))
	Next

	sqlFilter = data
End Function

'XSS ��� ���� �Լ�
'XSS ���� �Լ�
'$str - ���͸��� ��°�
'$avatag - ����� �±� ����Ʈ ��)  $avatag = "p,br"
Function clearXSS(strString, avatag)
	'XSS ���͸�
	strString = replace(strString, "<", "&lt;")
	strString = replace(strString, ">", "&gt;")
	strString = replace(strString, "(", "&#40;")
	strString = replace(strString, ")", "&#41;")
	'strString = replace(strString, "#", "&#35;")
	'strString = replace(strString, "&", "&#38;")
	strString = replace(strString, "\0", "")

	'����� �±� ��ȯ
	avatag = replace(avatag, " ", "")		'���� ����
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

'Ȯ���� �˻�
'$filename: ���ϸ�
'$avaext: ����� Ȯ���� ��) $avaext = "jpg,gif,pdf"
'���ϰ�: true-"ok", false-"error"
Function Check_Ext(filename,avaext)
	Dim bad_file, FileStartName, FileEndName
	Dim p
	Dim ok_file

	Check_Ext = "error"
'	If instr(filename, "\0") Then
'		Response.Write "������� �ʴ� �Է°�"
'		Response.End
'	End If

	'���ε� ���� Ȯ���� üũ
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

	'����� Ȯ���� üũ
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
'		Response.Write "������� �ʴ� �Է°�"
'		Response.End
'	End If

	'�Է� ���� ���� üũ
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

'�ٿ�ε� ��� üũ �Լ�
'$dn_dir - �ٿ�ε� ���丮 ���(path)
'$fname - �ٿ�ε� ���ϸ�
'���� - true:�Ŀ�ε� ���� ���, false: "error"
Function Check_Path(dn_dir, fname)
	'���丮 �����ڸ� �ϳ��� ����
	dn_dir = Replace(dn_dir, "/", "\")
	fname = Replace(fname, "/", "\")

	strFile = Server.MapPath(dn_dir) & "\" & fname '���� ������

	strFname = Mid(fname,InstrRev(fname,"\")+1) '���� �̸� ����, ..\ ���� ���� ��� Ž���� ���� ��
	Response.Write strFname

	strFPath = Server.MapPath(dn_dir) & "\" & strFname '�������� ���� �ٿ�ε� ���� ���

	If strFPath = strFile Then
		Check_Path = strFile '������ ��� ���� ��� ����
	Else
		Check_Path = "error"
	End If
End Function

'IP üũ �Լ�
Function Check_IP(IP_Addr)
	If Request.Servervariables("REMOTE_ADDR") = IP_Addr Then
		Check_IP = "TRUE"
	Else
		Check_IP = "FALSE"
	End If
End Function
%>
