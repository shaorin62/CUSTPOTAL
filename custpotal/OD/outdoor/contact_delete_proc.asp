<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	'����ȣ�� �ش��ϴ� ��� ��ü �� �ݾ� ���� ������ �����Ŀ� ó�� �Ѵ�.

	dim contidx						'������ ����ȣ
	dim objrs, objrs2
	dim sql
	dim arySidx
	dim cyear
	dim cmonth
	dim custcode
	dim custcode2
	dim searchstring

	call dbcon

	cyear = request("cyear")
	custcode = request("custcode")
	custcode2 = request("custcode2")
	cmonth = request("cmonth")
	contidx = request("contidx")
	searchstring = request("searchstring")

	sql = "select contidx, m.custcode as custcode2 , c.highcustcode as custcode from dbo.wb_contact_mst m inner join dbo.sc_cust_temp c on m.custcode = c.custcode where contidx = " & contidx
	call get_recordset(objrs, sql)

	if objrs.eof then
		response.write "<script> alert('��������� �̹� ����, �Ǵ� ����� ��������Դϴ�.'); history.back(); </script>"
	else

	end if
	objrs.close


	sql = "select * from dbo.wb_contact_mst m inner join dbo.wb_contact_md m2 on m.contidx = m2.contidx inner join dbo.wb_contact_md_dtl d on m2.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on d.idx = a.idx where m.contidx = "&contidx&" and a.isPerform = 1"
	call set_recordset(objrs, sql)

	'��ϵ� ��ü ������ ���ٸ� ��� ������ �ٷ� �����Ѵ�.
	if objrs.eof then
		objconn.execute "delete from dbo.wb_contact_md_dtl_account where idx in (select idx from dbo.wb_contact_md_dtl where sidx in (select sidx from dbo.wb_contact_md where contidx = " & contidx &"))"
		objconn.execute "delete from dbo.wb_contact_md_dtl where sidx in (select sidx from dbo.wb_contact_md where contidx = " & contidx &")"
		objconn.execute "delete from dbo.wb_contact_md where contidx = " & contidx
		objconn.execute "delete from dbo.wb_contact_mst where contidx = " & contidx
		response.write "<script> this.close(); </script>"
	else
		response.write "<script> alert('����� ������ �����ϴ� ����� ������ �Ұ����մϴ�.\n\n����� ������ ����Ͻ� �� �����ϼ���'); history.back(); </script>"

	end if

%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	window.opener.location.href = "contact_list.asp?cyear=<%=cyear%>&cmonth=<%=cmonth%>&selcustcode=<%=custcode%>&selcustcode2=<%=custcode2%>&txtsearchstring=<%=searchstring%>";
	this.close();
//-->
</SCRIPT>