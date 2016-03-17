<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	'계약번호에 해당하는 모든 매체 및 금액 정보 삭제에 동의후에 처리 한다.

	dim contidx						'삭제할 계약번호
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
		response.write "<script> alert('계약정보가 이미 삭제, 또는 변경된 계약정보입니다.'); history.back(); </script>"
	else

	end if
	objrs.close


	sql = "select * from dbo.wb_contact_mst m inner join dbo.wb_contact_md m2 on m.contidx = m2.contidx inner join dbo.wb_contact_md_dtl d on m2.sidx = d.sidx inner join dbo.wb_contact_md_dtl_account a on d.idx = a.idx where m.contidx = "&contidx&" and a.isPerform = 1"
	call set_recordset(objrs, sql)

	'등록된 매체 정보가 없다면 계약 정보를 바로 삭제한다.
	if objrs.eof then
		objconn.execute "delete from dbo.wb_contact_md_dtl_account where idx in (select idx from dbo.wb_contact_md_dtl where sidx in (select sidx from dbo.wb_contact_md where contidx = " & contidx &"))"
		objconn.execute "delete from dbo.wb_contact_md_dtl where sidx in (select sidx from dbo.wb_contact_md where contidx = " & contidx &")"
		objconn.execute "delete from dbo.wb_contact_md where contidx = " & contidx
		objconn.execute "delete from dbo.wb_contact_mst where contidx = " & contidx
		response.write "<script> this.close(); </script>"
	else
		response.write "<script> alert('정산된 내역이 존재하는 계약은 삭제가 불가능합니다.\n\n정산된 내역을 취소하신 후 삭제하세요'); history.back(); </script>"

	end if

%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	window.opener.location.href = "contact_list.asp?cyear=<%=cyear%>&cmonth=<%=cmonth%>&selcustcode=<%=custcode%>&selcustcode2=<%=custcode2%>&txtsearchstring=<%=searchstring%>";
	this.close();
//-->
</SCRIPT>