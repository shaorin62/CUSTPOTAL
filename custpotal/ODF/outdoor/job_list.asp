<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim gotopage : gotopage = request.querystring("gotopage")
	if gotopage = "" then gotopage = 1
	dim searchstring : searchstring = request("txtsearchstring")
	dim custcode : custcode = request("selcustcode")
	dim seqno : seqno = request("selseqno")
	if seqno = "" then seqno = null

	dim sql : sql = "select count(jobidx) from dbo.wb_jobcust; " &_
						"select s.jobidx, s.seqno, s.thema, j.seqname, c.custname, c2.custname as custname2 from dbo.wb_jobcust s " &_
						"inner join dbo.sc_jobcust j on s.seqno = j.seqno " &_
						"inner join dbo.sc_cust_temp c on j.custcode = c.custcode " & _
						"inner join dbo.sc_cust_temp c2 on j.clientsubcode = c2.custcode "&_
						"where c.custcode like '" & custcode & "%' and s.thema like '%" & searchstring & "%' "
						if not isnull(seqno) then sql = sql & " and s.seqno = " & seqno
						sql = sql & "order by c.custcode, c2.custcode, seqno, thema"
	dim objrs
	call get_recordset(objrs, sql)

	dim cnt : cnt = objrs(0)
	set objrs = objrs.nextrecordset
	dim jobidx, seqname, thema, startdate, enddate, custname, custname2
	if not objrs.eof then
		set jobidx = objrs("jobidx")				' �����ȣ
		set seqname = objrs("seqname")			' �귣���ڵ�
		set thema = objrs("thema")			' �귣���
		set custname = objrs("custname")		' �����ָ�
		set custname2 = objrs("custname2")	' ����θ�
	end if
	dim mode : mode = null
	dim url : url = "job_list.asp"
%>
<html>
<head>
<title>�Ƣ� SK M&C | Media Management System �Ƣ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  oncontextmenu="return false">
<form name="form1" method="post" action="">
<!--#include virtual="/od/top.asp" -->
  <table id="Table_01" width="1240"  border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/od/left_outdoor_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_outdoor.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td align="left" valign="top" height="600" >
	  <table width="1002" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="19">&nbsp;</td>
          </tr>
          <tr>
            <td height="17"><TABLE  width="100%">
            <TR>
				<TD width="50%"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle"> ������Ȳ </span></TD>
				<TD width="50%" align="right"><span class="navigator" > ���ܰ��� &gt; ��ü���� &gt; ������Ȳ</span></TD>
            </TR>
            </TABLE>
			</td>
          </tr>
          <tr>
            <td height="15">&nbsp;</td>
          </tr>
          <tr>
            <td>
			<table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13"><img src="/images/bg_search_left.gif" width="13" height="35"></td>
                  <td align="left" background="/images/bg_search.gif"><%call get_custcode_mst(custcode, mode, url)%><% call get_jobcust(custcode, seqno, null, null)%><input type="text" name="txtsearchstring"> <img src="/images/btn_search.gif" width="39" height="20" align="absmiddle" onClick="get_search();" class="styleLink" ></td>
                  <td align="right" background="/images/bg_search.gif"><img src="/images/btn_job_reg.gif" width="86" height="18" align="absmiddle" border="0" class="stylelink" onclick="pop_job_reg();"> </td>
                  <td width="13"><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td height="26" >&nbsp;</td>
          </tr>
          <tr>
            <td ><table width="1030" height="31" border="3" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
                <tr>
                  <td><table width="1024" border="0" cellspacing="0" cellpadding="0" class="header">
                      <tr>
                        <td width="40" align="center" class="header">No</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="214" align="center" >������ </td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="170" align="center" >�����</td>
                        <td width="3" align="center" ><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="250" align="center" >�귣��</td>
                        <td width="3" align="center"><img src="/images/ico_head_clip.gif" width="3" height="25"></td>
                        <td width="350" align="center">�����</td>
                      </tr>
                  </table></td>
                </tr>
              </table>
                <table width="1024" height="31" border="0" cellpadding="0" cellspacing="0" bordercolor="#8D652B">
				<% do until objrs.eof %>
                  <tr onclick="go_medium_view('<%=jobidx%>')" class="styleLink">
                    <td width="40" height="31" align="center"><%=cnt%></td>
                    <td width="3">&nbsp;</td>
                    <td width="214" align="center"> <%=custname%> </td>
                    <td width="3">&nbsp;</td>
                    <td width="170" align="center"> <%=custname2%> </td>
                    <td width="3">&nbsp;</td>
                    <td width="250" align="left"> &nbsp;&nbsp;&nbsp;<%=seqname%> </td>
                    <td width="3">&nbsp;</td>
                    <td width="350" align="left"> &nbsp;&nbsp;&nbsp;<%=thema%> </td>
                  </tr>
                  <tr>
                    <td height="1" bgcolor="#E7E9E3" colspan="11"></td>
                  </tr>
				<%
						cnt = cnt - 1
						objrs.movenext
					loop
					objrs.close
					set objrs = nothing
				%>
              </table></td>
          </tr>
      </table></td>
    </tr>
  </table>
<!--#include virtual="bottom.asp" -->
</form>
</body>
</html>
<script language="JavaScript">
<!--
	function go_medium_view(idx) {
		var url = "pop_job_view.asp?jobidx=" + idx + "&gotopage=<%=gotopage%>&searchstring=<%=searchstring%>";
		var name = "pop_job_view";
		var opt = "width=540, height=236, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function pop_job_reg() {
		var url = "pop_job_reg.asp";
		var name = "pop_job_reg";
		var opt = "width=540, height=236, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open(url, name, opt);

	}

	function go_page(url) {
		var frm = document.forms[0];
		location.href="job_list.asp?selcustcode="+frm.selcustcode.options[frm.selcustcode.selectedIndex].value;
	}

	function get_search() {
		var frm = document.forms[0];
		frm.action = "job_list.asp";
		frm.method = "post";
		frm.submit();
	}
//-->
</script>