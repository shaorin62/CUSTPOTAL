<!--#include virtual="/hq/outdoor/inc/function.asp" -->
<%
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")
	if cyear = "" then cyear = year(date)
	if cmonth = "" then cmonth = month(date)
	if Len(cmonth) = 1 then cmonth = "0" & cmonth
	Dim pempid : pempid = request.cookies("userid")

	dim sdate : sdate = dateserial(cyear, cmonth, "01")
	dim edate : edate = dateadd("d", -1, dateadd("m", 1,  sdate))

	If request.cookies("userid") = "" Then response.redirect "/"

	sql = "select distinct a.contidx, a.custcode, a.title, a.startdate, a.enddate from wb_contact_mst a left outer join wb_contact_md b on a.contidx=b.contidx inner join wb_contact_exe c on b.mdidx=c.mdidx and c.cyear=? and c.cmonth=? where a.contidx in (select contidx from wb_contact_md where empid=?) "
'	response.write sql

	Dim cmd : Set cmd = server.CreateObject("adodb.command")
	cmd.activeconnection = application("connectionstring")
	cmd.commandText = sql
	cmd.commandType = adCmdText
	cmd.parameters.append cmd.createparameter("cyear", adchar, adparaminput, 4)
	cmd.parameters.append cmd.createparameter("cmonth", adchar, adparaminput, 2)
	cmd.parameters.append cmd.createparameter("empid", adchar, adparaminput, 9)
	cmd.parameters("cyear").value = cyear
	cmd.parameters("cmonth").value = cmonth
	cmd.parameters("empid").value = pempid
	Dim rs : Set rs = cmd.execute
	clearparameter(cmd)

	If Not rs.eof Then
		Dim contidx : Set contidx = rs(0)
		Dim custcode : Set custcode = rs(1)
		Dim title : Set title = rs(2)
		Dim startdate : Set startdate = rs(3)
		Dim enddate : Set enddate = rs(4)
	End If

	sql = "select sum(isnull(c.expense,0))  expense, a.region, a.locate, sum(isnull(c.qty,0)) qty, a.contidx, filename, a.mdidx, reportname, a.unit from wb_contact_md a  inner join wb_contact_exe c on a.mdidx=c.mdidx left outer join wb_report_dtl d on c.mdidx=d.mdidx and c.cyear = d.cyear and c.cmonth = d.cmonth where c.cyear='"&cyear&"' and c.cmonth='" &cmonth&"' and a.mdidx in (select mdidx from wb_contact_md where empid='"&pempid&"') group by a.region, a.locate, a.contidx, filename, a.mdidx, reportname, unit "

	Dim rs2 : Set rs2 = server.CreateObject("adodb.recordset")
	rs2.activeconnection = application("connectionstring")
	rs2.cursorlocation = aduseclient
	rs2.cursortype = adopenstatic
	rs2.locktype = adLockOptimistic
	rs2.source = sql
	rs2.open

	If Not rs2.eof Then
		Dim expense : Set expense = rs2(0)
		Dim region : Set region = rs2(1)
		Dim locate : Set locate = rs2(2)
		Dim qty : Set qty = rs2(3)
		Dim filename : Set filename = rs2(5)
		Dim mdidx : Set mdidx = rs2(6)
		Dim reportname : Set reportname = rs2(7)
		Dim Unit : Set unit = rs2(8)
	End If

	Set cmd = Nothing
%>

<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/hq/outdoor/style.css" rel="stylesheet" type="text/css">
<script type="text/javascript">
<!--
	var reportPop ;
	document.domain = "mms.raed.co.kr";
	function go_search() {
		var frm = document.forms[0];
		frm.action = "/med/";
		frm.method = "post";
		frm.submit();
	}

	function getupload(mdidx, title, locate, crud) {
		if (crud == 'd') {if (!confirm("선택한 광고의 리포트를 삭제하시겠습니까?")) return false;}
		var url = "/med/popup/view_report.asp?mdidx="+mdidx+"&title="+title+"&locate="+locate+"&cyear=<%=cyear%>&cmonth=<%=cmonth%>&crud="+crud;
		var name = "reportPop";
		var left = screen.width / 2 - 600 / 2;
		var top = screen.height / 2 - 180 / 2;
		var opt = "width=600, height=180, resizable=no, scrollbars=no, status=yes, left="+left+",top="+top;
		reportPop = window.open(url, name, opt);
	}

	function download(file) {
		location.href='/med/download.asp?filename='+file;
	}

	window.onload = function () {
		self.focus();
	}

	window.onunload = function () {
		try {
			reportPop.close();
		} catch (e) {
			return true;
		}
	}
//-->
</script>
</head>
<!-- <body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"   oncontextmenu="return false">
 --><body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"   >
<form>
<table width="1240" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td height="24" background="/images/pop_top.gif" valign="top" >
	<% if ucase(request.cookies("class")) = "M" then %><table width="700"  border="0" align="right" cellpadding="0" cellspacing="0" height="60">
      <tr style="padding-top:10">
        <td>&nbsp;</td>
        <td width="244" align="right" valign="top" ><span class="log">&nbsp;<%=request.cookies("username")%></span> &nbsp;</td>
        <td width="1" valign="top"><img src="/images/top_vline_bg.gif" width="1" height="32"></td>
        <td width="104" align="right" valign="top" ><span class="log">&nbsp;<%=request.cookies("userid")%></span> &nbsp;</td>
        <td width="1" valign="top" ><img src="/images/top_vline_bg.gif" width="1" height="32"></td>
        <td width="164" align="right" valign="top" ><span class="log"><%=request.cookies("logtime")%>&nbsp;</span></td>
        <td width="1" valign="top"><img src="/images/top_vline_bg.gif" width="1" height="32"></td>
        <td width="85" align="center"valign="top" ><A HREF="/Log_out.asp"><img src="/images/btn_logout.gif" width="64" height="19" border="0"></A></td>
      </tr>
    </table>
	<% else %><table width="700"  border="0" align="right" cellpadding="0" cellspacing="0" height="60">
      <tr style="padding-top:10">
        <td height="24">&nbsp;</td>
      </tr>
    </table>
	<% end if %></td>
  </tr>
  <tr>
    <td height="24">&nbsp;</td>
  </tr>
  <tr>
    <td height="17"  align="center"><table width="1030" border="0" cellpadding="0" cellspacing="0" >
    <tr>
		<td class='title'><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle" > <%=request.cookies("custname")%> > 광고매체 보고서 관리 </td>
    </tr>
    </table></td>
  </tr>
  <tr>
    <td height="27">&nbsp;</td>
  </tr>
  <tr>
    <td valign="top" class="bdpdd" align="center">

			<table width="1030" height="35" border="0" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="13" valign="top" ><img src="/images/bg_search_left.gif" width="13" height="35" ></td>
                  <td background="/images/bg_search.gif"> <img src="/images/icon_dot_search.gif" width="4" height="3" align="absmiddle"> 검색년월
				  <%call getyear(cyear)%> <%call getmonth(cmonth)%> &nbsp; &nbsp; <a href="#" onClick="go_search();"><img src="/images/btn_search.gif" width="39" height="20" align="top" ></a> </td>
                  <td  align="right" background="/images/bg_search.gif" >&nbsp; </td>
                  <td width="13" ><img src="/images/bg_search_right.gif" width="13" height="35"></td>
                </tr>
        </table>

	</td>
  </tr>
  <tr>
    <td height="15">&nbsp;</td>
  </tr>
  <tr>
    <td height="600"  align='center'  valign="top">
	<!-- -->
	  <table border="0"width="1030" cellpadding="0" cellspacing="0" bordercolor="#8D652B" id="contact">
	  <thead>
			<tr height='30' align='center'>
				<th  width="30" class="hd left">No</th >
				 <th  width="240" class="hd center">계약명</th >
				 <th  width="80" class="hd center">시작일자</th >
				 <th  width="80" class="hd center">종료일자</th >
				 <th  width="80" class="hd center">월지급액</th >
				 <th  width="50" class="hd center">수량</th >
				 <th  width="320" class="hd center">매체위치</th >
				 <th  width="110" class="hd center">광고주</th >
				 <th  width="30" class="hd center">파일</th >
				 <th  width="40" class="hd right">관리</th >
			</tr>
		</thead>
		<tbody id='tbody'>
		<%
				Dim intLoop : intLoop = 1
				Do Until rs.eof
		%>
			<tr height='32'>
				<td  class="hd none" style='text-align:center;'  width="30"> <%=intLoop%></td>
				<td  class="hd none" style='text-align:left;' width="240"><span title='<%=title%>'><%=cutTitle(title, 40)%></span></td>
				<td  class="hd none" style='text-align:center;' width="80"><%=startdate%></td>
				<td  class="hd none" style='text-align:center;' width="80"><%=enddate%></td>
				<td  class="hd none" colspan='8'><table  width='620' border=0 style="table-layout:fixed;">
					<%
						rs2.Filter = "contidx="&contidx
						Do Until rs2.eof
					%>
						<tr height='32'>
							<td  width="80" style='text-align:right; padding-right:10px;'><%=FormatNumber(expense,0)%></td>
							<td  width="50"   style='text-align:center;'><%=FormatNumber(qty,0)%> <%=unit%></td>
							<td  width="320" style='padding-left:3px;'><%=locate%></td>
							<td  width="110" style='padding-left:3px;'><%=getcustname(custcode)%></td>
							<td  width="30"><% If Not IsNull(filename) Then response.write "<a href='#' onclick=""download('"&filename&"');""><img src='/images/m_ppt.gif' width='16' height='16' title='"&reportname&"'>" End If %></td>
							<td  width="40"><% if edate >= date then %><a href="#" onclick="getupload(<%=mdidx%>, '<%=title%>','<%=locate%>','c'); return false;"><img src='/images/m_upload.gif' width='16' height='16' alt="보고서 (재)등록"  ></a> <a href="#" onclick="getupload(<%=mdidx%>, '<%=title%>','<%=locate%>','d'); return false;"><img src='/images/m_delete.gif' width='16' height='15' alt="보고서삭제" hspace=1></a><% end if %> </td>
						</tr>
					<%
									rs2.movenext
								Loop
					%>
						</table></td>
			</tr>
			<%
						intLoop = intLoop + 1
						rs.movenext
					Loop
			%>
		</tbody>
        </table>
	<!--  -->
	</td>
  </tr>
  <tr>
    <td height="24">&nbsp; </td>
  </tr>
  <tr>
    <td height="24"><img src="/images/pop_bottom.gif" width="1240" height="71" align="absmiddle"></td>
  </tr>
</table>
</form>
</body>
</html>