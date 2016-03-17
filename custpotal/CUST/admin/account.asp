<%@ language="vbscript" codepage="65001" %>
<!--#include virtual="/cust/outdoor/inc/Function.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%

		Dim sql : sql = "select distinct a.highcustcode, a.custname from sc_cust_hdr a inner join sc_cust_dtl b on a.highcustcode=b.highcustcode where a.medflag='B' and a.use_flag=1 and b.med_out = 1 order by a.custname"
		Dim cmd : Set cmd = server.CreateObject("adodb.command")
		cmd.activeconnection = application("connectionstring")
		cmd.commandText = sql
		cmd.commandType = adCmdText
		Dim rs : Set rs = cmd.execute
		Set cmd = Nothing

		Sub getmedcode
			Do Until rs.eof
				response.write "<option value='" & rs(0) & "'>" & rs(1) & "</option>"
				rs.movenext
			Loop
		End Sub
%>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<div style='margin-top:10px;'>
<TABLE  width="100%">
	<TR>
		<TD ><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle" id="subtitle"> 계정관리</span></TD>
		<TD  align="right" valign="top">  <span class="navigator"  id="navigate">관리모드 &gt; 계정관리</span></TD>
	</TR>
</TABLE>
</div>

<br><br>

<link href="/style.css" rel="stylesheet" type="text/css">
	<table width="1030" height="31" border="0" cellpadding="0" cellspacing="0" >	
		<input type="hidden" name="mstruserid">
		<tr>
			<td >
				<table border="0" cellspacing="0" cellpadding="0">
					<tr>
						<td width="380" height='30'> <img src='/images/m_arw.gif' width='5' height='8' hspace='2'>계정 리스트 </td>
						<td  width='75' align="center" ></td>
						<td width='240' ><img src='/images/m_arw.gif' width='5' height='8' hspace='2'> 광고주 리스트 </td>
						<td  width='75' align="center" ></td>
						<td width='240'><img src='/images/m_arw.gif' width='5' height='8' hspace='2'> 팀 리스트</td>
					</tr>
					<tr>
						<td height="2"  colspan="5"></td>
					</tr>
					<tr>
						<td>
							<table width="100%" border="1" cellspacing="0" cellpadding="0" class="header" bordercolor="#8D652B">
							  <tr>
								<td width="30" height="25" align="center" class="header">No</td>
								<td width="120" align="center" >아이디</td>
								 <td width="120" align="center" >이름</td>
								<td width="110" align="center" >권한</td>
							  </tr>
							 </table>
						</td>
						<td  width='75' align="center" ></td>
						<td>
							<table width="100%" border="1" cellspacing="0" cellpadding="0" class="header" bordercolor="#8D652B">
							  <tr>
								<td align="center" class="header" height="25" >광고주명</td>
							  </tr>
							 </table>
						</td>
						<td  width='75' align="center" ></td>
						<td>
							<table width="100%" border="1" cellspacing="0" cellpadding="0" class="header" bordercolor="#8D652B">
							  <tr>
								<td align="center" class="header" height="25" >팀명</td>
							  </tr>
							 </table>
						</td>
					</tr>

					<tr>
						<td height="6"  colspan="5"></td>
					</tr>

					<tr>
						<td>
							<table width="100%" border="0" cellspacing="0" cellpadding="0" class="header" bordercolor="#8D652B">
							  <tr>
								<td align="center" class="header">
									<iframe id="frmuser" src="account_fuserid.asp" frameborder="1" width="390" height="400" leftmargin="0" topmargin="0" scrolling="auto"></iframe>
								</td>
							  </tr>
							  <tr>
								<td align="right" class="header">
									<a onclick="getemployee('c'); return false;" style="cursor:hand;">[추가] </a> |  
									<a onclick="getemployee('d'); return false;" style="cursor:hand;">[삭제] </a></td>
							  </tr>
							 </table>
						</td>
						<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src='/images/btn_next.gif' width='14' height='15' alt="추가" align='absmiddle' ></td>
						<td>
							<table width="100%" border="0" cellspacing="0" cellpadding="0" class="header" bordercolor="#8D652B">
							  <tr>
								<td align="center" class="header">
									<iframe id="frmcust" src="account_fcust.asp" frameborder="1" width="240" height="400" leftmargin="0" topmargin="0" scrolling="auto"></iframe>
								</td>
							  </tr>
							  <tr>
								<td align="right" class="header">
									<a onclick="getemployee_cust('c'); return false;" style="cursor:hand;">[추가] </a> |  
									<a onclick="getemployee_cust('d'); return false;" style="cursor:hand;">[삭제] </a></td>
							  </tr>
							 </table>
						</td>
						<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<img src='/images/btn_next.gif' width='14' height='15' alt="추가" align='absmiddle' ></td>
						<td>
							<table width="100%" border="0" cellspacing="0" cellpadding="0" class="header" bordercolor="#8D652B">
							  <tr>
								<td align="center" class="header">
									<iframe id="frmtim" src="account_ftim.asp" frameborder="1" width="240" height="400" leftmargin="0" topmargin="0" scrolling="auto"></iframe>
								</td>
							  </tr>
							  <tr>
								<td align="right" class="header">
									<a onclick="getemployee_tim('c'); return false;" style="cursor:hand;">[추가] </a> |  
									<a onclick="getemployee_tim('d'); return false;" style="cursor:hand;">[삭제] </a></td>
							  </tr>
							 </table>
						</td>
					</tr>


					<!--삭제아이프레임-->
					<tr>
						<td>
							<table width="100%" border="0" cellspacing="0" cellpadding="0" class="header" bordercolor="#8D652B">
							  <tr>
								<td align="center" class="header">
									<iframe id="frmdeleteproc" src="" frameborder="1" width="0" height="0" leftmargin="0" topmargin="0" scrolling="auto" style="visibility: hidden"></iframe>
								</td>
							  </tr>
							
							 </table>
						</td>
					</tr>
					
				</table>
		</td>
	</tr>
	</table>

	<br><br><br>