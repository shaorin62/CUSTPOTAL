<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<!--#include virtual="/hq/outdoor/inc/Function.asp" -->

<%
	response.cookies("midx") = request("midx")

	dim gotopage : gotopage = request("gotopage")
	dim searchstring : searchstring = request("searchstring")
	dim pwd : pwd = request("txtpassword")
	dim ridx : ridx = request("ridx")
	dim midx : midx = request("midx")

	dim objrs, sql
	if midx = "" then
		sql = "select min(midx) midx from dbo.wb_menu_mst where custcode is null"
		call get_recordset(objrs, sql)
		if not objrs.eof then
			midx = objrs("midx")
		else
			midx = 0
		end if
		objrs.close
	end if

	
	'메뉴의 특성을 파악한다 코멘트 , 파일업로드 , 댓글 가능 한지...
	sql  = "select  isfile, iscomment, isemail, attr02 from dbo.wb_menu_mst where midx = " & midx

	call get_recordset(objrs, sql)

	dim isfile : isfile = objrs("isfile").value
	dim iscomment : iscomment = objrs("iscomment").value
	dim isemail : isemail = objrs("isemail").value
	dim attr02 : attr02 = objrs("attr02").value

	objrs.close
	sql  = "select  title, contents, mail, tomail, password, cuser, highcategory, category, custcode, cyear, cmonth from dbo.wb_report where ridx = " & ridx
	call get_recordset(objrs, sql)
	if objrs.eof then
		response.write "<script> alert('삭제된 게시물입니다.'); this.close();</script>"
	end if


	dim title, content, mail, tomail, attachfile, attachfile2, attachfile3, c_user, password, highcategory, category, custcode, cyear, cmonth
	title = objrs("title")
	content = objrs("contents")
	mail = objrs("mail")
	tomail = objrs("tomail")
	password = objrs("password")
	c_user = objrs("cuser")
	highcategory = objrs("highcategory")
	category = objrs("category")
	custcode = objrs("custcode")
	cyear = objrs("cyear")
	cmonth = objrs("cmonth")


	objrs.close
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<script language="javascript" src="/js/ajax.js"></script>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>

<body bgcolor="#857C7A"  oncontextmenu="return false">
<form >
<input type="hidden" name="ridx" value="<%=ridx%>">
<input type="hidden" name="midx" value="<%=midx%>">
<input type="hidden" name="userid" value="<%=request.Cookies("userid")%>">
<table width="640" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> <%=title%> </td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"> </td>
  </tr>
</table>
<table width="640" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="22"><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp; </td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td  bgcolor="#FFFFFF">
	<!--  -->
	<% if isnull(password) or pwd = password or c_user = request.cookies("userid")  then %>

	  <table border="0" cellpadding="0" cellspacing="0" align="center" width="588">
	  <% If midx = 16 Then %>
		  <tr>
			<td class="hw">대분류/중분류</td>
			<td class="bw bbd"><span id='highcategory'>대분류</span> <span id='category'>중분류</span></td>
		  </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
		 <% ElseIf attr02 = "inter" Then %>
		  <tr>
            <td class="hw">광고주</td>
			<td class="bw bbd">
				<table border="0" cellpadding="0" cellspacing="0" align="left" width="100%">
					<tr>
			            <td width="240px"><span id='custcode2'>광고주</span></td>
						<td class="hw" width="10px">년월</td>
						<td class="bw bbd"><%call getyear(cyear)%>&nbsp;<%call getmonth(cmonth)%></td>
					</tr>				
				</table>
			</td>
          </tr>
		   <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
		<% End If %>
		  <tr>
            <td class="hw">제 목</td>
            <td class="bw bbd"><%=title%></td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">내 용</td>
            <td class="bw bbd" style="padding-top:5px;padding-bottom:5px;"><div style="width:448px; height:360px;overflow-y:scroll;" id="txtcontents"><%if not isnull(content) then response.write replace(replace(content, chr(13)&chr(10), "<br>"), "''", "'")%></div></td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
         <!-- <tr id="mail" style="display:none;">
            <td class="hw" >담당 메일</td>
            <td class="bw bbd"><%=mail%> </td>
          </tr>
		  <tr id="mailline" style="display:none;">
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr id="tomail" style="display:none;">
            <td class="hw">받는 메일</td>
            <td class="bw bbd"><%=tomail%></td>
          </tr>
		  <tr id="tomailline" style="display:none;">
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>-->
		  <%
			sql = "select attachfile from dbo.wb_report_pds where ridx =  "& ridx
			call get_recordset(objrs, sql)

			if not objrs.eof then
		%>
          <tr id="attachfile" style="display:none;">
            <td class="hw"> 첨부 파일 </td>
            <td class="bw bbd">
			<% do until objrs.eof %>
			<span onClick="checkForDownload('<%=objrs("attachfile")%>');" class="styleLink"> <img src="/images/ico_attach.gif" width="7" height="12" hspace="3"> <%=objrs("attachfile")%></span>
			<%
				objrs.moveNext
				loop
			%>
			</td>
          </tr>
		  <tr id="attachfileline" style="display:none;">
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
		  <%
			 end if
			 objrs.close
		%>
		  <tr>
				<td colspan="2"  height="50" valign="bottom" align="right">
				<%if iscomment then %> <img src="/images/btn_comment_reg.gif" width="78" height="18"  vspace="5" style="cursor:hand" onClick="pop_comment_reg();"><% end if%><% if request.cookies("userid") = c_user or UCase(request.cookies("class")) = "A"  or UCase(request.cookies("class")) = "N"  then %><img src="/images/btn_edit.gif" width="59" height="18" hspace="10" vspace="5" style="cursor:hand" onClick="get_report_edit();"><img src="/images/btn_delete.gif" width="59" height="18" vspace="5" style="cursor:hand" onClick="get_report_delete();">
				<% end if%><img src="/images/btn_close.gif" width="57" height="18" vspace="5" style="cursor:hand" onClick="set_close();" hspace="10" ></td>
			</tr>
			<%
				sql = "select cidx, ridx, comment, attachfile, cuser, cdate from dbo.wb_report_comment where ridx="&ridx & " order by cidx desc"
				call get_recordset(objrs, sql)

				dim cidx, comment, c_attachfile, c_date

				do until objrs.eof

				cidx = objrs("cidx")
				comment = objrs("comment")
				c_attachfile = objrs("attachfile")
				c_user = objrs("cuser")
				c_date = objrs("cdate")
			%>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw" style="padding-top:5px;padding-bottom:5px;"><%=c_user%><br><%=c_date%></td>
            <td class="bw bbd" style="padding-top:5px;padding-bottom:5px;"><%=comment%>
			<% if not isnull(c_attachfile) then %><p><span onClick="checkForDownload('<%=c_attachfile%>');" class="styleLink"> <img src="/images/ico_attach.gif" width="7" height="12" vspace="3" align="absmiddle"> <%=c_attachfile%> </span> <% end if%> <%if request.cookies("userid") = c_user or request.cookies("class") = "A"  or UCase(request.cookies("class")) = "N"  then %><img src="/images/reply_view_lineone_close.gif" width="11" height="11" vspace="3" align="absmiddle" onclick="get_comment_delete_proc(<%=cidx%>);" class="stylelink" alt="댓글삭제"><% end if%>	</td>
          </tr>

		  <%
				objrs.movenext
				loop

				objrs.close
				set objrs = nothing
		  %>
      </table>
	  <% else %>

	  <table border="0" cellpadding="0" cellspacing="0" align="center" width="588">
          <tr>
            <td class="hw">제 목</td>
            <td class="bw bbd"><%=title%></td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">내 용</td>
            <td class="bw bbd" style="padding-top:5px;padding-bottom:5px;"><div style="width:448px; height:460px;overflow-y:scroll;" id="txtcontents">비공개 게시물입니다. <BR>비밀번호를 입력하세요</div></td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">비밀번호</td>
            <td class="bw bbd"><INPUT TYPE="password" NAME="txtpassword" maxlength="14"></td>
          </tr>
		  <tr>
				<td colspan="2"  height="50" valign="bottom" align="right"> <img src="/images/btn_confirm.gif" width="59" height="18" vspace="5" style="cursor:hand" onClick="check_form();"><img src="/images/btn_close.gif" width="57" height="18" vspace="5" style="cursor:hand" onClick="set_close();" hspace="10" ></td>
			</tr>
      </table>
	  <% end if%>
	<!--  -->
	</td>
    <td background="/images/pop_right_middle_bg.gif">&nbsp;</td>
  </tr>
  <tr>
    <td><img src="/images/pop_left_bottom_bg.gif" width="22" height="25"></td>
    <td background="/images/pop_center_bottom_bg.gif">&nbsp;</td>
    <td><img src="/images/pop_right_bottom_bg.gif" width="23" height="25"></td>
  </tr>
</table>
</form>
</body>
</html>
<script language="JavaScript">
<!--
	function pop_comment_reg() {
		var url = "pop_comment_reg.asp?ridx=<%=ridx%>";
		var name = "";
		var opt = "width=540, height=233, resziable=no, scrollbars = no, status=yes, top=100, left=770";
		window.open(url, name, opt);
	}

	function get_report_edit() {
			location.href="pop_report_edit.asp?ridx=<%=ridx%>&midx=<%=midx%>&gotopage=<%=gotopage%>&searchstring=<%=searchstring%>";
	}

	function get_report_delete() {
		if (confirm("게시물에 등록된 댓글, 첨부자료도 함께 삭제됩니다.\n\n게시물을 삭제하시겠습니까?")) {
			location.href="report_delete_proc.asp?ridx=<%=ridx%>" ;
		}
	}

	function checkForDownload(name) {
		location.href="download.asp?filename="+name;
	}

	function get_comment_delete_proc(cidx) {
		if (confirm("등록된 첨부파일도 함께 삭제됩니다.\n\n선택한 댓글을 삭제하시겠습니까?")) {
			location.href = "comment_delete_proc.asp?cidx="+cidx+"&ridx=<%=ridx%>&midx=<%=midx%>";
		}
	}


	function set_close() {
		window.opener.document.location.href = window.opener.document.URL;
		this.close();
	}

	function check_private() {
		var chk = document.getElementById("chkpassword")
		if (chk.checked) document.getElementById("password").style.display = "block";
		else document.getElementById("password").style.display = "none";

		document.getElementById("txtpassword").value = "";
	}

	function check_form() {
		var frm = document.forms[0];

		if (frm.txtpassword.value == "") {
			alert("비밀번호를 입력하세요");
			frm.txtpassword.focus();
			return false;
		}

		var password = "<%=password%>";
		if (password != frm.txtpassword.value) {
			alert("비밀번호가 잘못되었습니다.\n\n비밀번호를 정확하게 입력하세요");
			frm.txtpassword.value = "";
			frm.txtpassword.focus();
			return false;
		}

		frm.action = "pop_report_view.asp";
		frm.method = "POST";
		frm.submit();
	}

	window.onload=function () {
		self.focus();
		<% if isnull(password) or password = pwd or c_user = request.cookies("userid") then %>
		<% if isemail then %>
			//document.getElementById("mail").style.display = "block";
			//document.getElementById("mailline").style.display = "block";
			//document.getElementById("tomail").style.display = "block";
			//document.getElementById("tomailline").style.display = "block";
			//document.getElementById("mailexp").style.display = "block";
			document.getElementById("txtcontents").style.height = parseInt(document.getElementById("txtcontents").style.height) - 82+"px";
		<% end if%>
		<% if isfile then %>
			if (document.getElementById("attachfile")) {
				document.getElementById("attachfile").style.display = "block";
				document.getElementById("attachfileline").style.display = "block";
			}
			document.getElementById("txtcontents").style.height = parseInt(document.getElementById("txtcontents").style.height) + 118 +"px";
		<% end if%>

		<% if midx=16 then %>
				_sendRequest("/inc/getreporthighcategory.asp", "highcategory=<%=highcategory%>", _gethighcategorycombo, "GET");
				_sendRequest("/inc/getreportcategory.asp", "highcategory=<%=highcategory%>&category=<%=category%>", _getcategorycombo, "GET");
				document.getElementById("cmbhighcategory").disabled = true;
				document.getElementById("cmbcategory").disabled = true;
			<% end if%>

			<% if attr02 = "inter" then %>
				_sendRequest("/inc/getcustcombo_report.asp", "custcode=<%=custcode%>", _getcustcombo_report, "GET");
			<% end if%>	
		<% end if%>

	}

// 추가
	function gethighcategorycombo() {
			// 광고주 콤보 박스 가져오기
			var highcategory = null;
			var params = "highcategory="+highcategory;

			sendRequest("/inc/getreporthighcategory.asp", params, _gethighcategorycombo, "GET");
		}

		function _gethighcategorycombo() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
				
						var highcategory = document.getElementById("highcategory");
						highcategory.innerHTML = xmlreq.responseText ;
						getcategorycombo();
				}
			}
		}

		function getcategorycombo() {
			// 운영팀 콤보 박스 가져오기
			var highcategory = document.getElementById("cmbhighcategory").value;
			var category = null;

			var params = "highcategory="+highcategory+"&category="+category;

			sendRequest("/inc/getreportcategory.asp", params, _getcategorycombo, "GET");
		}

		function _getcategorycombo() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
						var category = document.getElementById("category");
						category.innerHTML = xmlreq.responseText ;
				}
			}
		}

		function getcustcombo_report() {
			// 광고주 콤보 박스 가져오기
			var scope = null;
			var custcode = null;
			var params = "scope="+scope+"&custcode="+custcode;
			sendRequest("/inc/getcustcombo_report.asp", params, _getcustcombo_report, "GET");
		}

		function _getcustcombo_report() {
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
						var custcode = document.getElementById("custcode2");
						custcode.innerHTML = xmlreq.responseText ;
						document.getElementById("cmbcustcode").disabled  = true;
						document.getElementById("cyear").disabled  = true;
						document.getElementById("cmonth").disabled  = true;
				}
			}
		}

	function checkForSearch() {
	}
	//-->
</script>

