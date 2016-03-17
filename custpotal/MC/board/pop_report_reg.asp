<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<%
	dim midx : midx = request("midx")
	dim custcode : custcode = request("custcode")
	dim flag : flag = request("flag")
	dim highcategory : highcategory = request("highcategory")
	dim category : category = request("category")


	if midx = "" then
		sql = "select min(midx) from dbo.wb_menu_mst where custcode is null"
		call get_recordset(objrs, sql)
		if not objrs.eof then
			midx = objrs(0)
		else
			midx = 0
		end if
		objrs.close
	end if

	dim objrs, sql
	sql  = "select  title, isfile, iscomment, isemail from dbo.wb_menu_mst where midx = " & midx

	call get_recordset(objrs, sql)

	dim title : title = objrs("title").value
	dim isfile : isfile = objrs("isfile").value
	dim iscomment : iscomment = objrs("iscomment").value
	dim isemail : isemail = objrs("isemail").value

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
<form enctype="multipart/form-data">
<input type="hidden" name="midx" value="<%=midx%>">
<input type="hidden" name="custcode" value="<%=custcode%>">
<input type="hidden" name="flag" value="<%=flag%>">
<input type="hidden" name="thighcategory" value="<%=highcategory%>">
<input type="hidden" name="tcategory" value="<%=category%>">
<input type="hidden" name="userid" value="<%=request.Cookies("userid")%>">
<table width="640" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> <%=title%> 등록 </td>
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
	  <table border="0" cellpadding="0" cellspacing="0" align="center" width="588">
		<% If midx = 16 Then %>
		  <tr>
            <td class="hw">대분류/중분류</td>
            <td class="bw bbd"><span id='highcategory'>대분류</span> <span id='category'>중분류</span></td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
		<% End If %>
          <tr>
            <td class="hw">제 목</td>
            <td class="bw bbd"><input name="txttitle" type="text" id="txttitle" style="width:430px;" class="kor" maxlength="50"></td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr>
            <td class="hw">내 용</td>
            <td class="bw bbd" style="padding-top:5px;padding-bottom:5px;"><textarea name="txtcontents"  id="txtcontents" style="width:430px;height:310px;" class="kor" onblur="check_length(this, 1000);"></textarea></td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr id="mail" style="display:none;">
            <td class="hw" >담당 메일</td>
            <td class="bw bbd"><input name="txtmail" type="text" id="txtmail" style="width:430px;" class="eng"> </td>
          </tr>
		  <tr id="mailline" style="display:none;">
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr id="tomail" style="display:none;">
            <td class="hw">받는 메일</td>
            <td class="bw bbd"><input name="txttomail" type="text" id="txttomail" style="width:430px;" class="eng"></td>
          </tr>
		  <tr id="tomailline" style="display:none;">
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr id="attachfile" style="display:none;">
            <td class="hw">첨부 파일</td>
            <td class="bw bbd"><input type="file" name="txtfile" style="width:430px;"></td>
          </tr>
		  <tr id="attachfileline" style="display:none;">
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr id="attachfile2" style="display:none;">
            <td class="hw">첨부 파일</td>
            <td class="bw bbd"><input type="file" name="txtfile" style="width:430px;"></td>
          </tr>
		  <tr id="attachfileline2" style="display:none;">
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr id="attachfile3" style="display:none;">
            <td class="hw">첨부 파일</td>
            <td class="bw bbd"><input type="file" name="txtfile" style="width:430px;"></td>
          </tr>
		  <tr id="attachfileline3" style="display:none;">
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr  >
            <td class="hw">비공개 <INPUT TYPE="checkbox" NAME="chkpassword" ID="chkpassword" onclick="check_public()"> </td>
            <td class="bw bbd"> <span id="password" style="display:none;">비밀번호 <INPUT TYPE="text" NAME="txtpassword" ID="txtpassword"></span></td>
          </tr>
		  <tr>
			<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
          <tr id="mailexp" style="display:none;">
            <td colspan="2" height="20" valign="bottom">* 받는메일이 1명 이상인 경우에는 ,로 구분하여 입력하세요</td>
          </tr>
		  <tr>
				<td colspan="2"  height="50" valign="bottom" align="right"> <img src="/images/btn_close.gif" width="57" height="18" vspace="6" style="cursor:hand" onClick="set_close();" hspace="10" >
	</td>
      </table>
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
		function check_submit() {
			var frm = document.forms[0];
			if (frm.txttitle.value == "") {
				alert("제목을 입력하세요");
				frm.txttitle.focus();
				return false;
			}
			if (frm.txtcontents.value == "") {
				alert("내용을 입력하세요");
				frm.txtcontents.focus();
				return false ;
			}
			if (frm.chkpassword.checked) {
				if (frm.txtpassword.value == "") {
					alert("비공개 설정인 경우 비밀번호를 입력하셔야 합니다.");
					frm.txtpassword.focus();
					return false;
				}
			}

			frm.action = "report_reg_proc.asp";
			frm.method = "post";
			frm.submit();

		}

		function set_reset() {
			document.forms[0].reset();
		}

		function set_close() {
			this.close();
		}

		function check_public() {
			var chk = document.getElementById("chkpassword");
			if (chk.checked) {
				document.getElementById("password").style.display = "block";
			}	else {
				document.getElementById("password").style.display = "none";
			}
			document.getElementById("txtpassword").value = "";
		}

		window.onload=function () {
			self.focus();
			<% if isemail then %>
				document.getElementById("mail").style.display = "block";
				document.getElementById("mailline").style.display = "block";
				document.getElementById("tomail").style.display = "block";
				document.getElementById("tomailline").style.display = "block";
				document.getElementById("mailexp").style.display = "block";
				document.getElementById("txtcontents").style.height = parseInt(document.getElementById("txtcontents").style.height) - 82+"px";
			<% end if%>
			<% if isfile then %>
				document.getElementById("attachfile").style.display = "block";
				document.getElementById("attachfileline").style.display = "block";
				document.getElementById("attachfile2").style.display = "block";
				document.getElementById("attachfileline2").style.display = "block";
				document.getElementById("attachfile3").style.display = "block";
				document.getElementById("attachfileline3").style.display = "block";
				document.getElementById("txtcontents").style.height = parseInt(document.getElementById("txtcontents").style.height) + 56+"px";
			<% end if%>
			<% if midx=16 then %>
				_sendRequest("/inc/getreporthighcategory.asp", "highcategory=<%=highcategory%>", _gethighcategorycombo, "GET");
				_sendRequest("/inc/getreportcategory.asp", "highcategory=<%=highcategory%>&category=<%=category%>", _getcategorycombo, "GET");
				document.getElementById("cmbhighcategory").attachEvent("onchange", getcategorycombo);
			<% end if%>
		}

	function check_length(x, num) {
		var maxlen = num ;
		if (x.value.length > maxlen ) {
			alert(num + "자 이내로 입력하세요");
			x.value = x.value.substring(0, maxlen);
			return false ;
		}
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

	function checkForSearch() {
	}
	//-->
	</script>
