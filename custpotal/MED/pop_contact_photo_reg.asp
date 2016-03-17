<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim idx : idx = request("idx")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")

	dim objrs, sql

	sql = "select photo_1, photo_2, photo_3, photo_4 from dbo.wb_contact_md_dtl_account where idx = " & idx & " and cyear = '" & cyear & "' and cmonth = '" & cmonth & "' "
	call get_recordset(objrs, sql)

	dim photo_1, photo_2, photo_3, photo_4
	if not objrs.eof then
	photo_1 = objrs("photo_1")
	photo_2 = objrs("photo_2")
	photo_3 = objrs("photo_3")
	photo_4 = objrs("photo_4")
	end if
	objrs.close

	sql = "select m.contidx, title, locate from dbo.wb_contact_mst m inner join dbo.wb_contact_md m2 on m.contidx = m2.contidx inner join dbo.wb_contact_md_dtl d on m2.sidx = d.sidx where idx = " & idx
	call get_recordset(objrs, sql)

	dim title : title = objrs("title")
	dim locate : locate = objrs("locate")
	dim contidx : contidx = objrs("contidx")
	dim intLoop, intCheck

	objrs.close

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
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

<body  background="/images/pop_bg.gif"  oncontextmenu="return false">
<form>
<table width="686" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="top"> <%=title%> <%if not isnull(locate) then response.write  " : " & locate %> </td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="686" border="0" cellspacing="0" cellpadding="0"  bgcolor="#FFFFFF">
  <tr>
    <td width="22" ><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td>
	<!--  -->
	  <input type="hidden" name="idx" value="<%=idx%>">
	  <input type="hidden" name="cyear" value="<%=cyear%>">
	  <input type="hidden" name="cmonth" value="<%=cmonth%>">
	  <input type="hidden" name="contidx" value="<%=contidx%>">
	  <table border="0" cellpadding="0" cellspacing="0" align="center" >
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
	  <tr >
       <td   align="center" valign="top"><img src="<%if photo_1 <> "" then response.write "/pds/media/"&photo_1& """ class='stylelink' " else response.write "/images/noimage.gif"%>" width="150" height="100" border="0" vspace="5" onClick="pop_medium_photo('<%=photo_1%>');"  ></td>
		<td  align="center" valign="top"><img src="<%if photo_2 <> "" then response.write "/pds/media/"&photo_2& """   class='stylelink' "else response.write "/images/noimage.gif"%>" width="150" height="100" border="0" vspace="5"  onClick="pop_medium_photo('<%=photo_2%>');"></td>
		<td   align="center" valign="top"><img src="<%if photo_3 <> "" then response.write "/pds/media/"&photo_3 &""" class='stylelink' "else response.write "/images/noimage.gif"%>" width="150" height="100" border="0"  vspace="5" onClick="pop_medium_photo('<%=photo_3%>');"></td>
		<td  align="center" valign="top"><img src="<%if photo_4 <> "" then response.write "/pds/media/"&photo_4 &""" class='stylelink' "else response.write "/images/noimage.gif"%>" width="150" height="100" border="0"  vspace="5" onClick="pop_medium_photo('<%=photo_4%>');"></td>
	  </tr>
		<%

		sql = "select idx, dtlIdx, cyear, cmonth , comment from dbo.wb_contact_photo_mst where dtlIdx = "&idx&" and cyear = '"&cyear&"' and cmonth = '"&cmonth&"' "

		call get_recordset(objrs, sql)

		if not objrs.eof then
			do until objrs.eof
		%>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="1"></td>
		  </tr>
		  <tr>
            <td  colspan="4" align="center" Height="31" ><B>&lt;<%=objrs("comment")%>&gt;</B></td>
		  </tr>
          <tr>
		  <%
			dim objrs2
			sql = "select idx, mstidx, chk, filename, note from dbo.wb_contact_photo_dtl where mstidx = " & objrs("idx")
			call get_recordset(objrs2, sql)

			if not objrs.eof then
			intCheck = 0
				For intLoop = 0 To 3
		  %>
			<% if objrs2.eof then %>
             <td class="bw" valign="top"  width="150"><img src="/images/noimage.gif" width="150" height="100" border="0" vspace="5" hspace="5"><br>&nbsp;<br>&nbsp;</td>
			 <% else %>
            <td class="bw"  valign="top"  width="150"><img src="<%if not isnull(objrs2("filename")) then response.write "/pds/media/"&objrs2("filename")& """ class='stylelink' onclick='pop_view_photo("&objrs2("idx")&")'" else response.write "/images/noimage.gif"%>" width="150" height="100" border="0" vspace="5" hspace="5"><br> <INPUT TYPE="checkbox" NAME="photoIdx" value="<%=objrs2("idx")%>" onclick="checkCount(this);" >  <br><%=replace(objrs2("note"), chr(13)&chr(10), "<br>")%></td>
			<%
				intCheck = intCheck + 1
				objrs2.movenext
				end if
				Next
			end if
				objrs2.close
				set objrs2 = nothing
			%>
          </tr>
		  <tr>
			<td colspan="4"  height="31" align="right" style="padding-right:22px;"> <% if intCheck <  4 then %><A HREF="#" onclick="pop_contact_photo_add(<%=idx%>, '<%=cyear%>','<%=cmonth%>');" class="pagesplit">추가</A><% end if%> </td>
		  </tr>
		  <tr>
			<td colspan="4" bgcolor="#E7E7DE" height="2"></td>
		  </tr>
		  <%
			objrs.movenext
			Loop
	end if
			objrs.close
			set objrs = nothing
			%>
		  <tr>
				<td colspan="4"  height="50" valign="bottom" align="right"> * 웹사이트에 게시할 사진 4장을 선택하세요. <img src="/images/btn_photo_reg.gif" width="78" height="18" class="stylelink" onClick="pop_photo_reg(<%=idx%>, '<%=cyear%>','<%=cmonth%>');"  hspace="5"><img src="/images/btn_save.gif" width="59" height="18"   style="cursor:hand" onClick="check_submit();" ><img src="/images/btn_close.gif" width="57" height="18"style="cursor:hand" onClick="set_close();" hspace="5" >
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
			var cnt  = 0 ;
			for (var i = 0 ; i < frm.photoIdx.length ; i++) {
				if (frm.photoIdx[i].checked) cnt += 1 ;
			}

			if (cnt == 0) {
				alert("웹사이트에 게시할 사진을 선택하세요");
				return false;
			}

			frm.action = "contact_photo_reg_proc.asp";
			frm.method = "post";
			frm.submit();

		}

		function pop_photo_reg(idx, cyear, cmonth) {
		var url = "pop_photo_reg.asp?idx="+idx+"&cyear="+cyear+"&cmonth="+cmonth;
		var name = "pop_photo_reg";
		var opt = "width=558, height=433, resizable=no, scrollbars=yes, status=yes, left=100, top=100";
		window.open(url, name, opt);
		}

		function pop_contact_photo_add(idx, cyear, cmonth) {
			// idx <= wb_contact_md_dtl idx
		var url = "pop_photo_add.asp?idx="+idx+"&cyear="+cyear+"&cmonth="+cmonth;
		var name = "pop_photo_add";
		var opt = "width=540, height=212, resizable=no, scrollbars=no, status=yes, left=100, top=100";
		window.open(url, name, opt);

		}

		function pop_view_photo(idx) {
			var url = "pop_medium_photo.asp?idx=<%=idx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>&photoIdx="+idx;
			var name = "pop_medium_photo";
			var opt = "width=668, height=620, resizable=no, scrollbars=yes, status=yes, left=100, top=100";
			window.open(url, name, opt);
		}

	function pop_medium_photo(photo) {
		if (photo != "") {
			var url = "/hq/outdoor/pop_medium_photo.asp?photo=" + photo+"&idx=<%=idx%>&cyear=<%=cyear%>&cmonth=<%=cmonth%>" ;
			var name = "pop_medium_photo";
			var opt = "width=668, height=550, resizable=no, scrollbars=yes, status=yes, left=100, top=100";
			window.open(url, name, opt);
		}
	}

		function checkCount(p) {
			var frm = document.forms[0];
			var cnt  = 0 ;
			for (var i = 0 ; i < frm.photoIdx.length ; i++) {
				if (frm.photoIdx[i].checked) cnt += 1 ;
			}

			if (cnt > 4 ) {
				alert("선택은 4개까지 가능합니다.");
				p.checked = false ;
			}
		}

		function set_reset() {
			document.forms[0].reset();
		}

		function set_close() {
			this.close();
		}

		window.onload=function () {
			self.focus();
		}
	//-->
	</script>
