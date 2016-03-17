<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->
<%
	dim gotopage : gotopage = request("gotopage")
	dim searchstring : searchstring = request("searchstring")

	dim mdidx : mdidx = request("mdidx")
	'response.write mdidx
	dim sql : sql = "select  mdidx, title, custcode, categoryidx, unit, region, locate, map, comment, cuser, cdate, uuser, udate from dbo.WB_MEDIUM_MST where mdidx = " & mdidx
	dim objrs
	call get_recordset(objrs, sql)

	if objrs.eof then response.write "<script> alert('삭제 또는 잘못된 매체정보 입니다.'); location.href='medium_list.asp?gotopage=" & gotopage & "&serarchstring=" & searchstring

	dim title : title = objrs.fields("title")
	dim custcode : custcode = objrs.fields("custcode")
	dim categoryidx : categoryidx = objrs.fields("categoryidx")
	dim unit : unit = objrs.fields("unit")
	dim region : region = objrs.fields("region")
	dim locate : locate = objrs.fields("locate")
	dim map : map = objrs.fields("map")
	dim comment : comment = objrs.fields("comment")

	objrs.close
	set objrs = nothing
%>
<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  oncontextmenu="return false">
<form enctype="multipart/form-data">
<!--#include virtual="/od/top.asp" -->
  <table id="Table_01" width="1240" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td rowspan="2" valign="top"><!--#include virtual="/od/left_outdoor_menu.asp"--></td>
      <td height="65"><img src="/images/middle_navigater_outdoor.gif" width="1030" height="65" alt=""></td>
    </tr>
    <tr>
      <td height="400" align="left" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="24">&nbsp;</td>
          </tr>
          <tr>
            <td height="19" valign="top" class="navigator" > 옥외관리 &gt; 매체관리 &gt; 매체정보변경</td>
          </tr>
          <tr>
            <td height="17"><img src="/images/ico_subtitle.gif" width="28" height="17" align="absmiddle"><span class="subtitle">매체정보변경</span></td>
          </tr>
          <tr>
            <td height="27">&nbsp;</td>
          </tr>
          <tr>
            <td valign="top" class="bdpdd">
			<table border="0" cellpadding="0" cellspacing="0" >
			<tr>
				<td colspan="2" bgcolor="#cacaca" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd">매체명</td>
				<td class="tdbd"><input name="txttitle" type="text" size="70" maxlength="100" class="kor" value="<%=title%>"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd">매체사</td>
				<td class="tdbd"> <% call get_medium_custcode(custcode, null)%></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd" >매체분류</td>
				<td  class="tdbd""><span id="category"><% call get_medium_catetory(categoryidx)%></span> <img src="/images/btn_find.gif" width="39" height="20" align="absmiddle" class="stylelink" onClick="pop_medium_category();"> <input name="txtcategoryidx" type="hidden" id="txtcategoryidx" value="<%=categoryidx%>"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd">수량단위</td>
				<td class="tdbd"><input name="rdounit" type="radio" value="구좌" checked onclick="check_valid_unit();" <%if unit = "구좌" then response.write "checked"%>> 구좌 <input name="rdounit" type="radio" value="기" onclick="check_valid_unit();" <%if unit = "기" then response.write "checked"%>> 기 <input name="rdounit" type="radio" value="면" onclick="check_valid_unit();" <%if unit = "면" then response.write "checked"%>> 면 <input name="rdounit" type="radio" value="기타" onclick="check_valid_unit();" <%if unit <> "구좌" and unit <> "면" and unit <> "기" then response.write "checked"%>> 직접입력 <input type="text" name="txtunit" value="<%if unit <> "구좌" and unit <> "면" and unit <> "기" then response.write unit %>"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd">설치지역</td>
				<td class="tdbd"><%call get_region_code(region, null) %> </td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd">위치정보</td>
				<td class="tdbd"><input name="txtlocate" type="text" size="70" maxlength="100" class="kor" value="<%=locate%>"></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd">약도파일</td>
				<td class="tdbd"><%if not isnull(map) then%><img src="/pds/media/<%=map%>" width="500" height="350" border="1" alt="" vspace="11"><%end if%><input name="txtmap" type="file" id="txtmap" size="50" ></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			<tr>
				<td class="tdhd">특이사항</td>
				<td class="tdbd"><textarea name="txtcomment" rows="5" style="width:612px;" class="kor"><%if not isnull(comment) then response.write comment %></textarea></td>
			</tr>
			<tr>
				<td colspan="2" bgcolor="#E7E7DE" height="1"></td>
			</tr>
			</table>
			  <table width="756" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="50%" height="50" align="left" valign="bottom"><a href="/od/outdoor/md_list.asp"><img src="/images/btn_list.gif" width="59" height="20" border="0"></a></td>
                  <td width="50%" align="right" valign="bottom"><img src="/images/btn_save.gif" width="59" height="20" hspace="10" vspace="5" style="cursor:hand" onClick="checkForSubmit();"><img src="/images/btn_init.gif" width="67" height="20" vspace="5" style="cursor:hand" onClick="set_reset();"></td>
                </tr>
              </table></td>
          </tr>
      </table>
	  </td>
    </tr>
  </table>
<!--#include virtual="bottom.asp" -->
<input type="hidden" name="mdidx" value="<%=mdidx%>">
<input type="hidden" name="gotopage" value="<%=gotopage%>">
<input type="hidden" name="searchstring" value="<%=searchstring%>">
</form>
</body>
</html>
<script language="javascript">
<!--
	window.onload = function () {
		var frm = document.forms[0];
		if (!frm.rdounit[3].checked)
			frm.txtunit.disabled = true;
	}
	function checkForSubmit() {
		var frm = document.forms[0];
		if (frm.txttitle.value == "") {
			alert("매체명은 필수입력 사항입니다.");
			frm.txttitle.focus();
			return false;
		}
		if (frm.selcustcode.value == "") {
			alert("매체사는 필수입력 사합니다.");
			frm.selcustcode.focus();
			return false;
		}
		if (frm.txtcategoryidx.value == "") {
			alert("매체분류는 필수입력 사항입니다.");
			pop_medium_category();
			return false;
		}

		frm.method = "POST";
		frm.action = "medium_edit_proc.asp";
		frm.submit();
	}

	function pop_medium_custcode() {
		var url = "pop_medium_custcode.asp";
		var name = "pop_medium_custcode";
		var opt = "width=500, height=500, resziable=no, scrollbars = yes, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function pop_medium_category() {
		var url = "pop_medium_category.asp";
		var name = "pop_medium_category";
		var opt = "width=500, height=500, resziable=no, scrollbars = yes, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function check_valid_unit() {
		var frm = document.forms[0];
		var bln = frm.rdounit[3].checked ;
		frm.txtunit.disabled = !bln;
		if (bln) {
			frm.txtunit.focus();
			frm.txtunit.value = "";
		}
	}

	function set_reset() {
		document.forms[0].reset();
	}
//-->
</script>