<!--#include virtual="/MP/outdoor/inc/Function.asp" -->
<%
	On Error Resume Next
	'Call getquerystringparameter
	Dim pmdidx : pmdidx = request("mdidx")
	Dim pside : pside = request("side")
	Dim plastdate : plastdate = request("lastdate")
	dim cyear : cyear = request("cyear")
	dim cmonth : cmonth = request("cmonth")

	Dim sql_ :  sql_ = "select seq, cyear, cmonth, desc_01, desc_02, desc_03, desc_04 from wb_contact_photo where mdidx = " & pmdidx & " and side = '" & pside & "' order by seq desc "
	Dim cmd_ : Set cmd_ = server.CreateObject("adodb.command")
	cmd_.activeconnection = application("connectionstring")
	cmd_.commandText = sql_
	cmd_.commandType = adCmdText
	Dim rs_ : Set rs_ = cmd_.execute
	Set cmd_ = Nothing
	Dim rs_seq
	Dim rs_cyear
	Dim rs_cmonth
	Dim rs_desc_01
	Dim rs_desc_02
	Dim rs_desc_03
	Dim rs_desc_04
	If Not rs_.eof Then
		Set rs_seq = rs_(0)
		Set rs_cyear = rs_(1)
		Set rs_cmonth = rs_(2)
		Set rs_desc_01 = rs_(3)
		Set rs_desc_02 = rs_(4)
		Set rs_desc_03 = rs_(5)
		Set rs_desc_04 = rs_(6)
	End If
%>
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
			<link href="/MP/outdoor/style.css" rel="stylesheet" type="text/css">
			<title>�Ƣ� SK M&C | Media Management System �Ƣ�  </title>
			<script type='text/javascript' src='/js/ajax.js'></script>
			<script type='text/javascript' src='/js/script.js'></script>
			<script type="text/javascript" src="/js/calendar.js"></script>
			<script type="text/javascript">
			<!--

				var orgElement ;

				function get() {
					var params = "";
					sendRequest(url, params, _get, "GET");
				}

				function _get() {
					if (xmlreq.readyState == 4) {
						if (xmlreq.status == 200) {
						}
					}
				}

				function submitchange() {
				// ���� �ű� ���� & ���� ����
					var blank = true;
					var frm = document.forms[0];
					var colElement = document.getElementsByTagName("input");
					for (var i = 0 ; i < colElement.length ; i++) {
						if (colElement[i].getAttribute("type") == "file") {
							if (colElement[i].value) blank = false;
						}
					}
					if (document.getElementById("crud").value == "d") {blank = false;}

					if (blank) {alert("������ �̹����� �ϳ� �̻��� �����ϼ���"); return false;}
					frm.action = "/MP/outdoor/process/db_photo.asp";
					frm.method = "post";
					frm.submit();
				}

				function hiddenLayer() {
				// ���̾� ���߱�
					document.getElementById("photoLayer").style.display = "none";
				}

				function reloading(p) {
					// �̹��� �ʱ�ȭ
					var elem = document.getElementById(p.getAttribute('className'));
					elem.select();
					document.selection.clear();
					document.getElementById(elem.getAttribute("className")).setAttribute("src","/images/noimage.gif");
				}

				function prephoto(p) {
				// ���� ���ε� �̹��� �̸�����
					var img = document.getElementById(p.getAttribute('className'));
					var src = p.value;
					if (src=="") img.setAttribute("src", "/images/noimage.gif");
					else img.setAttribute("src", src);
				}

				function modPhoto() {
				// ���� ���� ��ư ����
					var file = document.getElementById("file");
					file.innerHTML = "<input type='file' id='file05' name='file05' class='showimg' style='width:406px;margin-left:150px;' onchange='prephoto(this);'>";
					var mng = document.getElementById("mng");
					mng.innerHTML = "<a href='#' onclick='submitchange(); return false;'><strong>����</strong></a> | <a href='#' onclick='canclePhoto(); return false;'><strong>���</strong></a>";
				}

				function canclePhoto() {
					// ���� ���� ��� (���̾� ���߱�)
					var img = document.getElementById("showimg");
					var file = document.getElementById("file");
					var mng = document.getElementById("mng");
					img.setAttribute("src", orgElement.src);
					file.innerHTML = "";
					mng.innerHTML = "<a href='#' onclick='modPhoto();return false;'><strong>����</strong></a> | <a href='#'><strong>����</strong></a>";
				}

				function deletePhoto() {
					if (confirm("������ ������ �����Ͻðڽ��ϱ�?")) {
						document.getElementById("crud").value = 'd';
						submitchange();
					}
				}

				function showLayer(seq, col) {
					// ���� ���� ���̾� ����
					var photoLayer = document.getElementById("photoLayer");
					photoLayer.style.display = "block";
					orgElement = event.srcElement;
					photoLayer.innerHTML = "<span id='closed' style='width:400px;height:20px;margin-left:150px;margin-top:70px;text-align:right;'><a href='#' onclick='hiddenLayer(); return false;'><strong>�ݱ�</strong></a></span>";
					photoLayer.innerHTML += "<img src="+orgElement.src+" width='400' height='320' align='absmiddle' style='border: 3px solid #FFFFFF; margin-left:150px; padding-top:80px;' id='showimg'>";
					photoLayer.innerHTML += "<div id='file' style='height:30px;'>&nbsp;</div>";
					photoLayer.innerHTML +="<span id='mng' style='margin-left:500px;'><a href='#' onclick='modPhoto(); return false;'><strong>����</strong></a> | <a href='#' onclick='deletePhoto(); return false;'><strong>����</strong></a> </span>";
					document.getElementById("seq").value = seq;
					document.getElementById("col").value= col;
					document.getElementById("crud").value = 'u';
				}

				function newLayer() {
					// �ű� ���� ��� ���̾� ����
					var photoLayer = document.getElementById("photoLayer");
					photoLayer.style.display = "block";
					photoLayer.innerHTML = "<div id='year' style='margin-left:73px;height:30px;'></div>";
					photoLayer.innerHTML += "<img src='/images/noimage.gif' id='img01' class='file01'  style='width:250px; height:190px; border: 2px solid #FFFFFF;' hspace='73' alt='ù��° ����'> <img src='/images/noimage.gif' id='img02' class='file02' style='width:250px; height:190px; border: 2px solid #FFFFFF;'  alt='�ι�° ����'> <input type='file' id='file01' name='file01' class='img01' style='width:215px;margin-left:73px;' onchange='prephoto(this);'  alt='ù��° ����ã��'/>   <a href='#' class='file01' onclick='reloading(this);' ><img src='/images/reset.jpg' width='35' height='19' align='absmiddle'  alt='ù��° ���¹�ư'></a>   <input type='file' id='file02' name='file02' class='img02'  style='width:215px;margin-left:73px;' onchange='prephoto(this);' alt='�ι�° ����ã��'/>   <a href='#' class='file02' onclick='reloading(this);' ><img src='/images/reset.jpg' width='35' height='19' align='absmiddle'  alt='�ι�° ���¹�ư'></a> <img src='/images/noimage.gif' id='img03' class='file03' style='width:250px; height:190px; border: 2px solid #FFFFFF; margin-top:10px;' hspace='73' alt='����° ����'> <img src='/images/noimage.gif' id='img04' class='file04' style='width:250px; height:190px; border: 2px solid #FFFFFF;' alt='�׹�° ����'> <input type='file' id='file03' name='file03' class='img03' style='width:215px;margin-left:73px;'  onchange='prephoto(this);'  alt='����° ���� ã��'/> <a href='#' class='file03' onclick='reloading(this);' ><img src='/images/reset.jpg' width='35' height='19' align='absmiddle'  alt='����° ���¹�ư'></a> <input type='file' id='file04' name='file04' class='img04' style='width:215px;margin-left:73px;' onchange='prephoto(this);'  alt='�׹�° ����ã��'/> <a href='#' class='file04' onclick='reloading(this);' ><img src='/images/reset.jpg' width='35' height='19' align='absmiddle' alt='�׹�° ���¹�ư'></a><p /><div id='mng' style='margin-left:605px;'> <a href='#' onclick='submitchange(); return false;'><strong>����</strong></a> | <a href='#' onclick='hiddenLayer(); return false;'><strong> �ݱ� </strong></a> </div>";
					document.getElementById("year").innerHTML = "<%Call getyear(cyear)%> <%Call getmonth(cmonth)%>";
					document.getElementById("crud").value = 'c';
				}

				window.onload = function () {
					self.focus();
				}

				window.onunload = function () {
//					try {
//						window.opener.getcontactphoto();
//					} catch(e) {
//						window.close();
//					}
				}
			//-->
		</script>
	</head>

<body>
<form onsubmit="return submitchange();"  enctype="multipart/form-data" >
<input type="hidden" id="mdidx" name="mdidx" value="<%=pmdidx%>" />
<input type="hidden" id="side" name="side" value="<%=pside%>" />
<input type='hidden' name='lastdate' id='lastdate' value="<%=plastdate%>"/>
<input type='hidden' name='seq' id='seq'>
<input type='hidden' name='col' id='col'>
<input type='hidden' name='crud' id='crud'>
<table width="720" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="/images/pop_bg.gif" height="50" align="left" valign="top" style="padding-left:18px; padding-top:27px;color:#FFFFFF; font-size:16px;font-weight:bolder;"><img src="/images/pop_title_dot.gif" width="5" height="14" align="absmiddle"> ���� ���� ���� </td>
    <td background="/images/pop_bg.gif" align="right"><img src="/images/pop_logo.gif" width="121" height="51"></td>
  </tr>
</table>
<table width="720" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="22"><img src="/images/pop_left_body_top.gif" width="22" height="16"></td>
    <td background="/images/pop_center_top_bg.gif">&nbsp;</td>
    <td width="23"><img src="/images/pop_right_body_top.gif" width="23" height="16"></td>
  </tr>
  <tr>
    <td background="/images/pop_left_middle_bg.gif">&nbsp;</td>
    <td height="458" valign='top'>
		<div style='overflow-y:scroll;width:675px;height:428px;'>
<!--  -->
		<table border="0" cellpadding="0" cellspacing="0">
			<tr height='20'>
				<td width='300' valign='top'><a href="#" onclick="newLayer(); return false;"><img src='/images/m_add.gif' width='14' height='15' alt="���� �߰�"></a> �����߰�</td>
				<td width='375' align='right' valign='top'> * ����, ������ ������ �����ϼ���.</td>
			</tr>
		</table>
		<table border="0" cellpadding="0" cellspacing="0">
			<tr>
				<th class='normal' width="50">����</th>
				<th class='normal' width="50">��</th>
				<th class='normal' width="135">��������</th>
				<th class='normal' width="135">��������</th>
				<th class='normal' width="135">��������</th>
				<th class='normal' width="135">��������</th>
			</tr>
			<%
				Do Until rs_.eof
				If IsNull(rs_desc_01) Or rs_desc_01 = "" Then rs_desc_01 = ""
				If IsNull(rs_desc_02) Or rs_desc_02 = "" Then rs_desc_02 = ""
				If IsNull(rs_desc_03) Or rs_desc_03 = "" Then rs_desc_03 = ""
				If IsNull(rs_desc_04) Or rs_desc_04 = "" Then rs_desc_04 = ""
			%>
			<tr>
				<td class='normal'><%=rs_cyear%></td>
				<td class='normal'><%=rs_cmonth%></td>
				<td class='normal'><a href="#" onclick="showLayer(<%=rs_seq%>, 'desc_01');"><%=getimage(rs_desc_01, 115)%></a></td>
				<td class='normal'><a href="#" onclick="showLayer(<%=rs_seq%>, 'desc_02');"><%=getimage(rs_desc_02, 115)%></a></td>
				<td class='normal'><a href="#" onclick="showLayer(<%=rs_seq%>, 'desc_03');"><%=getimage(rs_desc_03, 115)%></a></td>
				<td class='normal'><a href="#" onclick="showLayer(<%=rs_seq%>, 'desc_04');"><%=getimage(rs_desc_04, 115)%></a></td>
			</tr>
			<%
					rs_.movenext
				Loop
			%>
		</table>
		</div>
<!--  -->
	</td>
    <td background="/images/pop_right_middle_bg.gif"></td>
  </tr>
  <tr>
    <td><img src="/images/pop_left_bottom_bg.gif" width="22" height="25"></td>
    <td background="/images/pop_center_bottom_bg.gif">&nbsp;</td>
    <td><img src="/images/pop_right_bottom_bg.gif" width="23" height="25"></td>
  </tr>
</table>

<div id='buttonLayer' style='left:612px; top:505px;width:120px;height:18px;position:absolute; z-index:9;' ><a href="#" onclick='window.close();'><img src='/images/btn_close.gif' width='57' height='18'></a></a> </div>
<div id='photoLayer' style='left:0px; top:0px;width:720px;height:550px;position:absolute;z-index:10;display=none;background-color:#CCCCCC;filter:alpha(opacity=100);padding-top:10px;' ></div>
</form>
</body>
</html>