<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<html>
<head>
<title>�Ƣ� SK M&C | Media Management System �Ƣ� </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<script language="javascript" src="/js/ajax.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!--

//==================================�����������κ�=============================================
		function getdata(page) {
//			 ������ �޺� �ڽ� ��������
			page = (!!!page) ? "account.asp" : page ;
			var findID = 	 (document.getElementById("txtfindID")) ? document.getElementById("txtfindID").value : "" ;
			var findNAME = 	 (document.getElementById("txtfindNAME")) ? document.getElementById("txtfindNAME").value : "" ;
			var cmbFLAG = 	 (document.getElementById("cmbFLAG")) ? document.getElementById("cmbFLAG").value : "" ;

			var params = "findID="+findID + "&findNAME="+escape(findNAME) + "&cmbFLAG="+cmbFLAG;
			sendRequest(page, params, _getdata, "GET");
			document.getElementById("process").innerHTML = "<img src='http://ymail.net/skin/default/images/loding_clock.gif' style='margin-top:50px;'>" ;
		}

		function _getdata() {
			var process = document.getElementById("process");
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
						process.innerHTML = xmlreq.responseText ;
				} else {
						process.innerHTML = "<img src='http://ymail.net/skin/default/images/loding_clock.gif' style='margin-top:50px;'>" ;
				}
			}
		}

		function left_getdata(page) {
//			 ������ �޺� �ڽ� ��������
			page = (!!!page) ? "account.asp" : page ;
//			var cyear = 	 (document.getElementById("cyear")) ? document.getElementById("cyear").value : "" ;
//			var cmonth = 	 (document.getElementById("cmonth")) ? document.getElementById("cmonth").value : "" ;
//			var cyear2 = 	 (document.getElementById("cyear2")) ? document.getElementById("cyear2").value : "" ;
//			var cmonth2 = (document.getElementById("cmonth2")) ? document.getElementById("cmonth2").value : "" ;
//			var custcode2 = (document.getElementById("custcode2")) ? document.getElementById("custcode2").value : "" ;
			var params = "";
			sendRequest(page, params, _left_getdata, "GET");
			document.getElementById("process").innerHTML = "<img src='http://ymail.net/skin/default/images/loding_clock.gif' style='margin-top:50px;'>" ;

			var tcustcode = document.getElementById("tcustcode");
			tcustcode.innerText = "" ;
			var tflag = document.getElementById("tflag");
			tflag.innerText = "" ;

		}

		function _left_getdata() {
			var process = document.getElementById("process");
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
						process.innerHTML = xmlreq.responseText ;
				} else {
						process.innerHTML = "<img src='http://ymail.net/skin/default/images/loding_clock.gif' style='margin-top:50px;'>" ;
				}
			}
		}


		function left_menu_getdata(page,custcode,FLAG) {
//			 ������ �޺� �ڽ� ��������

			page = (!!!page) ? "account.asp" : page ;
//			var cyear = 	 (document.getElementById("cyear")) ? document.getElementById("cyear").value : "" ;
//			var cmonth = 	 (document.getElementById("cmonth")) ? document.getElementById("cmonth").value : "" ;
//			var cyear2 = 	 (document.getElementById("cyear2")) ? document.getElementById("cyear2").value : "" ;
//			var cmonth2 = (document.getElementById("cmonth2")) ? document.getElementById("cmonth2").value : "" ;

			var params = "custcode=" + custcode + "&FLAG=" + FLAG ;
			sendRequest(page, params, _left_menu_getdata, "GET");

			document.getElementById("process").innerHTML = "<img src='http://ymail.net/skin/default/images/loding_clock.gif' style='margin-top:50px;'>" ;

			var tcustcode = document.getElementById("tcustcode");
			tcustcode.innerText = custcode ;

			var tflag = document.getElementById("tflag");
			tflag.innerText = FLAG ;


		}

		function _left_menu_getdata() {
			var process = document.getElementById("process");
			if (xmlreq.readyState == 4) {
				if (xmlreq.status == 200) {
						process.innerHTML = xmlreq.responseText ;
				} else {
						process.innerHTML = "<img src='http://ymail.net/skin/default/images/loding_clock.gif' style='margin-top:50px;'>" ;
				}
			}

		}


//==================================������������������κ�=============================================
//==================================account.asp=============================================

	//���� ����
	function getemployee(crud) {
		
		var userid = window.frames[0].framefrm.txtuserid.value

		switch (crud) {
			case "v":
					var url = "pop_account_view.asp?userid=" + userid;
					var name = "pop_account_view";
					var opt = "width=540, height=350, resziable=no, scrollbars = no, status=yes, top=100, left=100";
					window.open(url, name, opt);
				break;					
			case "c":
					var url = "pop_account_reg.asp" ;
					var name = "pop_reg";
					var opt = "width=540, height=300, resziable=no, scrollbars = no, status=yes, top=100, left=100";
					window.open(url, name, opt);
				break;
			case "d":
					if (userid == "")
					{
						alert("������ ���̵� �����ϼ���."); 
						return false;
					}

					if (confirm("������ ������ �����Ͻðڽ��ϱ�? \n�ش� ������ ������, ���� �����˴ϴ�.")) {
						document.frames['frmdeleteproc'].location.href="account_delete_user_proc.asp?userid="+userid;
						document.frames['frmuser'].location.href="account_fuserid.asp?userid="+userid;
					}
				break ;
		}
	}
	
	//������ ����
	function getemployee_cust(crud) {
		
		var userid = window.frames[0].framefrm.txtuserid.value
		var userclass = window.frames[0].framefrm.txtclass.value
		var custcode = window.frames[1].framefrm_cust.txtcustcode.value
		
		
	
		if (userid == "") {
			alert("���������� �ʼ��Դϴ�.");
			return false;
		} 

		switch (crud) {				
			case "c":
					if (userclass =="A")
					{
						alert("�����ڴ� �����ָ� �߰��Ҽ� �����ϴ�."); 
						return false;
					}

					var url = "pop_custcode.asp?strUserid="+userid;
					var name = "pop_custcode";
					var opt = "width=540, height=480, resizable=no, scrollbars=yes, top=100, left=660";
					window.open(url, name, opt);
			
				break;
			case "d":
					if (userid == "")
					{
						alert("������ ���̵� �����ϼ���."); 
						return false;
					}
					if (custcode == "")
					{	
						alert("������ �����ָ� �����ϼ���."); 
						return false;
					}


					if (confirm("������ �����ָ� �����Ͻðڽ��ϱ�? \n�ش� ������ �����ֿ� ���� �����˴ϴ�.")) {
						document.frames['frmdeleteproc'].location.href="account_delete_cust_proc.asp?userid="+userid+"&custcode="+custcode;
						document.frames['frmcust'].location.href="account_fcust.asp?strUserid="+userid;
						document.frames['frmtim'].location.href="account_ftim.asp?strUserid="+userid;
					}
				break ;
		}
	}


	//�� ����
	function getemployee_tim(crud) {
		
		var userid = window.frames[0].framefrm.txtuserid.value
		var custcode = window.frames[1].framefrm_cust.txtcustcode.value
		var timcode = window.frames[2].framefrm_tim.txttimcode.value

		
		if (userid == "") {
			alert("���������� �ʼ��Դϴ�.");
			return false;
		} 
		
		if (custcode == "") {
			alert("�����ּ����� �ʼ��Դϴ�.");
			return false;
		} 

		switch (crud) {				
			case "c":
		
					var url = "pop_timcode.asp?strUserid="+userid + "&strCustcode="+ custcode;
					var name = "pop_timcode";
					var opt = "width=540, height=480, resizable=no, scrollbars=yes, top=100, left=660";
					window.open(url, name, opt);
			
				break;
			case "d":
					if (userid == "")
					{
						alert("������ ���̵� �����ϼ���."); 
						return false;
					}
					if (custcode == "")
					{	
						alert("������ �����ָ� �����ϼ���."); 
						return false;
					}
				
					if (timcode == "" )
					{	
						alert("������ ���� �����ϼ���."); 
						return false;
					}
				

					if (confirm("������ ���� �����Ͻðڽ��ϱ�?")) {
						document.frames['frmdeleteproc'].location.href="account_delete_tim_proc.asp?userid="+userid+"&custcode="+custcode+"&timcode="+timcode;
						document.frames['frmtim'].location.href="account_ftim.asp?strUserid="+userid+"&strCustcode="+custcode;
					}
				break ;
		}
	}

	function user_Class_src(userid){
		frmcust.location.href="account_fcust.asp?strUserid="+userid;
		frmtim.location.href="account_ftim.asp?strUserid="+userid;
		//frmtim.location.reload();
	}


	function cust_Class_src(userid,custcode){
		//frmcust.location.href="account_fcust.asp?strUserid="+userid;
		frmtim.location.href="account_ftim.asp?strUserid="+userid+"&strCustcode="+custcode;
		//frmtim.location.reload();
	}



//
//	function checkForView(uid,c_class) {
//		var url = "pop_account_view.asp?userid=" + uid + "&c_class=" + c_class;
//		var name = "pop_account_view";
//		var opt = "width=540, height=320, resziable=no, scrollbars = no, status=yes, top=100, left=100";
//		window.open(url, name, opt);
//	}
//	function pop_reg() {
//		var p = document.getElementById("btnReg") ;
//		var custcode = document.forms[0].tcustcode.value.replace("null","") ;
//
//		if (p.getAttribute("class") == "account" || p.getAttribute("class") == null) {
//			var url = "pop_account_reg.asp?tcustcode="+custcode ;
//			var name = "pop_reg";
//			var opt = "width=540, height=380, resziable=no, scrollbars = no, status=yes, top=100, left=100";
//		} else {
//			var url = "pop_menu_reg.asp?tcustcode="+custcode;
//			var name ="pop_menu_reg" ;
//			var opt = "width=540, height=215, resziable=no, scrollbars = no, status=yes, top=100, left=100";
//		}
//		window.open(url, name, opt);
//	}

	function checkForSearch(str) {
		var frm = document.forms[0];
		if (str !="") {
			if (str.indexOf("--") != -1) {
				alert("����� �� ���� ���ڸ� �Է��ϼ̽��ϴ�.");
				frm.txtsearchstring.value = "";
				frm.txtsearchstring.focus();
				return false;
			}
		}
		frm.action = frm.actionurl.value;
		frm.method = "post";
		frm.submit();
	}

	document.onkeydown = function() {
		if (event.keyCode == "13") return false;
	}




//==================================�������account.asp=============================================




//==================================menu.asp=============================================
   function pop_menu_reg() {
		var custcode = document.forms[0].tcustcode.value.replace("null","") ;
		var flag = document.forms[0].tflag.value.replace("null","") ;
		var url = "pop_menu_reg.asp?custcode="+custcode +"&tflag=" +flag ;
		var name ="pop_menu_reg" ;
		var opt = "width=540, height=205, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function go_submenu_reg(idx) {
		var flag = document.forms[0].tflag.value.replace("null","") ;

		var url = "pop_submenu_reg.asp?midx="+idx  +"&tflag=" +flag ;
		var name ="pop_submenu_reg" ;
		var opt = "width=540, height=205, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}

	function go_menu_view(idx) {

		var url = "pop_menu_view.asp?midx=" + idx;
		var name ="pop_menu_view" ;
		var opt = "width=540, height=205, resziable=no, scrollbars = no, status=yes, top=100, left=100";
		window.open(url, name, opt);
	}
//==================================�������menu.asp=============================================


window.onload = getdata;
-->
</SCRIPT>

</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"  >
<form target="scriptFrame">
<!--#include virtual="/cust/top.asp" -->

  <input type="hidden" name="tcustcode">
  <input type="hidden" name="tflag">
  <table width="1240" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF" id="Table_01">
    <tr>
      <td valign="top"><!--#include virtual="/cust/left_admin_menu.asp"--></td>
      <td align="left" valign="top"><img src="/images/middle_navigater_admin.gif" width="1030" height="65" alt="">
	  <div id="process" style="text-align:center;"></div>
	  </td>
    </tr>
	<tr>
	<td colspan="2"><!--#include virtual="/bottom.asp" --></td>
	</tr>
  </table>
</form>
</body>
</html>

