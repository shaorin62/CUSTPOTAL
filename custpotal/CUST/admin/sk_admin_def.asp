<!--#include virtual="/inc/getdbcon.asp" -->
<!--#include virtual="/inc/func.asp" -->

<html>
<head>
<title>▒▒ SK M&C | Media Management System ▒▒ </title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="/style.css" rel="stylesheet" type="text/css">
<script language="javascript" src="/js/ajax.js"></script>
<SCRIPT LANGUAGE="JavaScript">
<!--

//==================================공통페이지부분=============================================
		function getdata(page) {
//			 광고주 콤보 박스 가져오기
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
//			 광고주 콤보 박스 가져오기
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
//			 광고주 콤보 박스 가져오기

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


//==================================여기까지공통페이지부분=============================================
//==================================account.asp=============================================

	//계정 정보
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
						alert("삭제할 아이디를 선택하세요."); 
						return false;
					}

					if (confirm("선택한 계정을 삭제하시겠습니까? \n해당 계정과 광고주, 팀도 삭제됩니다.")) {
						document.frames['frmdeleteproc'].location.href="account_delete_user_proc.asp?userid="+userid;
						document.frames['frmuser'].location.href="account_fuserid.asp?userid="+userid;
					}
				break ;
		}
	}
	
	//광고주 정보
	function getemployee_cust(crud) {
		
		var userid = window.frames[0].framefrm.txtuserid.value
		var userclass = window.frames[0].framefrm.txtclass.value
		var custcode = window.frames[1].framefrm_cust.txtcustcode.value
		
		
	
		if (userid == "") {
			alert("계정선택은 필수입니다.");
			return false;
		} 

		switch (crud) {				
			case "c":
					if (userclass =="A")
					{
						alert("관리자는 광고주를 추가할수 없습니다."); 
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
						alert("삭제할 아이디를 선택하세요."); 
						return false;
					}
					if (custcode == "")
					{	
						alert("삭제할 광고주를 선택하세요."); 
						return false;
					}


					if (confirm("선택한 광고주를 삭제하시겠습니까? \n해당 계정의 광고주와 팀도 삭제됩니다.")) {
						document.frames['frmdeleteproc'].location.href="account_delete_cust_proc.asp?userid="+userid+"&custcode="+custcode;
						document.frames['frmcust'].location.href="account_fcust.asp?strUserid="+userid;
						document.frames['frmtim'].location.href="account_ftim.asp?strUserid="+userid;
					}
				break ;
		}
	}


	//팀 정보
	function getemployee_tim(crud) {
		
		var userid = window.frames[0].framefrm.txtuserid.value
		var custcode = window.frames[1].framefrm_cust.txtcustcode.value
		var timcode = window.frames[2].framefrm_tim.txttimcode.value

		
		if (userid == "") {
			alert("계정선택은 필수입니다.");
			return false;
		} 
		
		if (custcode == "") {
			alert("광고주선택은 필수입니다.");
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
						alert("삭제할 아이디를 선택하세요."); 
						return false;
					}
					if (custcode == "")
					{	
						alert("삭제할 광고주를 선택하세요."); 
						return false;
					}
				
					if (timcode == "" )
					{	
						alert("삭제할 팀을 선택하세요."); 
						return false;
					}
				

					if (confirm("선택한 팀을 삭제하시겠습니까?")) {
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
				alert("사용할 수 없는 문자를 입력하셨습니다.");
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




//==================================여기까지account.asp=============================================




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
//==================================여기까지menu.asp=============================================


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

