	function checkForKey() {
		var str = document.forms[0].txtaccount.value;
		var rgex = /^(\w|\d)$/;
		if (!rgex.test(str.charAt(str.length-1)) ){
			document.forms[0].txtaccount.value = str.substring(0, str.length-1);
			return false;
		}
	}

	function checkForCustomer1(val) {
		var frm = document.forms[0];
		frm.txtdeptcode.value = "";
		frm.txtdeptname.value = "";
		frm.txtcustcode.value = "";
		frm.txtcustname.value = "";
		document.getElementById("employee").style.display = "none";
		switch(val.toString()) {
			case "1":
			case "2":
				window.open("employee_List.asp", "Authority", "width=500, height=500, resizable=no, scrollbars=yes, top=100, left=100");
				break;
			case "3":
				window.open("customer_List.asp", "Authority", "width=500, height=500, resizable=no, scrollbars=yes, top=100, left=100");
				break;
			case "4":
				window.open("dept_List.asp", "Authority", "width=500, height=500, resizable=no, scrollbars=yes,  top=100, left=100");
				break;
			case "5":
				window.open("media_List.asp", "Authority", "width=500, height=500, resizable=no, scrollbars=yes, top=100, left=100");
				break;
			case "6":
				window.open("outside_List.asp", "Authority", "width=500, height=500, resizable=no, scrollbars=yes, top=100, left=100");
				break;
		}
	}
	
	function checkForCustomer() {
		window.open("customer_List.asp", "Authority", "width=500, height=500, resizable=no, scrollbars=yes, top=100, left=100");
	}

	function checkForAuthority() {
		var val = 0;
		for (var i = 0 ; i < document.forms[0].rdoauthority.length ;i++) {
			if (document.forms[0].rdoauthority[i].checked) {
				val = i+1;
			}
		}
		if(val) {
			checkForCustomer(val.toString());
		} else {
			alert("접속권한을 선택하세요");
			document.forms[0].rdoauthority[0].focus();
			return false;
		}
	}