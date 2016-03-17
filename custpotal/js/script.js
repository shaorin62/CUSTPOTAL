function checknumber(elem) {
	if (isNaN(elem.value)) {
		alert("숫자만 입력하세요");
		elem.value = "";
		elem.focus();
		return false;
	}
}

function comma(elem) {
	var str = elem.value.replace(/[^\d]/g, "");
	var retval = Number(String(str)).toLocaleString().slice(0,-3);
	if (retval < 0) retval = 0 ;
    var mi =  elem.value.indexOf("-");
	if (mi ==  1 || mi ==  0 )
	{
		retval = '-' + retval
	}
	elem.value = retval ;
	var rng = elem.createTextRange();
	rng.collapse();
	rng.moveStart('character', 100); // input type='text style='text-align:right'
	rng.select();
}

function checktextlength(elem, max) {
	var len = elem.value.length;
		if (max < len) {
			elem.value = elem.value.substring(0, max);
		}
}

function getpopup(url, name, opt, outerWidth, outerHeight) {	
	var left = screen.width / 2 - outerWidth / 2;
	var top = screen.height / 2 - outerHeight / 2;
	var opt = opt + ",  left="+left+",top="+top ;

	if (name == "") {name = "popup"; }
	window.open(url, name, opt)
}

//	function insertRow(id, ary) {
//		var tableElement = document.getElementById(id);
//		var rowElement = tableElement.insertRow(0);
//		
//		for (var i = 0 ; i < ayrvalue.length ; i++) {
//			var cellElement = rowElement.insertCell(-1);
//			cellElement.appendChild(document.createTextNode(ary));
//		}
//	}
//
//	function deleteRow(id, idx) {
//		var tableElement = document.getElementById(id);
//		var rows = tableElement.rows.length;
//		var frm = document.forms[0];
//			for (var i =0 ; i < frm.chk.length ; i++) {
//				tableElement.deleteRow(i);
//			}
//		if (frm.chk.length <= 1) {
//			tableELement.deleteRow(-1);
//		} else {
//
//		}
//
//	}
