
// F11 ���� ����
document.onkeydown = function() {
	if (event.keyCode == 122) {
		event.keyCode = 505;
	}

	if (event.keyCode == 505) { 
		return false;
	}
}

