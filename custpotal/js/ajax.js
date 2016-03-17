var xmlreq = null ;

function _getXMLHttpRequest() { 
	if (window.ActiveXObject) {
		try { return new ActiveXObject("Msxml2.XMLHTTP");}
		catch (e) {
			try { return new ActveXObject("Microsoft.XMLHTTP");}
			catch (e1) { return null;}
		}
	} else if (window.XMLHttpRequest) {
		new XMLHttpRequest();
	} else {
		return null;
	}
}

function sendRequest(url, params, callback, method) {
//	alert("sendRequest");
	xmlreq = _getXMLHttpRequest();
	var httpMethod = method ? method : 'GET';
	if (httpMethod != 'GET' && httpMethod != "POST") httpMethod = 'GET';
	var httpParams = (params == null || params == '') ? null : params ;
	var httpUrl = url ;
	if (httpMethod == 'GET' && httpParams != null) httpUrl = httpUrl + "?" + httpParams;

	xmlreq.open(httpMethod, httpUrl, true) ;
	xmlreq.setRequestHeader('Content-Type', 'application/x-www-form-urlendcoded');
	xmlreq.onreadystatechange = callback;
	xmlreq.send(httpMethod == 'POST' ? httpParams : null);
}

function log(msg) {
	var debugConsole = document.getElementById("debugConsole");
	debugConsole.innerHTML += msg + "<br>";
}

function _sendRequest(url, params, callback, method) {
	xmlreq = _getXMLHttpRequest();
	var httpMethod = method ? method : 'GET';
	if (httpMethod != 'GET' && httpMethod != "POST") httpMethod = 'GET';
	var httpParams = (params == null || params == '') ? null : params ;
	var httpUrl = url ;
	if (httpMethod == 'GET' && httpParams != null) httpUrl = httpUrl + "?" + httpParams;

	xmlreq.open(httpMethod, httpUrl, false) ;
	xmlreq.setRequestHeader('Content-Type', 'application/x-www-form-urlendcoded');
	xmlreq.onreadystatechange = callback;
	xmlreq.send(httpMethod == 'POST' ? httpParams : null);
}


function _callback() {
	if (xmlreq.readyState == 4) {
		if (xmlreq.status == 200) {
				/* code */
		}
	}
}
