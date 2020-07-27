//模拟WScript对象。
var WScript = {
	"CreateObject" : function(ProgID, Prefix) {
		if(Prefix) {
			return new ActiveXObject(ProgID, Prefix);
		} else {
			return new ActiveXObject(ProgID);
		}
	},
	"BuildVersion" : 16384,
	"Echo" : function(){
		var strText;
		if(arguments.length){
			strText = arguments[0];
			for(var i = 1; i < arguments.length; i++) {
				strText+='	' + arguments[i];
			}
			TextOutPut(strText);
		}
		return;
	}
}
var WSH = WScript;