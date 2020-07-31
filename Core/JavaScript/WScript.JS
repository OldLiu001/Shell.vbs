//模拟WScript对象。
//部分摘自CrLf兄的WSH.js，在此表示感谢。
window.WScript = window.WSH = {
    'Application': oMSHTA.applicationName,
    'CreateObject': function (ProgID, Prefix) {
        if (Prefix) {
            return new ActiveXObject(ProgID, Prefix);
        } else {
            return new ActiveXObject(ProgID);
        }
    },
    'BuildVersion': navigator.appVersion,
    'Echo': function () {
        var strText;
        if (arguments.length) {
            strText = arguments[0];
            for (var i = 1; i < arguments.length; i++) {
                strText += '    ' + arguments[i];
            }
            TextOutPut(strText);
        }
        return;
    },
    'Name': 'HTML Application',
    'FullName': 'mshta.exe',
    'ScriptName': oMSHTA.commandLine.replace(/^(")?[^"]+[\\\/]([^\\\/\?]+?)(?:[\?#][^\\\/\?#]*)?\1.*$/, "$2"),
    'ScriptFullName': location.pathname,
    'Timeout': 0,
    'Interactive': 1,
    'Quit': window.close
}