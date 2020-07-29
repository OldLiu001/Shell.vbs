function TextOutPut(strText) {
    var pNode = document.createElement("p");
    pNode.innerText = strText;
    document.getElementById("Console").appendChild(pNode);
}