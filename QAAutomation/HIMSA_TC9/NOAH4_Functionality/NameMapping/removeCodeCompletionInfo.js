var objArgs = WScript.Arguments;

if (objArgs.length  < 1) {
   WScript.Echo("Please specify the NameMapping.tcNM file as a parameter");
}
else {
   WScript.Echo("Processing file " + objArgs(0));
   removeCodeCompletionInfo(objArgs(0));
}

function removeCodeCompletionInfo(fileName)
{
   WScript.Echo("Function start")
   var xml = new ActiveXObject("MSXML2.DOMDocument.4.0");
   xml.load(fileName);

   var ccInfoNodes = xml.selectNodes("//Node[@name=\"typeinfo\"]");
   for (var i = 0; i < ccInfoNodes.length; i++) {
     var parent = ccInfoNodes.item(i).parentNode;
     parent.removeChild(ccInfoNodes.item(i));
   }
   xml.save(fileName);
   WScript.Echo("Completed");
}


