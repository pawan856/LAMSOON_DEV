function addInterfaceData(formName, interfaceSpan, aName, aValue)
{	
	var targetInput = eval("document." + formName + "." + aName);
	if(targetInput ==  null)
	{
		var interfaceElement = document.createElement("input");
		interfaceElement.type = "hidden";
		interfaceElement.name = ""+aName;
		interfaceElement.value = ""+aValue;
		interfaceSpan.appendChild(interfaceElement);
	}
	else
	{	targetInput.value = aValue;
	}
}

function clearInterfaceData(interfaceSpan)
{	interfaceSpan.innerHTML = "";
}

/*
Newer version of addInterfaceData. 
Created by Clive on 21 Jan 2003
*/
function setInterfaceData(aName, aValue)
{
	var targetInput = eval("document.forms[0]." + aName);
	if(targetInput ==  null)
	{
		//var interfaceElement = document.createElement("<INPUT TYPE='hidden' NAME='" + aName + "' VALUE='" + aValue + "'>");
	    //var interfaceElement = document.createElement("<INPUT TYPE='hidden' NAME='" + aName + "'>");
	    try {
	        interfaceElement = document.createElement("<INPUT TYPE='hidden' NAME='" + aName + "'>");
	    }
	    catch (e) {
	        interfaceElement = document.createElement("INPUT");
	        interfaceElement.setAttribute("TYPE", "hidden");
	        interfaceElement.setAttribute("NAME", aName);
	    }
		document.forms[0].appendChild(interfaceElement);
		eval("document.forms[0]." + aName).value = aValue;
	}
	else
	{
		targetInput.value = aValue;
	}
}

function setInterfaceDataToForm(aForm, aName, aValue)
{
	var targetInput = eval("aForm." + aName);

	if (targetInput == null)
	{
		//var interfaceElement = document.createElement("<INPUT TYPE='hidden' NAME='" + aName + "' VALUE='" + aValue + "'>");
	    var interfaceElement;
	    try {
	        interfaceElement = document.createElement("<INPUT TYPE='hidden' NAME='" + aName + "'>");
	    }
	    catch (e) {
	        interfaceElement = document.createElement("INPUT");
	        interfaceElement.setAttribute("TYPE", "hidden");
	        interfaceElement.setAttribute("NAME", aName);
	    }
		
		aForm.appendChild(interfaceElement);
		eval("aForm." + aName).value = aValue;
	}
	else
	{
		targetInput.value = aValue;
	}
}

function removeAllElementFromForm(aForm)
{
	var i;

	for (i = aForm.childNodes.length - 1; i >=0; i--)
		aForm.removeChild(aForm.childNodes[i]);
}