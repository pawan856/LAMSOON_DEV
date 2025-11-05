/*
function isEmail(fieldValue) 
{
   var newString = fieldValue.match(/\b(^(\S+@).+((\.com)|(\.net)|(\.info)|(\.edu)|(\.mil)|(\.gov)|(\.org)|(\..{2,2}))$)\b/gi);
   if (!newString) 
   	return false;
   else
   	return true;
}
*/

function isEmail(str) {
// are regular expressions supported
  var supported = 0 ;
  if (window.RegExp) {
	var tempStr = "a";
    var tempReg = new RegExp(tempStr);
    if (tempReg.test(tempStr)) supported = 1;
  }
  if (!supported)
    return (str.indexOf(".") >2) &&  (str.indexOf("@") > 0);
  var r1 = new RegExp("(@.*@)|(\\.\\.)|(@\\.)|(^\\.)");
  var r2 = new RegExp("^.+\\@(\\[?)[a-zA-Z0-9\\-\\.]+\\.([a-zA-Z]{2,4}|[0-9]{1,3})(\\]?)$");
  return (!r1.test(str) && r2.test(str));
}

function isEmpty(fieldValue)
{
	var regEx = /\S+/i;
	if(regEx.test(fieldValue))
		return false;
	else
		return true;
}

function isInteger(fieldValue)
{	var regEx=/^\d+$/;
	if (regEx.test(fieldValue))
		return true;
	else
		return false;
}

function isDecimal(fieldValue)
{
	return isFloat(fieldValue);
}

function isFloat(fieldValue)
{	
	var regEx=/(^\d+$)|(^\d+\.\d+$)|(^-\d+$)|(^-\d+\.\d+$)/;
	if (regEx.test(fieldValue))
		return true;
	else
		return false;
}

function chkDateFmt(fieldValue, aFormat)
{
	var regEx = /^\d?\d\/\d?\d\/\d\d\d\d$/;

    if (fieldValue.length == 0)
    {
        return true;
    }

	fieldValue = javaDateConverter('DD/MM/YY', 'MM/DD/YYYY', fieldValue)

	if (regEx.test(fieldValue))
	{
		var dmy;
		var dd;
		var mm;
		var yy;		

		dmy = fieldValue.split("/");
		if(dmy[0].charAt(0) == '0')
			dmy[0] = dmy[0].substring(1, 2);

		if(dmy[1].charAt(0) == '0')
			dmy[1] = dmy[1].substring(1, 2);
			
		if(parseInt(aFormat) == 1)
		{	// check for this format "MM/DD/YYYY"			
			dd = parseInt(dmy[1]);
			mm = parseInt(dmy[0])-1;
			yy = parseInt(dmy[2]);			
		}		
		else if(parseInt(aFormat) == 2)
		{	// check for this format "MM/DD/YYYY"
			dd = parseInt(dmy[0]);
			mm = parseInt(dmy[1])-1;			
			yy = parseInt(dmy[2]);				
		}

		var aDate = new Date(yy, mm, dd);

		if(aDate.getMonth() != mm)
			return false;
		else if(yy < 1980 || yy > 2100)
			return false;

		return true;
	}
	else
	{
		return false;		
	}
}

function isDate(fieldValue)
{	/* check for this format "MM/DD/YYYY"*/	

	fieldValue = javaDateConverter(jsDateFormat,"MM/DD/YYYY",fieldValue)

	var regEx=/^\d?\d\/\d?\d\/\d\d\d\d$/;
	if (regEx.test(fieldValue))
	{
		
		var dmy = fieldValue.split("/");
		if(dmy[0].charAt(0) == '0')
		{	dmy[0] = dmy[0].substring(1, 2);
		}

		if(dmy[1].charAt(0) == '0')
		{	dmy[1] = dmy[1].substring(1, 2);
		}
		
		var dd = parseInt(dmy[1]);
		var mm = parseInt(dmy[0])-1;
		var yy = parseInt(dmy[2]);
		
		var aDate = new Date(yy, mm, dd);		
		if(aDate.getMonth() != mm)
		{
			return false;
		}
		else if(yy < 1980 || yy > 2100)
		{	
			return false;
		}
		return true;
	}
	else
		return false;
}

function validateFieldNoMessage(fieldRef, criteria)		/*** USAGE NOT RECOMMENDED ***/
{	/* Requires: listUtil.js */
	/* This function assumes the criteria format "|CRITERIA_ONE||CRITERIA_TWO||CRITERIA_THREE|*/
	/*
		The checking will be "AND" based.
		Currently supported Criteria:
		DATE, NOT_EMPTY, INTEGER, FLOAT:9, MAX_LENGTH:123
	*/
	var tempArray;
	tempArray = listToArray(criteria);
	if(tempArray == null)		//No validation required.
		return true;
		
	for(var i=0; i<tempArray.length; i++)
	{	var tempIndex = tempArray[i].indexOf(":");
	
		if(tempIndex > 0)
		{	//Complex validation
			var tempCriteria = tempArray[i].substring(0, tempIndex);
			switch(tempCriteria)
			{
				case "MAX_LENGTH":
					var tempMaxLength = Number(tempArray[i].substring(tempIndex+1, tempArray[i].length));
					if(fieldRef.value.length > tempMaxLength)
						return false;
					break;
				
				case "FLOAT":					
					if(!isEmpty(fieldRef.value) && !isFloat(fieldRef.value))
						return false;
					else
					{	//check precision					
						var tempPrecision = tempArray[i].substring(tempIndex+1, tempArray[i].length);						
						var tempLength = fieldRef.value.length;
												
						var dotPos = fieldRef.value.indexOf(".");
						if(dotPos > 0) //dot found
						{	tempLength = dotPos;
						}
						if(Number(tempLength) > Number(tempPrecision))
							return false;
					}
					break;
					
				case "POSITIVE_INCLUDEZERO_FLOAT":					
					if((!isEmpty(fieldRef.value) && !isFloat(fieldRef.value)) || (!isEmpty(fieldRef.value) && isFloat(fieldRef.value) && (fieldRef.value < 0)))
						return false;
					else
					{	//check precision					
						var tempPrecision = tempArray[i].substring(tempIndex+1, tempArray[i].length);						
						var tempLength = fieldRef.value.length;
												
						var dotPos = fieldRef.value.indexOf(".");
						if(dotPos > 0) //dot found
						{	tempLength = dotPos;
						}
						if(Number(tempLength) > Number(tempPrecision))
							return false;
					}
					break;
			
				default:					
					alert("System error, validation type " + tempArray[i] + " not supported.\n Validation on this type is ignored." );
			}
		}
		else
		{	//Simple validation
			switch(tempArray[i])
			{
				case "DATE":
					if(!isEmpty(fieldRef.value) && !isDate(fieldRef.value))
						return false;
					break;
					
				case "NOT_EMPTY":
					if(isEmpty(fieldRef.value))
						return false;
					break;
				
				case "INTEGER":
					if(!isEmpty(fieldRef.value) && !isInteger(fieldRef.value))
						return false;
					break;
				case "FLOAT":
					if(!isEmpty(fieldRef.value) && !isFloat(fieldRef.value))
						return false;
					break;
				
				default:
					alert("System error, validation type " + tempArray[i] + " not supported.\n Validation on this type is ignored." );
			}
		}
	}
	return true;
}

function validateField(fieldRef, fieldDesc, criteria)
{	/* Requires: listUtil.js */
	/* This function assumes the criteria format "|CRITERIA_ONE||CRITERIA_TWO||CRITERIA_THREE|*/
	/*
		The checking will be "AND" based.
		Currently supported Criteria:
		DATE, NOT_EMPTY, INTEGER, FLOAT:9, MAX_LENGTH:123, POSITIVE_INCLUDEZERO_FLOAT:9
	*/	
	
	var tempArray;
	tempArray = listToArray(criteria);
	
	if(tempArray == null)		//No validation required.
		return true;

	for(var i=0; i<tempArray.length; i++)
	{	var tempIndex = tempArray[i].indexOf(":");
	
		if(tempIndex > 0)
		{	//Complex validation
			var tempCriteria = tempArray[i].substring(0, tempIndex);
			switch(tempCriteria)
			{
				case "MAX_LENGTH":
					var tempMaxLength = Number(tempArray[i].substring(tempIndex+1, tempArray[i].length));
					if(fieldRef.value.length > tempMaxLength)
					{
						alert("Field value too large, maximum length is " + tempMaxLength + "! " + fieldDesc);
						fieldRef.focus();
						return false;
					}					
					break;
				
				case "FLOAT":					
					if(!isEmpty(fieldRef.value) && !isFloat(fieldRef.value))
					{	
						alert("Invalid number! " + fieldDesc);
						fieldRef.focus();
						return false;
					}
					else
					{	//check precision					
						var tempPrecision = tempArray[i].substring(tempIndex+1, tempArray[i].length);						
						var tempLength = fieldRef.value.length;
												
						var dotPos = fieldRef.value.indexOf(".");
						if(dotPos > 0) //dot found
						{	tempLength = dotPos;
						}
						if(Number(tempLength) > Number(tempPrecision))
						{	alert("Field value too large, maximum precision is " + tempPrecision + "! " + fieldDesc);
							fieldRef.focus();
							return false;
						}
					}
					break;
				
				case "FLOAT_SIZE":					
					if(!isEmpty(fieldRef.value) && !isFloat(fieldRef.value))
					{	
						alert("Invalid number! " + fieldDesc);
						fieldRef.focus();
						return false;
					}
					else
					{	//check precision					
						var tempSize = tempArray[i].substring(tempIndex+1, tempArray[i].length);						
						var tempLength = fieldRef.value.length;
						
						var tempIndex2 = tempSize.indexOf(",");
						var tempPrecision = tempSize.substring(0, tempIndex2);
						var tempDecimal = tempSize.substring(tempIndex2+1, tempSize.length);
									
						var dotPos = fieldRef.value.indexOf(".");
						if(dotPos > 0) //dot found
						{
							tempLength = dotPos;
							var tempLength2 = fieldRef.value.length - dotPos - 1;
						}
						if(Number(tempLength) > Number(tempPrecision))
						{
							alert("Field value too large, maximum precision is " + tempPrecision + "! " + fieldDesc);
							fieldRef.focus();
							return false;
						}
						
						if(Number(tempLength2) > Number(tempDecimal))
						{
							alert("Decimal place too many, maximum decimal place is " + tempDecimal + "! " + fieldDesc);
							fieldRef.focus();
							return false;
						}
					}
					break;
					
				case "POSITIVE_INCLUDEZERO_FLOAT":	
						
					if((!isEmpty(fieldRef.value) && !isFloat(fieldRef.value)) || (!isEmpty(fieldRef.value) && isFloat(fieldRef.value) && (fieldRef.value < 0)))
					{	
						alert("Invalid number! " + fieldDesc);
						fieldRef.focus();
						return false;
					}
					else
					{	//check precision					
						var tempPrecision = tempArray[i].substring(tempIndex+1, tempArray[i].length);						
						var tempLength = fieldRef.value.length;
												
						var dotPos = fieldRef.value.indexOf(".");
						if(dotPos > 0) //dot found
						{	tempLength = dotPos;
						}
						if(Number(tempLength) > Number(tempPrecision))
						{	alert("Field value too large, maximum precision is " + tempPrecision + "! " + fieldDesc);
							fieldRef.focus();
							return false;
						}
					}
					break;
					
				default:					
					alert("System error, validation type " + tempArray[i] + " not supported.\n Validation on this type is ignored." );
			}
		}
		else
		{	//Simple validation
			switch(tempArray[i])
			{
				case "DATE":
					
					if(!isEmpty(fieldRef.value) && !isDate(fieldRef.value))
					{	
						alert("Invalid date! " + fieldDesc);
						fieldRef.focus();
						return false;
					}
					break;
					
				case "NOT_EMPTY":
					if(isEmpty(fieldRef.value))
					{	
						alert("Entry required! " + fieldDesc);
						if(fieldRef.getAttribute('type')!='hidden')
							fieldRef.focus();
						return false;
					}			
					break;
				
				case "INTEGER":
					if(!isEmpty(fieldRef.value) && !isInteger(fieldRef.value))
					{	alert("Invalid whole number! " + fieldDesc);
						fieldRef.focus();
						return false;
					}			
					break;
				case "PRICE":
					if(!isEmpty(fieldRef.value.replace(/\,/g, "")) && !isInteger(fieldRef.value.replace(/\,/g, "")))
					{	alert("Invalid whole number! " + fieldDesc);
						fieldRef.focus();
						return false;
					}			
					break;
//					regEx.test((fieldValue).replace(/\,/g, ""))
				case "FLOAT":
					if(!isEmpty(fieldRef.value) && !isFloat(fieldRef.value))
					{	alert("Invalid number! " + fieldDesc);
						fieldRef.focus();
						return false;
					}
					break;
				
				default:
					alert("System error, validation type " + tempArray[i] + " not supported.\n Validation on this type is ignored." );
			}
		}
	}
	return true;
}

function LTrim(str)
{
    var whitespace = new String(" \t\n\r");
    var s = new String(str);

    if (whitespace.indexOf(s.charAt(0)) != -1)
    {
        var j=0, i = s.length;
        while (j < i && whitespace.indexOf(s.charAt(j)) != -1)
        {
            j++;
        }
        s = s.substring(j, i);
    }
    return s;
}

function RTrim(str)
{
    var whitespace = new String(" \t\n\r");
    var s = new String(str);

    if (whitespace.indexOf(s.charAt(s.length-1)) != -1)
    {
        var i = s.length - 1;
        while (i >= 0 && whitespace.indexOf(s.charAt(i)) != -1)
        {
            i--;
        }
        s = s.substring(0, i+1);
    }
    return s;
}

function Trim(str)
{
    return RTrim(LTrim(str));
}

function isValidPwd(fieldValue)
{
	var matchCount = 0
	
	if(fieldValue.match(/\d/))
		matchCount++;
	
	if(fieldValue.match(/[a-z]/))
		matchCount++;
	
	if(fieldValue.match(/[A-Z]/))
		matchCount++;
	
	if(fieldValue.match(/\W/))
		matchCount++;
	
	if(matchCount>=3)
		return(true);
	else
		return(false);
}

function maskDecimal(inputElement, maxLength, decimalPlaces) {
    var i, exceptions = [8, 46, 37, 39, 13, 9];   // backspace, delete, arrowleft & right, enter, tab
    var isException = false;
    var isDot = ((46 == event.keyCode) && (new String(inputElement.value).indexOf(".") <= 0));       // dot
    var k = String.fromCharCode(event.keyCode);
    //alert(event.keyCode);
    var sel, rng, r2, curPos = -1;

    if (event.keyCode == 46) {
        if (new String(inputElement.value).indexOf(".") > 0) {
            return false;
        }
    } else {
        if (inputElement.value.length == maxLength) {
            if (new String(inputElement.value).indexOf(".") <= 0) {
                inputElement.value = inputElement.value + ".";
                return true;
            }
        }
    }

    if (event.keyCode == 32) {
            return false;
        }
    
    if (typeof inputElement.selectionStart == "number") {
        curPos = inputElement.selectionStart;
    } else if (document.selection && inputElement.createTextRange) {
        sel = document.selection;
        if (sel) {
            r2 = sel.createRange();
            rng = inputElement.createTextRange();
            rng.setEndPoint("EndToStart", r2);
            curPos = rng.text.length;
        }
    }
    
    for (i = 0; i < exceptions.length; i++)
        if (exceptions[i] == event.keyCode)
        isException = true;

    if (isNaN(k) && (!isException) && (!isDot))
        event.returnValue = false;
    else {
        var p = new String(inputElement.value + k).indexOf(".");

        if (((p < inputElement.value.length - decimalPlaces && curPos > p) || isDot) && p > -1 && (!isException))
            event.returnValue = false;
        else if (inputElement.value.length >= ((p > -1) || isDot ? maxLength + 1 + decimalPlaces : maxLength + decimalPlaces) && (!isException))         
            event.returnValue = false;
        else if (decimalPlaces == 0 && isDot)
            event.returnValue = false;
    }

}

function viewModeSetting(getForm) {
    var mode = getForm.editMode.value;

    if (mode == "V") {
        var link = getForm.document.getElementsByTagName('a');
        var count = link.length;

        if (count > 1) {
            for (i = 0; i < count; i++) {
                link[i].style.display = "none"; ;
            }

        }
    }
}

function maskKey(objEvent) {
    var iKeyCode;
    iKeyCode = objEvent.keyCode;
    if ((iKeyCode >= 48 && iKeyCode <= 57) || (iKeyCode == 46)) return true;
    return false;
}

function maskDate(objEvent) {
    var iKeyCode;
    iKeyCode = objEvent.keyCode;
    if ((iKeyCode >= 48 && iKeyCode <= 57) || (iKeyCode == 47)) return true;
    return false;
}

function maskTel(objEvent) {
    var iKeyCode;
    iKeyCode = objEvent.keyCode;
    if ((iKeyCode >= 48 && iKeyCode <= 57) || (iKeyCode == 47)) return true;
    return false;
}