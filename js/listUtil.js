function updateDirtyList(dirtyList, anId, oldValue, newValue)	
{
	if(oldValue != newValue)
	{	//append to list
		dirtyList.value = appendToList(dirtyList.value, anId);
	}
	else
	{	//remove from list
		dirtyList.value = removeFromList(dirtyList.value, anId);
	}
	//alert(dirtyList.value);
}

function appendToList(originalList, newListItem)
{	/* List format: |abc||def||efg| */
	/* this function append the new Value to the list if it does not already exist in the list */
	if(newListItem == null)
		return originalList;
	if(originalList == null)
		return "|" + newListItem + "|";	
	if(originalList.indexOf("|" + newListItem + "|") < 0)
	{	originalList += "|" + newListItem + "|";
	}
	return originalList;
}

function removeFromList(originalList, removeListItem)
{
	if(originalList.indexOf("|" + removeListItem + "|") >= 0)
	{				
		originalList = originalList.replace("|" + removeListItem + "|", "");
	}
	return originalList;
}

function changeSeperator(originalList, newSeperator)
{
	originalList = originalList.replace(/\|\|/gi, newSeperator);
	originalList = originalList.replace(/\|/gi, "");
	return originalList;
}

function inList(originalList, checkValue)
{
	if(originalList == null || checkValue == null)
		return false;
	if(originalList == "" || checkValue == "")
		return false;
	if(originalList.indexOf("|" + checkValue + "|") >= 0)
		return true;
	else
		return false;	
}

function listToArray(aList)
{
	if(aList.length - 2 <= 0)
		return new Array(0);
				
	aList = aList.substr(1, aList.length - 2)
	return aList.split("||");	
}


function listUnion(listA, listB)
{	
	if(listB == "")
		return listA;
	
	if(listA == "")
		return listB;

	var finalList = listA;		
	listB = changeSeperator(listB, ", ");
	
	var tempArray = listB.split(", ");
	for(var i=0; i<tempArray.length; i++)
	{	if(!inList(finalList, tempArray[i]))
		{	finalList = appendToList(finalList, tempArray[i]);
		}
	}
	return finalList;
}

function countListElement(aList)
{
	if(aList.length>0)
		return listToArray(aList).length;
}

function decodeString(str, compareStringList, resultValueList, defaultVal)
{
	var i, compareStringArray, resultValueArray
	
	compareStringArray = listToArray(compareStringList)
	resultValueArray = listToArray(resultValueList)
	for(i=0; i<compareStringArray.length; i++)
	{
		if(str == compareStringArray[i])
			return resultValueArray[i]
	}		
	return defaultVal
}