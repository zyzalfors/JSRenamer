let FSObj = new ActiveXObject("Scripting.FileSystemObject");

function formatPath(path) { 
 return (path.trim() + "\\").replace(/\\+/g,"\\");
}

function openDirectory() {
 let dirPath = formatPath(document.getElementById("dirPath").value);
 if(FSObj.FolderExists(dirPath) && dirPath !== "\\") {
  try {
   window.open(dirPath); 
  }
  catch(error) { 
   alert("Directory cannot be opened"); 
  }
 }
 else { 
   alert("Directory not found"); 
 }
}

function scrollNew() {
 document.getElementById("newNames").scrollTop = document.getElementById("oldNames").scrollTop;
 document.getElementById("newNames").scrollLeft = document.getElementById("oldNames").scrollLeft;
}

function scrollOld() {
 document.getElementById("oldNames").scrollTop = document.getElementById("newNames").scrollTop;
 document.getElementById("oldNames").scrollLeft = document.getElementById("newNames").scrollLeft;
}

function removeBlankLines(filter) {
 let text = document.getElementById(filter).value.split("\n").filter(function (text){ return text.search(/[^\s]/g) !== -1; }).join("\n");
 document.getElementById(filter).value = text;
}

function getName(fileName) {
 let pos = fileName.lastIndexOf(".");
 return pos !== -1 ? fileName.substring(0, pos) : fileName;
}

function getExtension(fileName) {
 let pos = fileName.lastIndexOf(".");
 return pos !== -1 ? fileName.substring(pos + 1) : "";
}

function toAlternatingCase1(text) {
 let newText = "";
 let toUp = true;
 for(let i = 0; i < text.length; i++) {
  if(text[i].toUpperCase() !== text[i].toLowerCase()) {
   if(toUp === true) {
    newText += text[i].toUpperCase();
    toUp = false;
   }
   else {
    newText += text[i].toLowerCase();
    toUp = true;
   }
  }
  else { 
  newText += text[i]; 
  }
 }
 return newText;
}

function toAlternatingCase2(text) {
 let newText = "";
 let toUp = false;
 for(let i = 0; i < text.length; i++) {
  if(text[i].toUpperCase() !== text[i].toLowerCase()) {
   if(toUp === true) {
    newText += text[i].toUpperCase();
    toUp = false;
   }
   else {
    newText += text[i].toLowerCase();
    toUp = true;
   }
  }
  else { 
   newText += text[i]; 
  }
 }
 return newText;
}

function padText(text, padder, length, leftOrRight) {
 let paddedTextChars = v.graphemes(text);
 let padderChars = v.graphemes(padder);
 let nPaddings = Math.ceil((length - paddedTextChars.length)/padderChars.length);
 if(!leftOrRight) { 
   for(let i = 0; i < nPaddings; i++) {
    paddedTextChars = paddedTextChars.concat(padderChars); 
   } 
 }
 else {
  for(let i = 0; i < nPaddings; i++) {
   padderChars = padderChars.concat(padderChars); 
  }
  paddedTextChars = padderChars.concat(paddedTextChars);
 } 
 let nCharsToBeRemoved = paddedTextChars.length - length;
 if(!leftOrRight) {
  while( nCharsToBeRemoved > 0 ) {
   paddedTextChars.pop();
   nCharsToBeRemoved--;
  }
 }
 else {
  while( nCharsToBeRemoved > 0 ) {
   paddedTextChars.shift();
   nCharsToBeRemoved--;
  }
 }
 let paddedText = "";
 for(let i = 0; i < paddedTextChars.length; i++) { 
   paddedText += paddedTextChars[i]; 
 }
 return paddedText;
}

function toCapitalizedCase(text) {
 let newText = "";
 let toUp = true;
 for (let i = 0; i < text.length; i++) {
  if(text[i].search(/\s/) === -1) {
   if (toUp === true) {
	newText += text[i].toUpperCase();
	toUp = false;
   }
   else {
    newText += text[i].toLowerCase(); 
   }
  }
  else {
   newText += text[i];
   toUp = true;      
  }
 }
 return newText;
}

function numbering(fileName, appendNum, prependNum, number) {
 let name = getName(fileName);
 let extension = getExtension(fileName);
 if(appendNum) { 
   name = name + number; 
 }
 if(prependNum) { 
  name = number + name; 
 }
 if(extension.trim() !== "") { 
  return name + "." + extension; 
 }
 return name;
}

function replaceWithMethod(fileName, regExp, method, modName, modExt) {
 let name = getName(fileName);
 let extension = getExtension(fileName);
 if(modExt && extension.trim() !== "") { 
  extension = extension.replace(regExp, function(match) { return method(match); }); 
 }
 if(modName) { 
  name = name.replace(regExp, function(match) { return method(match); }); 
 }
 if(extension.trim() !== "") { 
  return name + "." + extension; 
 }
 return name;
}

function replaceWithPad(fileName, regExp, padder, length, leftOrRight, modName, modExt) {
 let name = getName(fileName);
 let extension = getExtension(fileName);
 if(modExt && extension.trim() !== "") {
  extension = extension.replace(regExp, function(match) { return padText(match, padder, length, leftOrRight); }); 
 }
 if(modName) {
  name = name.replace(regExp, function(match) { return padText(match, padder, length, leftOrRight); }); 
 }
 if(extension.trim() !== "") {
  return name + "." + extension; 
 }
 return name;
}

function replaceWithReplacement(fileName, regExp, replacement, modName, modExt) {
 let name = getName(fileName);
 let extension = getExtension(fileName);
 if(modExt && extension.trim() !== "") { 
  extension = extension.replace(regExp, replacement); 
 }
 if(modName) { 
  name = name.replace(regExp, replacement); 
 }
 if(extension.trim() !== "") { 
  return name + "." + extension; 
 }
 return name;
}

function clearArea(filter) {
 if(filter === "Clear new names") { 
  document.getElementById("newNames").value = ""; 
 }
 if(filter === "Clear old names") { 
  document.getElementById("oldNames").value = ""; 
 }
 if(filter === "Clear methods") { 
  document.getElementById("methods").value = "{ \"methods\" : [] }"; 
 } 
}

function addMethod(filter) {
 let json = JSON.parse(document.getElementById("methods").value);
 let number = parseInt(document.getElementById("number").value);
 let increment = parseInt(document.getElementById("increment").value);
 let appendNum = document.getElementById("appendNum").checked;
 let prependNum = document.getElementById("prependNum").checked;
 let regex = document.getElementById("regex").value;
 let replacement = document.getElementById("replacement").value;
 let modExt = document.getElementById("modExt").checked;
 let modName = document.getElementById("modName").checked;
 let global = document.getElementById("global").checked;
 let caseIns = document.getElementById("caseIns").checked;
 let length = parseInt(document.getElementById("length").value);
 let padder = document.getElementById("padder").value;
 if(filter === "Numbering") {
  json.methods.push(
	            {
                     "name": filter, 
                     "startingNumber": number,
		     "increment" : increment, 
		     "prependNumber": prependNum, 
		     "appendNumber": appendNum
		    }
                   );
 }
 else if(filter === "Replace") {
  json.methods.push(
	            {
                     "name": filter, 
		     "regularExpression": regex,
		     "global" : global, 
		     "caseInsensitive" : caseIns, 
		     "replacement" : replacement, 
		     "modifyName" : modName, 
		     "modifyExtension" : modExt
		    }
                   );
 }
 else if(filter === "Pad right" || filter === "Pad left") {
  json.methods.push(
	            {
                     "name": filter, 
		     "regularExpression": regex,
		     "global" : global, 
		     "caseInsensitive" : caseIns, 
		     "padderText" : padder,
		     "length" : length,
		     "modifyName" : modName, 
		     "modifyExtension" : modExt
		    }
                   );
 }
 else {
  json.methods.push(
	            {
                     "name": filter, 
		     "regularExpression": regex, 
		     "global" : global, 
		     "caseInsensitive" : caseIns,
		     "modifyName" : modName, 
	    	     "modifyExtension" : modExt
		    }
                   );
 }
 document.getElementById("methods").value = JSON.stringify(json,null,1);
}

function getRegex(regex, global, caseIns) {
 try {
  if(global && caseIns) {
   return new RegExp(regex,"gi");
  }
  if(global && !caseIns) {
   return new RegExp(regex,"g");								
  }
  if(!global && caseIns) {
   return new RegExp(regex,"i");								
  }
  return new RegExp(regex);
 }
 catch(error) {
  return new RegExp(".+");
 }
}

function getNewName(fileName, method, regex, global, caseIns, replacement, number, increment, appendNum, prependNum, padder, length, modName, modExt) {
 let regExp = getRegex(regex, global, caseIns);
 switch (method) {	 
   case "Numbering" :
      number += increment;
      return numbering(fileName, appendNum, prependNum, number);
   break;

   case "Replace" :
      return replaceWithReplacement(fileName, regExp, replacement, modName, modExt);
   break;

   case "Upper case" :
      return replaceWithMethod(fileName, regExp, v.upperCase, modName, modExt);
   break;
   
   case "Lower case" :
      return replaceWithMethod(fileName, regExp, v.lowerCase, modName, modExt);
   break;
 
   case "Inverse case" :
      return replaceWithMethod(fileName, regExp, v.swapCase, modName, modExt);
   break;
   
   case "Alternating case 1" :
      return replaceWithMethod(fileName, regExp, toAlternatingCase1, modName, modExt);
   break;
   
   case "Alternating case 2" :
      return replaceWithMethod(fileName, regExp, toAlternatingCase2, modName, modExt);
   break;
   
   case "Capitalized case" :
      return replaceWithMethod(fileName, regExp, toCapitalizedCase, modName, modExt);
   break;
   
   case "Latinise" :
      return replaceWithMethod(fileName, regExp, v.latinise, modName, modExt);
   break;
   
   case "Reverse" :
      return replaceWithMethod(fileName, regExp, v.reverseGrapheme, modName, modExt);
   break;
   
   case "Trim" :
      return replaceWithMethod(fileName, regExp, v.trim, modName, modExt);
   break;
   
  case "Pad right" :
     return replaceWithPad(fileName, regExp, padder, length, false, modName, modExt);
  break;
  
  case "Pad left" :
     return replaceWithPad(fileName, regExp, padder, length, true, modName, modExt);
  break;
  
  default :
  return fileName;
 }
}

function getNewNames() {
 let oldNames = document.getElementById("oldNames").value.split("\n").filter(function (text){ return text.search(/[^\s]/g) !== -1; });
 let methods = JSON.parse(document.getElementById("methods").value).methods;
 let newNamesList = "";
 for(let i = 0; i < oldNames.length; i++) {
  let newName = oldNames[i];
  for(let j = 0; j < methods.length; j++) {
   newName = getNewName(newName, methods[j].name, methods[j].regularExpression, methods[j].global, methods[j].caseInsensitive, methods[j].replacement, methods[j].startingNumber, i*methods[j].increment, methods[j].appendNumber, methods[j].prependNumber, methods[j].padderText, methods[j].length, methods[j].modifyName, methods[j].modifyExtension);
  }
  newNamesList += newName + "\n";
 }
 document.getElementById("newNames").value = newNamesList.slice(0, -1);
}

function includes(container, content) {
 return container !== null ? container.indexOf(content) !== -1 : false;
}

function sortOldNames() {
 let dirPath = formatPath(document.getElementById("dirPath").value);
 let caseInsFilter = document.getElementById("caseInsFilter").checked;
 let regExp = getRegex(document.getElementById("filter").value, false, caseInsFilter);
 if(!FSObj.FolderExists(dirPath) || dirPath === "\\") {
  alert("Directory not found");
 }
 else {
  let files = new Enumerator(FSObj.GetFolder(dirPath).Files);
  let names = [];
  for(;!files.atEnd(); files.moveNext()) {
   let name = files.item().Name;
   if(includes(name.match(regExp), name)) {
    names.push(name);
   }
  }
  names.sort(new Intl.Collator(undefined, {numeric: true, sensitivity: 'base'}).compare);
  let namesList = names.join("\n");
  document.getElementById("oldNames").value = namesList;
  if(namesList === "") {
   alert("No file found");
  }
 }
}

function renameFiles() {
 let openDirAfter = document.getElementById("openDirAfter").checked;
 let dirPath =  formatPath(document.getElementById("dirPath").value);
 let oldNames = document.getElementById("oldNames").value.split("\n").filter(function (text){ return text.search(/[^\s]/g) !== -1; });
 let newNames = document.getElementById("newNames").value.split("\n").filter(function (text){ return text.search(/[^\s]/g) !== -1; });
 let n = 0;
 for(let i = 0; i < oldNames.length; i++) {
  if(oldNames[i] !== undefined && newNames[i] !== undefined && newNames[i] !== oldNames[i]) {
   if(FSObj.FileExists(dirPath + oldNames[i])) {
    n++;
    FSObj.MoveFile(dirPath + oldNames[i], dirPath + newNames[i]);
   }
  }
 }
 alert(n + " file(s) renamed");
 if (FSObj.FolderExists(dirPath) && dirPath !== "\\" && openDirAfter) {
  window.open(dirPath);
 }
}
