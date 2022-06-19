let FSObj = new ActiveXObject("Scripting.FileSystemObject");

function formatPath(path)
{
 return (path.trim() + "\\").replace(/\\+/g, "\\");
}

function openDirectory()
{
 let dirPath = formatPath(document.getElementById("dirPath").value);
 if(FSObj.FolderExists(dirPath) && dirPath !== "\\")
 {
  try
  {
   window.open(dirPath);
  }
  catch(error)
  {
   window.alert("Directory cannot be opened");
  }
 }
 else
 {
  window.alert("Directory not found");
 }
}

function scrollArea(filter)
{
 switch(filter)
 {
  case "oldNames":
   document.getElementById("newNames").scrollTop = document.getElementById("oldNames").scrollTop;
   document.getElementById("newNames").scrollLeft = document.getElementById("oldNames").scrollLeft;
   break;
  case "newNames":
   document.getElementById("oldNames").scrollTop = document.getElementById("newNames").scrollTop;
   document.getElementById("oldNames").scrollLeft = document.getElementById("newNames").scrollLeft;
   break;
  default:
 }
}

function removeEmptyLines(filter)
{
 document.getElementById(filter).value = document.getElementById(filter).value.split("\n").filter(function(text){ return text.search(/[^\s]/g) !== -1; }).join("\n");
}

function getName(fileName)
{
 let pos = fileName.lastIndexOf(".");
 return pos !== -1 ? fileName.substring(0, pos) : fileName;
}

function getExtension(fileName)
{
 let pos = fileName.lastIndexOf(".");
 return pos !== -1 ? fileName.substring(pos + 1) : "";
}

function toAlternating(text, toUp)
{
 let manipulatedText = "";
 for(let i = 0; i < text.length; i++)
 {
  if(text[i].toUpperCase() !== text[i].toLowerCase())
  {
   if(toUp === true)
   {
    manipulatedText += text[i].toUpperCase();
    toUp = false;
   }
   else
   {
    manipulatedText += text[i].toLowerCase();
    toUp = true;
   }
  }
  else
  {
   manipulatedText += text[i];
  }
 }
 return manipulatedText;
}

function toAlternating1(text)
{
 return toAlternating(text, true);
}
 
function toAlternating2(text)
{
 return toAlternating(text, false);
}

function numbering(fileName, appendNum, prependNum, number)
{
 let name = getName(fileName);
 let extension = getExtension(fileName);
 if(appendNum)
 {
  name = name + number;
 }
 if(prependNum)
 {
  name = number + name;
 }
 if(extension.trim() !== "")
 {
  return name + "." + extension;
 }
 return name;
}

function padText(text, pad, length, leftOrRight)
{
 return leftOrRight ? v.padLeft(text, length, pad) : v.padRight(text, length, pad);
}

function removeSpaces(text)
{
 return text.replace(/\s+/g, "");
}

function removeRedundantSpaces(text)
{
 return text.replace(/\s+/g, " ");
}

function replaceUsingMethod(fileName, regExp, method, modName, modExt)
{
 let name = getName(fileName);
 let extension = getExtension(fileName);
 if(modExt && extension.trim() !== "")
 {
  extension = extension.replace(regExp, function(match){ return method(match); });
 }
 if(modName)
 {
  name = name.replace(regExp, function(match){ return method(match); });
 }
 if(extension.trim() !== "")
 {
  return name + "." + extension;
 }
 return name;
}

function replaceUsingPad(fileName, regExp, pad, length, leftOrRight, modName, modExt)
{
 let name = getName(fileName);
 let extension = getExtension(fileName);
 if(modExt && extension.trim() !== "")
 {
  extension = extension.replace(regExp, function(match){ return padText(match, pad, length, leftOrRight); });
 }
 if(modName)
 {
  name = name.replace(regExp, function(match){ return padText(match, pad, length, leftOrRight); });
 }
 if(extension.trim() !== "")
 {
  return name + "." + extension;
 }
 return name;
}

function replaceUsingReplacement(fileName, regExp, replacement, modName, modExt)
{
 let name = getName(fileName);
 let extension = getExtension(fileName);
 if(modExt && extension.trim() !== "")
 {
  extension = extension.replace(regExp, replacement);
 }
 if(modName)
 {
  name = name.replace(regExp, replacement);
 }
 if(extension.trim() !== "")
 {
  return name + "." + extension;
 }
 return name;
}

function clearMethods()
{ 
 document.getElementById("methods").value = "{ \"methods\" : [] }";
}

function addMethod()
{
 let method = document.getElementById("methodList").value;
 if(method === "null")
 {
  return;
 }
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
 let pad = document.getElementById("pad").value;
 switch(method)
 {
  case "Numbering":
   json.methods.push(
                     {
                      "method": method,
                      "startNum": number,
                      "increment" : increment,
                      "prependNum": prependNum,
                      "appendNum": appendNum
                     }
                    );
   break;
  case "Replace":
   json.methods.push(
                     {
                      "method": method,
                      "regex": regex,
                      "global" : global,
                      "caseIns" : caseIns,
                      "replacement" : replacement,
                      "modName" : modName,
                      "modExt" : modExt
                     }
                    );
   break;
  case "Pad right":
  case "Pad left":
   json.methods.push(
                     {
                      "method": method,
                      "regex": regex,
                      "global" : global,
                      "caseIns" : caseIns,
                      "pad" : pad,
                      "length" : length,
                      "modName" : modName,
                      "modExt" : modExt
                     }
                    );
   break;
  default:
   json.methods.push(
                     {
                      "method": method,
                      "regex": regex,
                      "global" : global,
                      "caseIns" : caseIns,
                      "modName" : modName,
                      "modExt" : modExt
                     }
                    );
 }
 document.getElementById("methods").value = JSON.stringify(json, null, 1);
}

function getRegex(regex, global, caseIns)
{
 if(global && caseIns)
 {
  return new RegExp(regex,"gi");
 }
 if(global && !caseIns)
 {
  return new RegExp(regex,"g");
 }
 if(!global && caseIns)
 {
  return new RegExp(regex,"i");
 }
 return new RegExp(regex);
}

function getNewName(fileName, method, regex, global, caseIns, replacement, number, increment, appendNum, prependNum, pad, length, modName, modExt)
{
 let regExp = getRegex(regex, global, caseIns);
 switch(method)
 {
  case "Numbering":
   number += increment;
   return numbering(fileName, appendNum, prependNum, number);
   break;
  case "Replace":
   return replaceUsingReplacement(fileName, regExp, replacement, modName, modExt);
   break;
  case "Convert to upper case":
   return replaceUsingMethod(fileName, regExp, v.upperCase, modName, modExt);
   break;
  case "Convert to lower case":
   return replaceUsingMethod(fileName, regExp, v.lowerCase, modName, modExt);
   break;
  case "Convert to inverse case":
   return replaceUsinghMethod(fileName, regExp, v.swapCase, modName, modExt);
   break;
  case "Convert to alternating case 1":
   return replaceUsingMethod(fileName, regExp, toAlternating1, modName, modExt);
   break;
  case "Convert to alternating case 2":
   return replaceUsingMethod(fileName, regExp, toAlternating2, modName, modExt);
   break;
  case "Convert to title case":
   return replaceUsingMethod(fileName, regExp, v.titleCase, modName, modExt);
   break;
  case "Convert to camel case":
   return replaceUsingMethod(fileName, regExp, v.camelCase, modName, modExt);
   break;
  case "Convert to kebab case":
   return replaceUsingMethod(fileName, regExp, v.kebabCase, modName, modExt);
   break;
  case "Convert to slug case":
   return replaceUsingMethod(fileName, regExp, v.slugify, modName, modExt);
   break;
  case "Convert to snake case":
   return replaceUsingMethod(fileName, regExp, v.snakeCase, modName, modExt);
   break;
  case "Latinise":
   return replaceUsingMethod(fileName, regExp, v.latinise, modName, modExt);
   break;
  case "Reverse":
   return replaceUsingMethod(fileName, regExp, v.reverseGrapheme, modName, modExt);
   break;
  case "Pad right":
   return replaceUsingPad(fileName, regExp, pad, length, false, modName, modExt);
   break;
  case "Pad left":
   return replaceUsingPad(fileName, regExp, pad, length, true, modName, modExt);
   break;
  case "Remove spaces":
   return replaceUsingMethod(fileName, regExp, removeSpaces, modName, modExt);
   break;
  case "Remove redundant spaces":
   return replaceUsingMethod(fileName, regExp, removeRedundantSpaces, modName, modExt);
   break;
  default:
  return fileName;
 }
}

function getNewNames()
{
 let oldNames = document.getElementById("oldNames").value.split("\n").map(function(text){ return text.trim(); }).filter(function(text){ return text.search(/[^\s]/g) !== -1; });
 let methods = JSON.parse(document.getElementById("methods").value).methods;
 let newNamesList = "";
 for(let i = 0; i < oldNames.length; i++)
 {
  let newName = oldNames[i];
  for(let j = 0; j < methods.length; j++)
  {
   try
   {
    newName = getNewName(newName, methods[j].method, methods[j].regex, methods[j].global, methods[j].caseIns, methods[j].replacement, methods[j].startNum, i*methods[j].increment, methods[j].appendNum, methods[j].prependNum, methods[j].pad, methods[j].length, methods[j].modName, methods[j].modExt);
   }
   catch(error)
   {
    window.alert(error.toString());
    return;
   }
  }
  newNamesList += newName + "\n";
 }
 document.getElementById("newNames").value = newNamesList.slice(0, -1);
}

function includes(container, content)
{
 return container !== null ? container.indexOf(content) !== -1 : false;
}

function sortOldNames()
{
 let dirPath = formatPath(document.getElementById("dirPath").value);
 if(!FSObj.FolderExists(dirPath) || dirPath === "\\")
 {
  window.alert("Directory not found");
  return;
 }
 let caseInsFilter = document.getElementById("caseInsFilter").checked;
 let regExp;
 try
 {
  regExp = getRegex(document.getElementById("filter").value, false, caseInsFilter);
 }
 catch(error)
 {
  window.alert(error.toString());
  return;  
 }
 let files = new Enumerator(FSObj.GetFolder(dirPath).Files);
 let names = [];
 for(;!files.atEnd(); files.moveNext())
 {
  let name = files.item().Name;
  if(includes(name.match(regExp), name))
  {
   names.push(name);
  }
 }
 names.sort(new Intl.Collator(undefined, {numeric: true, sensitivity: 'base'}).compare);
 let namesList = names.join("\n");
 document.getElementById("oldNames").value = namesList;
 if(namesList === "")
 {
  window.alert("No file found");
 }
}

function renameFiles()
{
 let openDirAfter = document.getElementById("openDirAfter").checked;
 let dirPath = formatPath(document.getElementById("dirPath").value);
 let oldNames = document.getElementById("oldNames").value.split("\n").map(function(text){ return text.trim(); }).filter(function(text){ return text.search(/[^\s]/g) !== -1; });
 let newNames = document.getElementById("newNames").value.split("\n").map(function(text){ return text.trim(); }).filter(function(text){ return text.search(/[^\s]/g) !== -1; });
 let n = 0;
 for(let i = 0; i < oldNames.length; i++)
 {
  if(oldNames[i] !== undefined && newNames[i] !== undefined && newNames[i] !== oldNames[i])
  {
   if(FSObj.FileExists(dirPath + oldNames[i]))
   {
    n++;
    FSObj.MoveFile(dirPath + oldNames[i], dirPath + newNames[i]);
   }
  }
 }
 window.alert(n + " file(s) renamed");
 if(FSObj.FolderExists(dirPath) && dirPath !== "\\" && openDirAfter)
 {
  window.open(dirPath);
 }
}