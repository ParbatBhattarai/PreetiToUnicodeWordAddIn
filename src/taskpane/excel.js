/* eslint-disable no-unused-vars */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("replace").onclick = replace;
    document.getElementById("test").onclick = test;

  }
});

export async function run() {
  return Excel.run(async context => {
    /**
     * Insert your Word code here
     */

    // const text1 = "hello hello hi hello";
    // insert a paragraph at the end of the document.
    try {
      // const paragraph = context.document.body.insertParagraph(text1, Word.InsertLocation.end);
      // console.log("Test");
      // const paragraph = context.document.body.insertParagraph(JSON.stringify(all_rules), Word.InsertLocation.end);
      // change the paragraph color to green.
      // paragraph.font.color = "green";

      var all_rules = {
        "preeti": {
          "name": "Preeti",
          "post-rules": [["्ा", ""], ["(त्र|त्त)([^उभप]+?)m", "$1m$2"], ["त्रm", "क्र"], ["त्तm", "क्त"], ["([^उभप]+?)m", "m$1"], ["उm", "ऊ"], ["भm", "झ"], ["पm", "फ"], ["इ{", "ई"], ["ि((.्)*[^्])", "$1ि"], ["(.[ािीुूृेैोौंःँ]*?){", "{$1"], ["((.्)*){", "{$1"], ["{", "र्"], ["([ाीुूृेैोौंःँ]+?)(्(.्)*[^्])", "$2$1"], ["्([ाीुूृेैोौंःँ]+?)((.्)*[^्])", "्$2$1"], ["([ंँ])([ािीुूृेैोौः]*)", "$2$1"], ["ँँ", "ँ"], ["ंं", "ं"], ["ेे", "े"], ["ैै", "ै"], ["ुु", "ु"], ["ूू", "ू"], ["^ः", ":"], ["टृ", "ट्ट"], ["ेा", "ाे"], ["ैा", "ाै"], ["अाे", "ओ"], ["अाै", "औ"], ["अा", "आ"], ["एे", "ऐ"], ["ाे", "ो"], ["ाै", "ौ"]],
          "v": "1.0.1",
          "char-map": {
            "÷": "/", "v": "ख", "r": "च", "\"": "ू", "~": "ञ्", "z": "श", "ç": "ॐ", "f": "ा", "b": "द", "n": "ल", "j": "व", "×": "×", "V": "ख्", "R": "च्", "ß": "द्म", "^": "६", "Û": "!", "Z": "श्", "F": "ँ", "B": "द्य", "N": "ल्", "Ë": "ङ्ग", "J": "व्", "6": "ट", "2": "द्द", "¿": "रू", ">": "श्र", ":": "स्", "§": "ट्ट", "&": "७", "£": "घ्", "•": "ड्ड", ".": "।", "«": "्र", "*": "८", "„": "ध्र", "w": "ध", "s": "क", "g": "न", "æ": "“", "c": "अ", "o": "य", "k": "प", "W": "ध्", "Ö": "=", "S": "क्", "Ò": "¨", "_": ")", "[": "ृ", "Ú": "’", "G": "न्", "ˆ": "फ्", "C": "ऋ", "O": "इ", "Î": "ङ्ख", "K": "प्", "7": "ठ", "¶": "ठ्ठ", "3": "घ", "9": "ढ", "?": "रु", ";": "स", "'": "ु", "#": "३", "¢": "द्घ", "/": "र", "+": "ं", "ª": "ङ", "t": "त", "p": "उ", "|": "्र", "x": "ह", "å": "द्व", "d": "म", "`": "ञ", "l": "ि", "h": "ज", "T": "त्", "P": "ए", "Ý": "ट्ठ", "\\": "्", "Ù": ";", "X": "ह्", "Å": "हृ", "D": "म्", "@": "२", "Í": "ङ्क", "L": "ी", "H": "ज्", "4": "द्ध", "±": "+", "0": "ण्", "<": "?", "8": "ड", "¥": "र्‍", "$": "४", "¡": "ज्ञ्", ",": ",", "©": "र", "(": "९", "‘": "ॅ", "u": "ग", "q": "त्र", "}": "ै", "y": "थ", "e": "भ", "a": "ब", "i": "ष्", "‰": "झ्", "U": "ग्", "Q": "त्त", "]": "े", "˜": "ऽ", "Y": "थ्", "Ø": "्य", "E": "भ्", "A": "ब्", "M": "ः", "Ì": "न्न", "I": "क्ष्", "5": "छ", "´": "झ", "1": "ज्ञ", "°": "ङ्ढ", "=": ".", "Æ": "”", "‹": "ङ्घ", "%": "५", "¤": "झ्", "!": "१", "-": "(", "›": "द्र", ")": "०", "…": "‘", "Ü": "%"
          }
        },
        "pcs nepali": {
          "name": "PCS Nepali",
          "post-rules": [["्ा", ""], ["(त्र|त्त)([^उभप]+?)m", "$1m$2"], ["त्रm", "क्र"], ["त्तm", "क्त"], ["([^उभप]+?)m", "m$1"], ["उm", "ऊ"], ["भm", "झ"], ["पm", "फ"], ["इ{", "ई"], ["ि((.्)*[^्])", "$1ि"], ["(.[ािीुूृेैोौंःँ]*?){", "{$1"], ["((.्)*){", "{$1"], ["{", "र्"], ["([ाीुूृेैोौंःँ]+?)(्(.्)*[^्])", "$2$1"], ["्([ाीुूृेैोौंःँ]+?)((.्)*[^्])", "्$2$1"], ["([ंँ])([ािीुूृेैोौः]*)", "$2$1"], ["ँँ", "ँ"], ["ंं", "ं"], ["ेे", "े"], ["ैै", "ै"], ["ुु", "ु"], ["ूू", "ू"], ["^ः", ":"], ["टृ", "ट्ट"], ["ेा", "ाे"], ["ैा", "ाै"], ["अाे", "ओ"], ["अाै", "औ"], ["अा", "आ"], ["एे", "ऐ"], ["ाे", "ो"], ["ाै", "ौ"]],
          "v": "1.0.0",
          "char-map": {"t": "त", "÷": "/", "v": "ख", "ñ": "ङ", "p": "उ", "r": "च", "|": "्र", "~": "ङ", "x": "ह", "z": "श", "å": "द्व", "d": "म", "ç": "ॐ", "f": "ा", "`": "ञ्", "b": "द", "í": "ष", "l": "ि", "n": "ल", "é": "ङ्ग", "h": "ज", "j": "व", "T": "त्", "V": "ख्", "P": "ए", "R": "च्", "\\": "्", "ß": "द्म", "^": "ट", "Ù": "ह", "X": "ह्", "Z": "श्", "D": "म्", "F": "ा", "@": "द्द", "B": "द्य", "L": "ी", "N": "ल्", "H": "ज्", "J": "व्", "4": "४", "·": "ट्ठ", "6": "६", "0": "०", "2": "२", "<": "्र", "¿": "रु", ">": "श्र", "8": "८", ":": "स्", "¥": "ऋ", "$": "द्ध", "§": "ट्ट", "&": "ठ", "¡": "ज्ञ्", "£": "घ्", "\"": "ू", ",": ",", ".": "।", "©": "?", "(": "ढ", "*": "ड", "u": "ग", "w": "ध", "q": "त्र", "s": "क", "}": "ै", "y": "थ", "ø": "य्", "ú": "ू", "e": "भ", "g": "न", "æ": "“", "a": "ब", "c": "अ", "o": "य", "i": "ष्", "k": "प", "U": "ग्", "Ô": "क्ष", "W": "ध्", "Q": "त्त", "S": "क्", "Ò": "ू", "]": "े", "_": ")", "Y": "थ्", "Ø": "्य", "[": "ृ", "E": "भ्", "G": "न्", "Æ": "”", "A": "ब्", "C": "र्‍", "M": "ः", "O": "इ", "I": "क्ष्", "K": "प्", "5": "५", "´": "झ", "7": "७", "1": "१", "°": "ङ्क", "3": "३", "=": ".", "?": "रू", "9": "९", ";": "स", "%": "छ", "¤": "ँ", "'": "ु", "!": "ज्ञ", "#": "घ", "¢": "द्घ", "-": "(", "/": "र", "®": "+", ")": "ण्", "+": "ं", "ª": "ञ"}
        },
        "kantipur": {
          "name": "Kantipur",
          "post-rules": [["्ा", ""], ["(त्र|त्त)([^उभप]+?)m", "$1m$2"], ["त्रm", "क्र"], ["त्तm", "क्त"], ["([^उभप]+?)m", "m$1"], ["उm", "ऊ"], ["भm", "झ"], ["पm", "फ"], ["इ{", "ई"], ["ि((.्)*[^्])", "$1ि"], ["(.[ािीुूृेैोौंःँ]*?){", "{$1"], ["((.्)*){", "{$1"], ["{", "र्"], ["([ाीुूृेैोौंःँ]+?)(्(.्)*[^्])", "$2$1"], ["्([ाीुूृेैोौंःँ]+?)((.्)*[^्])", "्$2$1"], ["([ंँ])([ािीुूृेैोौः]*)", "$2$1"], ["ँँ", "ँ"], ["ंं", "ं"], ["ेे", "े"], ["ैै", "ै"], ["ुु", "ु"], ["ूू", "ू"], ["^ः", ":"], ["टृ", "ट्ट"], ["ेा", "ाे"], ["ैा", "ाै"], ["अाे", "ओ"], ["अाै", "औ"], ["अा", "आ"], ["एे", "ऐ"], ["ाे", "ो"], ["ाै", "ौ"]],
          "v": "1.0.1",
          "char-map": {"÷": "/", "v": "ख", "r": "च", "\"": "ू", "~": "ञ्", "z": "श", "ç": "ॐ", "f": "ा", "b": "द", "n": "ल", "j": "व", "V": "ख्", "R": "च्", "ß": "द्म", "^": "६", "Z": "श्", "F": "ा", "B": "द्य", "Ï": "फ्", "N": "ल्", "Ë": "ङ्ग", "J": "व्", "6": "ट", "2": "द्द", "¿": "रू", ">": "श्र", ":": "स्", "§": "ट्ट", "&": "७", "£": "घ्", "•": "ड्ड", "¯": "¯", ".": "।", "«": "्र", "*": "८", "„": "ध्र", "w": "ध", "s": "क", "g": "न", "æ": "“", "c": "अ", "o": "य", "k": "प", "W": "ध्", "S": "क्", "Ò": "¨", "_": ")", "[": "ृ", "Ú": "’", "G": "न्", "Æ": "”", "C": "ऋ", "Â": "र", "O": "इ", "Î": "फ्", "K": "प्", "7": "ठ", "¶": "ठ्ठ", "3": "घ", "9": "ढ", "?": "रु", ";": "स", "º": "फ्", "'": "ु", "#": "३", "¢": "द्घ", "/": "र", "®": "र", "+": "ं", "ª": "ङ", "t": "त", "p": "उ", "|": "्र", "x": "ह", "å": "द्व", "d": "म", "`": "ञ", "l": "ि", "h": "ज", "T": "त्", "P": "ए", "Œ": "त्त्", "\\": "्", "X": "हृ", "D": "म्", "@": "२", "Í": "ङ्क", "L": "ी", "H": "ज्", "µ": "र", "4": "द्ध", "±": "+", "0": "ण्", "<": "?", "8": "ड", "¥": "र्‍", "$": "४", "¡": "ज्ञ्", "†": "!", "™": "र", "­": "(", ",": ",", "©": "र", "(": "९", "“": "ँ", "‘": "ॅ", "u": "ग", "q": "त्र", "}": "ै", "y": "थ", "ø": "य्", "e": "भ", "a": "ब", "i": "ष्", "‰": "झ्", "U": "ग्", "Ô": "क्ष", "Q": "त्त", "œ": "त्र्", "]": "े", "˜": "ऽ", "Y": "थ्", "Ø": "्य", "E": "भ्", "A": "ब्", "M": "ः", "Ì": "न्न", "I": "क्ष्", "È": "ष", "5": "छ", "´": "झ", "1": "ज्ञ", "°": "ङ्ढ", "=": ".", "‹": "ङ्ग", "%": "५", "¤": "झ्", "!": "१", "-": "(", "¬": "…", "›": "ऽ", ")": "०", "¨": "ङ्ग", "…": "‘"}
        }
      };

      var font = 'Preeti';
      font = font.toLowerCase();
      var myFont = all_rules[font];
      if (!myFont) {
        throw 'font not included in module';
      }

      //this finally works
      const currentSelection = context.document.getSelection();
      currentSelection.load('text');
      await context.sync();
      var textJson = JSON.parse(JSON.stringify(currentSelection));
      var text = textJson['text'];
      //this finally works

      var output = '';
      for (var w = 0; w < text.length; w++) {
        var letter = text[w];
        output += myFont['char-map'][letter] || letter;
      }
      for (var r = 0; r < myFont['post-rules'].length; r++) {
        output = output.replace(new RegExp(myFont['post-rules'][r][0], 'g'), myFont['post-rules'][r][1]);
      }

      const replace = currentSelection.insertText(output, "Replace");
      currentSelection.font.color = "black";

      // const newparagraph = context.document.body.insertParagraph(output, Word.InsertLocation.end);

    } catch (error) {
      context.document.body.insertParagraph(convertToUpperCase(JSON.stringify(error)), "End");
      // console.log("Error: " + JSON.stringify(error));
    }

    await context.sync().then(function () {
      // console.log('Inserted the text at the end of the selection.');
    });  
  });
}

export async function replace() {
  return Excel.run(async context => {
   
    // insert a paragraph at the end of the document.

    // currentSelection.clear();
    try {
      const currentSelection = context.document.getSelection();
      // const text = "currentSelection.toJSON";
      // const selectedText = currentSelection.text(); // {"name":"RichApi.Error","code":"PropertyNotLoaded","traceMessages":[],"innerError":null,"debugInfo":{"code":"PropertyNotLoaded","message":"The property 'text' is not available. Before reading the property's value, call the load method on the containing object and call \"context.sync()\" on the associated request context.","errorLocation":"Range.text"},"httpStatusCode":400}

      //this finally works
      currentSelection.load('text');
      await context.sync();
      var textJson = JSON.parse(JSON.stringify(currentSelection));
      const replace = currentSelection.insertText(textJson['text'], "End");
      //this finally works

      // const text1 = "newsReplacement ";
      // const replace2 = currentSelection.insertText(text1, "Replace");
      // replace.font.color = "green";

      // context.document.body.insertParagraph(convertToUpperCase(selectedText), "End");
      // context.document.body.insertParagraph(getText(), "End");
    } catch (error) {
      context.document.body.insertParagraph(JSON.stringify(error), "End");
      // console.log("Error: " + JSON.stringify(error));
    }

    

    // const paragraph = context.document.body.insertParagraph(currentSelection, Word.InsertLocation.end);
    // // change the paragraph color to blue.
    // paragraph.font.color = "green";

    await context.sync();
  }).catch(function(error) {
    // console.log("Error: " + JSON.stringify(error));
    const paragraph = context.document.body.insertParagraph(JSON.stringify(error), Word.InsertLocation.end);
  });
}

export async function test() {
  return Word.run(async context => {
    context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
          write('Action failed. Error: ' + asyncResult.error.message);
      }
      else {
          write('Selected data: ' + asyncResult.value);
      }
    });
  });
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message; 
}

function convertToUpperCase(message) {
  var originalString = message;
  return originalString.toUpperCase();
}

function getText() { 
  Word.run(function (context) {
      var customSelection = context.document.getSelection();
      return customSelection.text();
   })
}

