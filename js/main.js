Office.onReady();

async function addSpace() {
     await PowerPoint.run(async (context) => {
          let slides = context.presentation.getSelectedSlides();
          slides.load("items");
          await context.sync();
          let curSlideID = slides.items[0].id;
          context.presentation.load("slides");
          await context.sync();
          let activeSlide = context.presentation.slides.getItem(curSlideID);
          activeSlide.load("shapes");
          await context.sync();
          let shapes = activeSlide.shapes;
          shapes.load("items");
          await context.sync();
          for (let i = 0; i < shapes.items.length; i++) {
               let tmpObj = new Object();
               tmpObj.slideID = curSlideID;
               tmpObj.shapeID = shapes.items[i].id;
               let curRange = shapes.items[i].textFrame.textRange;
               shapes.items[i].textFrame.textRange.load("text");
               await context.sync();
               let curTextContent = shapes.items[i].textFrame.textRange.text;
               let regex1 = /([\u4E00-\u9FFF])([A-Za-z0-9])/g;
               let regex2 = /([A-Za-z0-9])([\u4E00-\u9FFF])/g;
               await addspace(regex1, 1);
               await addspace(regex2, 1);
               async function addspace(regex, insertIndex) {
                    while ((match = regex.exec(curTextContent)) != null) {
                         let tmpRange = curRange.getSubstring(match.index, 1);
                         tmpRange.load("text");
                         await context.sync();
                         console.log("Matched text content: " + match[0]);
                         console.log("Matched index: " + match.index);
                         console.log("original texRange: " + tmpRange.text);
                         tmpRange.text = match[1] + " ";
                         console.log("added texRange: " + tmpRange.text);
                         curTextContent = curTextContent.slice(0, match.index + insertIndex) + " " + curTextContent.slice(match.index + insertIndex);
                         console.log("updated total text rage: " + curTextContent);
                         console.log("===========================");
                    }
                    return context.sync();
               }
          }
     });
}

function setSuperscript() {
     PowerPoint.run(function (context) {
          const selectedTextRange = context.document.getSelectedDataAsync(Office.CoercionType.Text);

          return context
               .sync()
               .then(function () {
                    const selectedText = selectedTextRange.value;
                    const range = context.document.getSelection();

                    // Apply superscript formatting to the selected text
                    range.font.superscript = true;

                    return context.sync();
               })
               .catch(function (error) {
                    // Handle any errors
                    console.error(error);
               });
     });
}
function setSubscript() {
     PowerPoint.run(function (context) {
          const selectedTextRange = context.document.getSelectedDataAsync(Office.CoercionType.Text);

          return context
               .sync()
               .then(function () {
                    const selectedText = selectedTextRange.value;
                    const range = context.document.getSelection();

                    // Apply subscript formatting to the selected text
                    range.font.subscript = true;

                    return context.sync();
               })
               .catch(function (error) {
                    // Handle any errors
                    console.error(error);
               });
     });
}

async function copyPureTextContent() {
     Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, async function (result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
               let str = result.value;
               if (str[str.length - 1].match(/[\n\r]/)) {
                    str = str.substr(0, str.length - 2);
               }
               await createTxtElement(document, str);
               console.log(str);

               function createTxtElement(document, str) {
                    // Create new element
                    var el = document.createElement("textarea");
                    // Set value (string to be copied)
                    el.value = str;
                    // Set non-editable to avoid focus and move outside of view
                    el.setAttribute("readonly", "");
                    el.style = { position: "absolute", left: "-9999px" };
                    document.body.appendChild(el);
                    // Select text inside element
                    el.select();
                    // Copy text to clipboard
                    document.execCommand("copy");
                    // Remove temporary element
                    document.body.removeChild(el);
                    return new Promise((r) => setTimeout(r, 10));
               }
          } else {
               const error = result.error;
               // 在此處處理錯誤
               console.error(error);
          }
     });
}

async function compareFileDifference_source() {
     await PowerPoint.run(async (context) => {
          let logAry = [];
          let slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          for (let i = 0; i < slides.items.length; i++) {
               curPageShapes = slides.items[i].shapes;
               curSlideID = slides.items[i].id;
               curPageShapes.load("items");
               await context.sync();
               for (let k = 0; k < curPageShapes.items.length; k++) {
                    let el = new Object();
                    el.p = curSlideID;
                    el.s = curPageShapes.items[k].id;
                    if (curPageShapes.items[k].type == "Group") {
                         let subgroup = curPageShapes.items[k].shapes;
                         console.log(curPageShapes.items[k].shapes);
                         subgroup.load("items");
                         await context.sync();
                    }
                    try {
                         curPageShapes.items[k].textFrame.textRange.load("text");
                         await context.sync();
                         el.t = curPageShapes.items[k].textFrame.textRange.text;
                    } catch (err) {
                         el.t = "";
                    }
                    logAry.push(el);
               }
          }
          logAry.sort((a, b) => {
               if (a.slideID != b.p) {
                    return a.p - b.p;
               } else {
                    return a.s - b.s;
               }
          });
          // console.log(JSON.stringify(logAry));

          outcome = JSON.stringify(logAry);
          await createTxtElement(document, outcome);
          console.log(outcome);
          // console.log(JSON.stringify(logAry));
          // createTxtElement(document, JSON.stringify(logAry));
          function createTxtElement(document, str) {
               tmpEle = document.getElementById("tmp");
               // Create new element
               var el = document.createElement("textarea");
               // Set value (string to be copied)
               el.value = str;
               // Set non-editable to avoid focus and move outside of view
               el.setAttribute("readonly", "");
               el.style = { position: "absolute", left: "-9999px" };
               document.body.appendChild(el);
               // Select text inside element
               el.select();
               // Copy text to clipboard
               console.log();
               document.execCommand("copy");
               // Remove temporary element
               document.body.removeChild(el);
               return new Promise((r) => setTimeout(r, 10));
          }
     });
}

async function compareFileDifference_target() {
     await PowerPoint.run(async (context) => {
          let sourceDataLog = false;
          await readPasteBoard();
          let logAry = [];
          let slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          for (let i = 0; i < slides.items.length; i++) {
               curPageShapes = slides.items[i].shapes;
               curSlideID = slides.items[i].id;
               curPageShapes.load("items");
               await context.sync();
               for (let k = 0; k < curPageShapes.items.length; k++) {
                    let el = new Object();
                    el.p = curSlideID;
                    el.s = curPageShapes.items[k].id;
                    try {
                         curPageShapes.items[k].textFrame.textRange.load("text");
                         await context.sync();
                         el.t = curPageShapes.items[k].textFrame.textRange.text;
                    } catch (err) {
                         console.log("Err" + el.s);
                         el.t = "";
                    }
                    logAry.push(el);
                    let findMatchContentIdx = sourceDataLog.findIndex((obj) => {
                         // console.log(obj);
                         return obj.p == el.p && obj.s == el.s && obj.t.replace(/[\n\s\r]/g, "") == el.t.replace(/[\n\s\r]/g, "");
                    });
                    console.log(findMatchContentIdx);
                    if (findMatchContentIdx != -1 && el.t != "") {
                         try {
                              if (el.t == sourceDataLog[findMatchContentIdx].t) {
                                   curPageShapes.items[k].textFrame.textRange.font.color = "#EEEEEE";
                                   curPageShapes.items[k].fill.setSolidColor("#FFFFFF");
                                   curPageShapes.items[k].fill.transparency = 0.1;
                              } else {
                                   curPageShapes.items[k].textFrame.textRange.font.color = "#EEE1E0";
                                   curPageShapes.items[k].fill.setSolidColor("#FFFFEE");
                                   curPageShapes.items[k].fill.transparency = 0.1;
                              }
                         } catch (err) {}
                    }
               }
          }
          logAry.sort((a, b) => {
               if (a.slideID != b.p) {
                    return a.p - b.p;
               } else {
                    return a.s - b.s;
               }
          });
          // console.log(JSON.stringify(logAry));

          outcome = JSON.stringify(logAry);
          // await createTxtElement(document, outcome);
          console.log(outcome);
          // console.log(JSON.stringify(logAry));
          // createTxtElement(document, JSON.stringify(logAry));
          function readPasteBoard() {
               if (navigator.clipboard) {
                    // Read the clipboard data
                    navigator.clipboard
                         .readText()
                         .then(function (clipboardData) {
                              console.log("Clipboard data:", clipboardData);
                              sourceDataLog = JSON.parse(clipboardData);
                         })
                         .catch(function (error) {
                              sourceDataLog = false;
                              console.log("Error reading clipboard data:", error);
                         });
               } else {
                    sourceDataLog = false;
                    console.log("Clipboard API is not supported in this browser.");
               }
               return new Promise((r) => setTimeout(r, 10));
               // return sourceDataLog;
          }
     });
}
