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
               console.log("start");
               if (shapes.items[i].type != "GeometricShape") {
                    console.log("not geoShape");
                    continue;
               }
               shapes.items[i].textFrame.load("hasText");
               await context.sync();
               if (!shapes.items[i].textFrame.hasText) {
                    console.log("not Text");
                    continue;
               }
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
                    // if (curPageShapes.items[k].type == "Group") {
                    //      let subgroup = curPageShapes.items[k].shapes;
                    //      console.log(curPageShapes.items[k].shapes);
                    //      subgroup.load("items");
                    //      await context.sync();
                    // }
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

let textStyles = [];

async function readItemStyle(presetIndex) {
     await PowerPoint.run(async (context) => {
          let selectedShapes = context.presentation.getSelectedShapes();
          let shapeCount = selectedShapes.getCount();
          selectedShapes.load("items");
          await context.sync();
          textStyles[presetIndex] = {
               left: Math.round(selectedShapes.items[0].left),
               top: Math.round(selectedShapes.items[0].top),
               width: Math.round(selectedShapes.items[0].width),
               height: Math.round(selectedShapes.items[0].height),
               right: Math.round(selectedShapes.items[0].left + selectedShapes.items[0].width),
               bottom: Math.round(selectedShapes.items[0].top + selectedShapes.items[0].height),
          };
          console.log(textStyles[presetIndex]);
          // Get the position information
          // Now you can use 'left' and 'top' to determine the position of the shape.
     });
     applyItemStyle(presetIndex, [-1, -1], false);
     $("button[data-textStyleIndex=" + presetIndex + "]").html(
          "←: " +
               ("___" + Math.floor(textStyles[presetIndex].left)).slice(-3) +
               "<br>  ↑: " +
               ("___" + Math.floor(textStyles[presetIndex].top)).slice(-3) +
               // "→: " +
               // Math.floor(textStyles[presetIndex].right) +
               // "; ↓: " +
               // Math.floor(textStyles[presetIndex].bottom) +
               "<br> 寬: " +
               ("___" + Math.floor(textStyles[presetIndex].width)).slice(-3) +
               "<br>  高: " +
               ("___" + Math.floor(textStyles[presetIndex].height)).slice(-3)
     );
}

async function applyItemStyle(presetIndex, posOpt, bySizeOpt) {
     let selectorEle = document.querySelectorAll('input[name="fontDropper' + presetIndex + '"]');
     let applyWidtHeight = [];
     console.log(bySizeOpt);
     if (bySizeOpt) {
          selectorEle.forEach((e) => {
               applyWidtHeight.push(e.checked);
          });
     } else {
          applyWidtHeight = [true, true];
     }
     console.log(applyWidtHeight);

     await PowerPoint.run(async (context) => {
          let selectedShapes = context.presentation.getSelectedShapes();
          let shapeCount = selectedShapes.getCount();
          selectedShapes.load("items");
          await context.sync();
          selectedShapes.items.map((shape) => {
               shape.load("left,top,width,height");
          });
          await context.sync();
          selectedShapes.items.map((shape) => {
               if (applyWidtHeight[0]) {
                    shape.width = textStyles[presetIndex].width;
               }
               if (applyWidtHeight[1]) {
                    shape.height = textStyles[presetIndex].height;
               }
               shape.load("width,height");
               if (posOpt[0] == -1) {
                    shape.left = textStyles[presetIndex].left;
               } else if (posOpt[0] == 1) {
                    shape.left = textStyles[presetIndex].right - shape.width;
               }
               if (posOpt[1] == -1) {
                    shape.top = textStyles[presetIndex].top;
               } else if (posOpt[1] == 1) {
                    shape.top = textStyles[presetIndex].bottom - shape.height;
               }
          });
          // Get the position information
          // Now you can use 'left' and 'top' to determine the position of the shape.
     });
}

let titelStyle = {};
async function readTitleStyle() {
     await PowerPoint.run(async (context) => {
          let selectedShapes = context.presentation.getSelectedShapes();
          selectedShapes.load("items");
          await context.sync();
          selectedShapes.items[0].name = "Title textbox";
          let textRange = selectedShapes.items[0].textFrame.textRange;
          selectedShapes.items[0].load("textFrame");
          await context.sync();
          textRange.load("font");
          await context.sync();
          textRange.load("paragraphFormat");
          await context.sync();
          console.log(selectedShapes.items[0].textFrame.textRange.font);
          // console.log(selectedShapes.items[0].type);
          titelStyle = {
               left: Math.round(selectedShapes.items[0].left),
               top: Math.round(selectedShapes.items[0].top),
               width: Math.round(selectedShapes.items[0].width),
               height: Math.round(selectedShapes.items[0].height),
               right: Math.round(selectedShapes.items[0].left + selectedShapes.items[0].width),
               bottom: Math.round(selectedShapes.items[0].top + selectedShapes.items[0].height),
               fontSize: selectedShapes.items[0].textFrame.textRange.font.size,
               fontColor: selectedShapes.items[0].textFrame.textRange.font.color,
               fontFamily: selectedShapes.items[0].textFrame.textRange.font.name,
               bold: selectedShapes.items[0].textFrame.textRange.font.bold,
               italic: selectedShapes.items[0].textFrame.textRange.font.italic,
               xAlign: selectedShapes.items[0].textFrame.textRange.paragraphFormat.horizontalAlignment,
               vAlign: selectedShapes.items[0].textFrame.verticalAlignment,
               ml: selectedShapes.items[0].textFrame.leftMargin,
               mr: selectedShapes.items[0].textFrame.rightMargin,
               mt: selectedShapes.items[0].textFrame.topMargin,
               mb: selectedShapes.items[0].textFrame.bottomMargin,
          };
     });
}

async function applyTitleStyle() {
     let selectorEle = document.querySelectorAll('input[name="fontDropperTitle"]');
     let applyFormat = [];

     selectorEle.forEach((e) => {
          applyFormat.push(e.checked);
     });
     console.log(applyFormat);
     console.log(titelStyle);
     await PowerPoint.run(async (context) => {
          let selectedSlides = context.presentation.getSelectedSlides();
          selectedSlides.load("items");
          await context.sync();

          for (let i = 0; i < selectedSlides.items.length; i++) {
               slide = selectedSlides.items[i];
               slide.load("shapes");
               await context.sync();
               let slideShapes = slide.shapes;
               slideShapes.load("items");
               await context.sync();
               for (let k = 0; k < slideShapes.items.length; k++) {
                    let shape = slideShapes.items[k];
                    if (shape.type != "GeometricShape") {
                         console.log("not geoShape");
                         continue;
                    }
                    if (shape.top > ($("#titleJudge_Top").val() * 540) / 100) {
                         continue;
                    }
                    if (shape.left > ($("#titleJudge_Left").val() * 960) / 100) {
                         continue;
                    }

                    let curRange = shape.textFrame.textRange;
                    try {
                         shape.textFrame.textRange.load("font");
                         await context.sync();
                    } catch (err) {
                         console.log("Err");
                         continue;
                    }

                    shape.load("width,height");
                    shape.name = "Title textbox";
                    console.log("A");
                    shape.width = titelStyle.width;
                    shape.height = titelStyle.height;
                    shape.left = titelStyle.left;
                    shape.top = titelStyle.top;
                    console.log("B");

                    shape.textFrame.textRange.paragraphFormat.horizontalAlignment = titelStyle.xAlign;
                    // shape.textFrame.verticalAlignment = PowerPoint.TextVerticalAlignment.top;
                    console.log("H");
                    if (applyFormat[0]) {
                         shape.textFrame.textRange.font.name = titelStyle.fontFamily;
                    }
                    console.log("K");
                    if (applyFormat[1]) {
                         shape.textFrame.textRange.font.color = titelStyle.fontColor;
                    }
                    console.log("R");
                    if (applyFormat[2]) {
                         shape.textFrame.textRange.font.size = titelStyle.fontSize;
                    }
                    console.log("Z");
                    if (applyFormat[3]) {
                         shape.textFrame.textRange.font.bold = titelStyle.bold;
                    }
                    if (applyFormat[4]) {
                         shape.textFrame.textRange.font.italic = titelStyle.italic;
                    }
                    if (applyFormat[5]) {
                         shape.textFrame.verticalAlignment = titelStyle.vAlign;
                    }
                    try {
                         shape.textFrame.leftMargin = titelStyle.ml;
                    } catch (err) {}
                    try {
                         shape.textFrame.rightMargin = titelStyle.mr;
                    } catch (err) {}
                    try {
                         shape.textFrame.topMargin = titelStyle.mt;
                    } catch (err) {}
                    try {
                         shape.textFrame.bottomMargin = titelStyle.mb;
                    } catch (err) {}
                    console.log("End");
               }
          }
     });
}

let refStyle = {};
async function readRefStyle() {
     await PowerPoint.run(async (context) => {
          let selectedShapes = context.presentation.getSelectedShapes();
          selectedShapes.load("items");
          await context.sync();
          let textRange = selectedShapes.items[0].textFrame.textRange;
          selectedShapes.items[0].name = "Ref textbox";
          selectedShapes.items[0].load("textFrame");
          await context.sync();
          textRange.load("font");
          await context.sync();
          textRange.load("paragraphFormat");
          await context.sync();
          console.log(selectedShapes.items[0].textFrame.textRange.font);
          refStyle = {
               left: Math.round(selectedShapes.items[0].left),
               top: Math.round(selectedShapes.items[0].top),
               width: Math.round(selectedShapes.items[0].width),
               height: Math.round(selectedShapes.items[0].height),
               right: Math.round(selectedShapes.items[0].left + selectedShapes.items[0].width),
               bottom: Math.round(selectedShapes.items[0].top + selectedShapes.items[0].height),
               fontSize: selectedShapes.items[0].textFrame.textRange.font.size,
               fontColor: selectedShapes.items[0].textFrame.textRange.font.color,
               fontFamily: selectedShapes.items[0].textFrame.textRange.font.name,
               bold: selectedShapes.items[0].textFrame.textRange.font.bold,
               italic: selectedShapes.items[0].textFrame.textRange.font.italic,
               xAlign: selectedShapes.items[0].textFrame.textRange.paragraphFormat.horizontalAlignment,
               vAlign: selectedShapes.items[0].textFrame.verticalAlignment,
               ml: selectedShapes.items[0].textFrame.leftMargin,
               mr: selectedShapes.items[0].textFrame.rightMargin,
               mt: selectedShapes.items[0].textFrame.topMargin,
               mb: selectedShapes.items[0].textFrame.bottomMargin,
          };
     });
}

async function applyRefStyle() {
     let selectorEle = document.querySelectorAll('input[name="fontDropperRef"]');
     let applyFormat = [];

     selectorEle.forEach((e) => {
          applyFormat.push(e.checked);
     });

     await PowerPoint.run(async (context) => {
          let selectedSlides = context.presentation.getSelectedSlides();
          selectedSlides.load("items");
          await context.sync();

          await context.sync();
          selectedSlides.items.sort(function (a, b) {
               return b.bottom - a.bottom;
          });
          console.log(selectedSlides.items);
          for (let i = 0; i < selectedSlides.items.length; i++) {
               slide = selectedSlides.items[i];
               slide.load("shapes");
               await context.sync();
               let slideShapes = slide.shapes;
               slideShapes.load("items");
               await context.sync();
               let curRefShape = null;
               let curRefShapeBottom = null;
               let curSlideRefShapes = [];
               refObjCount = 0;
               for (let k = 0; k < slideShapes.items.length; k++) {
                    let shape = slideShapes.items[k];
                    try {
                         console.log("start");
                         if (shape.type != "GeometricShape") {
                              console.log("not geoShape");
                              continue;
                         }
                         shape.textFrame.load("hasText");
                         await context.sync();
                         if (!shape.textFrame.hasText) {
                              console.log("not Text");
                              continue;
                         }
                         // console.log("loading pos");
                         shape.load("width");
                         await context.sync();
                         shape.load("height");
                         await context.sync();

                         // console.log("pos");
                         if (shape.top + shape.height < ((100 - $("#refJudge_Top").val()) * 540) / 100) {
                              continue;
                         }
                         if (shape.left > ($("#refJudge_Left").val() * 960) / 100) {
                              continue;
                         }

                         // let curRange = shape.textFrame.textRange;
                         try {
                              shape.load("textFrame");
                              await context.sync();
                              shape.textFrame.textRange.load("font");
                              await context.sync();
                              shape.textFrame.textRange.load("paragraphFormat");
                              await context.sync();
                         } catch (err) {
                              console.log("Err");
                              continue;
                         }
                         // console.log("loading text");
                         shape.textFrame.textRange.load("text");
                         await context.sync();
                         // console.log(shape.textFrame.textRange.text.toString()[0]);
                         // if (curRefShape == null) {
                         shape = shape;
                         shape.name = "Ref textbox";
                         if (refObjCount > 0) {
                              shape.name = "SubRef" + refObjCount;
                         }
                         // shape.bottom = shape.top + shape.height;
                         curSlideRefShapes.push(shape);
                         shape.textFrame.autoSizeSetting = PowerPoint.ShapeAutoSize.AutoSizeNone;
                         shape.width = refStyle.width;
                         shape.height = refStyle.height;
                         shape.left = refStyle.left;
                         shape.top = refStyle.top - refObjCount * 20;
                         refObjCount++;
                         // console.log("size");
                         // console.log(shape.textFrame.autoSizeSetting);
                         // console.log("autosize");
                         // console.log(refStyle.ml);

                         // console.log("margin");
                         if (applyFormat[0]) {
                              shape.textFrame.textRange.font.name = refStyle.fontFamily;
                         }
                         if (applyFormat[1]) {
                              shape.textFrame.textRange.font.color = refStyle.fontColor;
                         }
                         if (applyFormat[2]) {
                              shape.textFrame.textRange.font.size = refStyle.fontSize;
                         }
                         console.log(shape.textFrame.textRange.font.size);
                         if (applyFormat[3]) {
                              shape.textFrame.textRange.font.bold = refStyle.bold;
                         }
                         if (applyFormat[4]) {
                              shape.textFrame.textRange.font.italic = refStyle.italic;
                         }
                         if (applyFormat[5]) {
                              shape.textFrame.verticalAlignment = refStyle.vAlign;
                         }
                         // console.log("font");
                         shape.textFrame.textRange.paragraphFormat.horizontalAlignment = refStyle.xAlign;
                         // console.log("vA");
                         shape.textFrame.leftMargin = refStyle.ml;
                         shape.textFrame.rightMargin = refStyle.mr;
                         shape.textFrame.topMargin = refStyle.mt;
                         shape.textFrame.bottomMargin = refStyle.mb;
                         // console.log("xA");
                         // } else {
                         //      // shape.textFrame.textRange.load("text");
                         //      // await context.sync();
                         //      // console.log(shape.textFrame.textRange.text);
                         //      curRefShape.textFrame.textRange.load("text");
                         //      await context.sync();
                         //      console.log(curRefShape.textFrame.textRange.text);
                         //      if (curRefShapeBottom < shape.top + shape.height) {
                         //           let insertPoint = curRefShape.textFrame.textRange.text.length - 1;
                         //           let tmpRange = curRefShape.textFrame.textRange.getSubstring(curRefShape.textFrame.textRange.text.length - 1, 1);
                         //           tmpRange.load("text");
                         //           await context.sync();
                         //           console.log("A");
                         //           console.log(tmpRange.text[0]);
                         //           tmpRange.text = tmpRange.text[0] + "\n ";
                         //           let insertRange = curRefShape.textFrame.textRange.getSubstring(insertPoint + 2, 1);
                         //           insertRange.load("text");
                         //           await context.sync();
                         //           insertRange.text = shape.textFrame.textRange.text;
                         //           // insertRange = shape.textFrame.textRange;
                         //           // curRefShape.textFrame.textRange = +"\n" + curRefShape.shape.textFrame.textRangetextFrame.textRange;
                         //           // console.log(curRefShape.textFrame.textRange.text.toString()[0]);
                         //           // curRefShape.textFrame.textRange.text = curRefShape.textFrame.textRange.text + "\n" + shape.textFrame.textRange.text;
                         //      } else {
                         //           console.log("B: ");
                         //           let insertPoint = 0;
                         //           let tmpRange = curRefShape.textFrame.textRange.getSubstring(0, 1);
                         //           tmpRange.load("text");
                         //           await context.sync();
                         //           tmpRange.text = " \n" + curRefShape.textFrame.textRange.text[0];
                         //           let insertRange = curRefShape.textFrame.textRange.getSubstring(insertPoint, 1);
                         //           insertRange.load("text");
                         //           await context.sync();
                         //           insertRange.text = shape.textFrame.textRange.text;
                         //           // curRefShape.textFrame.textRange.text = shape.textFrame.textRange.text + "\n" + curRefShape.textFrame.textRange.text;
                         //      }
                         //      shape.delete();
                         //      k--;
                         //      console.log("k-");
                         //      continue;
                         // }
                    } catch (err) {
                         console.log(err);
                    }
                    console.log("nRound");
               }
               curSlideRefShapes.sort(function (a, b) {
                    return b.bottom - a.bottom;
               });
               console.log(curSlideRefShapes);
               for (let k = 0; k < curSlideRefShapes.length - 1; k++) {
                    console.log(curSlideRefShapes[k].name);
                    curSlideRefShapes[k].name = curSlideRefShapes[k].name + "-" + k;
                    console.log(curSlideRefShapes[k].name);
                    curSlideRefShapes[k].top = curSlideRefShapes[k].top - 5 * k;
               }
          }
     });
}

$(function () {
     $('[data-bs-toggle="tooltip"]').tooltip();
});
