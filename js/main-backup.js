Office.onReady();
let ary_registedAbbreObjIDs = [];
let ary_registedTableObjIDandContents = [];
// let abbreCsvFilePath = "";
async function registAbbreObj() {
     let alt = 0;
     try {
          if (event.altKey == 1) {
               alt = 1;
          }
     } catch (err) {}
     await PowerPoint.run(async (context) => {
          let slides = context.presentation.getSelectedSlides();
          slides.load("items");
          await context.sync();
          let curSlideID = slides.items[0].id;
          // console.log(slides.items[0].id);
          let shapes = context.presentation.getSelectedShapes();
          let shapeCount = shapes.getCount();
          shapes.load("items");
          await context.sync();
          shapes.items.map((shape) => {
               let tmpObj = new Object();
               tmpObj.slideID = curSlideID;
               tmpObj.shapeID = shape.id;
               let checkRegisted = ary_registedAbbreObjIDs.findIndex((obj) => {
                    return obj.slideID == tmpObj.slideID && obj.shapeID == tmpObj.shapeID;
               });
               if (alt && checkRegisted != -1) {
                    ary_registedAbbreObjIDs.splice(checkRegisted, 1);
               } else if (!alt && checkRegisted == -1) {
                    ary_registedAbbreObjIDs.push(tmpObj);
               }
               console.log(ary_registedAbbreObjIDs);
               // console.log(shape.id);
               // console.log(shape);
               // document.getElementById("outcome").innerText = shape.id;
          });
          await context.sync();
     });
}
async function registTableObj() {
     let alt = 0;
     try {
          if (event.altKey == 1) {
               alt = 1;
          }
     } catch (err) {}
     await PowerPoint.run(async (context) => {
          let slides = context.presentation.getSelectedSlides();
          slides.load("items");
          await context.sync();
          let curSlideID = slides.items[0].id;

          let shapes = context.presentation.getSelectedShapes();
          let shapeCount = shapes.getCount();
          if (shapes.getCount() > 0) {
               alert("一次只能記錄一個物件");
               return;
          }
          shapes.load("items");
          await context.sync();
          let tmpObj = new Object();
          tmpObj.slideID = curSlideID;
          tmpObj.shapeID = shapes.items[0].id;
          await Office.context.document.getSelectedDataAsync(
               "text", // coercionType
               {
                    valueFormat: "unformatted", // valueFormat
                    filterType: "all",
               }, // filterType
               function (result) {
                    // callback
                    console.log(result.value);
                    tmpObj.contents = result.value;
                    //   write('Selected data is: ' + dataValue);
               }
          );

          let checkRegistedID = ary_registedTableObjIDandContents.findIndex((obj) => {
               return obj.slideID == tmpObj.slideID && obj.shapeID == tmpObj.shapeID;
          });
          if (checkRegistedID != -1) {
               ary_registedTableObjIDandContents.splice(checkRegistedID, 1);
          }
          if (!alt) {
               ary_registedTableObjIDandContents.push(tmpObj);
          }

          console.log(ary_registedTableObjIDandContents.length);
          // console.log(shape.id);
          // console.log(shape);
          // document.getElementById("outcome").innerText = shape.id;
          await context.sync();
     });
}
async function listAbbreofActivePage() {
     await PowerPoint.run(async (context) => {
          let tmpCount = context.presentation.getSelectedShapes().getCount();
          await context.sync();
          // console.log(tmpCount.value);
          if (tmpCount.value > 0) {
               registAbbreObj();
          }
     });
     let existedAbbs = [];
     let allAbbs = [];
     let allRemoved = [];
     let registedAbbreContents = "";
     readAbbreCsvFile();
     await PowerPoint.run(async (context) => {
          let curPageContents = "";
          let IDofUndetectedItems = [];
          let slides = context.presentation.getSelectedSlides();
          slides.load("items");
          await context.sync();
          let curSlideID = slides.items[0].id;
          context.presentation.load("slides");
          await context.sync();
          // console.log(curSlideID);
          let activeSlide = context.presentation.slides.getItem(curSlideID);
          activeSlide.load("shapes");
          await context.sync();
          let shapes = activeSlide.shapes;
          shapes.load("items");
          await context.sync();
          for (let i = 0; i < shapes.items.length; i++) {
               // shapes.items.map((shape) => {
               let tmpObj = new Object();
               tmpObj.slideID = curSlideID;
               tmpObj.shapeID = shapes.items[i].id;
               let checkAbbreRegisted = ary_registedAbbreObjIDs.findIndex((obj) => {
                    return obj.slideID == tmpObj.slideID && obj.shapeID == tmpObj.shapeID;
               });
               let checkTableRegisted = ary_registedTableObjIDandContents.find((obj) => {
                    return obj.slideID == tmpObj.slideID && obj.shapeID == tmpObj.shapeID;
               });

               try {
                    if (checkTableRegisted != undefined) {
                         curPageContents += checkTableRegisted.contents;
                         console.log(checkTableRegisted.contents);
                    } else {
                         shapes.items[i].textFrame.textRange.load("text");
                         await context.sync();
                         if (checkAbbreRegisted == -1) {
                              curPageContents += shapes.items[i].textFrame.textRange.text;
                              curPageContents += "\n";
                         } else {
                              registedAbbreContents = shapes.items[i].textFrame.textRange.text;
                         }
                    }
               } catch (err) {
                    IDofUndetectedItems.push(shapes.items[i].id);
               }
          }
          if (IDofUndetectedItems.length > 0) {
               console.log("有無法偵測的物件");
               activeSlide.setSelectedShapes(IDofUndetectedItems);
          }
          console.log(curPageContents);

          let allEngWords = curPageContents.match(/[A-Za-z0-9αβγμ][A-Za-z0-9αβγμ\-\/]*[A-Za-z0-9αβγμ]+/g);

          allEngWords = allEngWords.filter(function (element, index, self) {
               return self.indexOf(element) === index;
          });

          if (curPageContents.match(/\b([a-z]\.)+\b/g) != null) {
               let tmpAry = curPageContents.match(/\b([a-z]\.)+\b/g);
               if (allEngWords == null) {
                    allEngWords = [];
               }
               for (let k = 0; k < tmpAry.length; k++) {
                    allEngWords.push(tmpAry[k]);
               }
          }
          let removedWords = [];

          if (allEngWords != null) {
               // sortIgnoreUpperCase(allEngWords);
               allEngWords.sort();
               // let tmpWord = "";
               for (let k = 0; k < allEngWords.length; k++) {
                    // if (tmpWord == allEngWords[k]) {
                    //      console.log("dup " + allEngWords[k]);
                    //      allEngWords.splice(k, 1);
                    //      continue;
                    // }
                    tmpWord = allEngWords[k].toString();
                    // await console.log("pre " + tmpWord);
                    if (allEngWords[k].match(/\./)) {
                         continue;
                    }
                    if (
                         allEngWords[k].match(/[A-Z]/g) == null ||
                         allEngWords[k].match("CODECODE") ||
                         (allEngWords[k].match(/[X0\-]/g) != null && allEngWords[k].match(/[X0\-]/g).length == allEngWords[k].length)
                    ) {
                         allEngWords.splice(k, 1);
                         k--;
                         continue;
                    }

                    if (allEngWords[k] == "mmHg") {
                         allEngWords.splice(k, 1);
                         k--;
                         continue;
                    }

                    if (allEngWords[k] == "Ph" || allEngWords[k] == "Hb" || allEngWords[k] == "Af") {
                         continue;
                    }

                    if (allEngWords[k] == "EP") {
                         removedWords.push(allEngWords[k]);
                         allEngWords.splice(k, 1);
                         k--;
                         continue;
                    }

                    if (allEngWords[k].match(/[0-9\-]/)) {
                         if (allEngWords[k].match(/[0-9\-]/g).length >= allEngWords[k].length / 2) {
                              if (allEngWords[k].match(/[0-9\-]/g).length != allEngWords[k].length) {
                                   removedWords.push(allEngWords[k]);
                              }
                              allEngWords.splice(k, 1);
                              k--;
                              continue;
                         }
                    }
                    if (allEngWords[k].match(/[a-z]/)) {
                         if (allEngWords[k].match(/[a-z\-]/g).length >= allEngWords[k].length / 2) {
                              if (
                                   allEngWords[k].match(/[A-Z][a-z\-\/][a-z\-\/]+/) != null &&
                                   allEngWords[k].match(/[A-Z][a-z\-\/][a-z\-\/]+/)[0].length == allEngWords[k].length
                              ) {
                              } else {
                                   removedWords.push(allEngWords[k]);
                              }
                              allEngWords.splice(k, 1);
                              k--;
                              continue;
                         }
                    }
               }
          }

          if (allEngWords != null) {
               // await sortIgnoreUpperCase(allEngWords);
               // allAbbs = allAbbs.concat(allEngWords).sort(function compare(a, b) {
               //      a.toUpperCase() - b.toUpperCase();
               // });
               sortIgnoreUpperCase(allEngWords);
          }
          if (removedWords != null) {
               sortIgnoreUpperCase(removedWords);
               // allRemoved = allRemoved.concat(removedWords).sort(function compare(a, b) {
               //      a.toUpperCase() - b.toUpperCase();
               // });
          }
          let registedAbbreContents_modifier = registedAbbreContents.split(/[\n\r]/);
          for (let i = 0; i < registedAbbreContents_modifier.length; i++) {
               let tmpIndex = registedAbbreContents_modifier[i].match(/[0-9]+\.[\s ]*[A-Z]/);
               // console.log(tmpIndex);
               if (tmpIndex != undefined) {
                    registedAbbreContents_modifier.splice(i, 1);
                    i--;
                    continue;
               }
               if (registedAbbreContents_modifier[i].match(/[*†‡§]/) != undefined) {
                    registedAbbreContents_modifier.splice(i, 1);
                    i--;
                    continue;
               }
          }

          let excistedAbbreList = registedAbbreContents_modifier.join().split(/[\s ]*;[\s ]*/);
          console.log(excistedAbbreList);
          let mainAbbreList = allEngWords;
          let suspectList = removedWords;
          let excistedAbbreList_ObjedAry = [];
          let mainAbbreList_filtered = mainAbbreList;
          for (let i = 0; i < excistedAbbreList.length; i++) {
               if (excistedAbbreList[i].toString().match(/[\:=,]/) != null && excistedAbbreList[i].toString().match(/[\:=,]/).length > 1) {
                    let tmpAry = excistedAbbreList[i].toString().split(/,[\s]*/);
                    excistedAbbreList.splice(i, 1, tmpAry);
                    continue;
               } else {
                    let tmpObj = new Object();
                    if (excistedAbbreList[i].split(/[\s ]*[\:=,][\s ]*/)[0] != "") {
                         tmpObj.abbre = excistedAbbreList[i].split(/[\s ]*[\:=,][\s ]*/)[0];
                         tmpObj.full = excistedAbbreList[i].split(/[\s ]*[\:=,][\s ]*/)[1];
                         excistedAbbreList_ObjedAry.push(tmpObj);
                    }
               }
          }

          let mainAbbreList_matched = [];
          let databaseRefedList = [];
          let unmatchedList = [];
          let newAbbreToUpdateAry = [];
          for (let i = 0; i < mainAbbreList_filtered.length; i++) {
               let exsistedMatchObj = excistedAbbreList_ObjedAry.find((obj) => {
                    return obj.abbre == mainAbbreList_filtered[i];
               });
               if (mainAbbreList_filtered[i].match(/[0-9]+\b/) && exsistedMatchObj == undefined) {
                    exsistedMatchObj = excistedAbbreList_ObjedAry.find((obj) => {
                         return obj.abbre == mainAbbreList_filtered[i].replace(/[0-9]+\b/, "");
                    });
                    if (exsistedMatchObj != undefined) {
                         mainAbbreList_filtered[i] = mainAbbreList_filtered[i].replace(/[0-9]+\b/, "");
                    }
               }

               let databaseMatchObjAry = abbreDatabaseAry.filter((obj) => {
                    return obj.abbre == mainAbbreList_filtered[i];
               });
               let full = "";
               if (exsistedMatchObj != undefined) {
                    full = exsistedMatchObj.full;
                    excistedAbbreList_ObjedAry.splice(
                         excistedAbbreList_ObjedAry.findIndex((obj) => {
                              return obj.abbre == exsistedMatchObj.abbre;
                         }),
                         1
                    );
                    if (
                         abbreDatabaseAry.filter((obj) => {
                              return obj.full.toLowerCase().replace(/./g, "") == exsistedMatchObj.full.toLowerCase().replace(/./g, "");
                         }).length == 0
                    ) {
                         // if (databaseMatchObjAry != undefined) {
                         //      alert(databaseMatchObjAry.full + "\r" + exsistedMatchObj.full);
                         // }
                         // newAbbreToUpdateAry_Obj.push("'" + exsistedMatchObj.abbre + "', '" + exsistedMatchObj.full + "'");

                         //待補充資料庫中有多筆不同資料時的處理
                         newAbbreToUpdateAry.push(mainAbbreList_filtered[i] + " = " + full);
                    }
               } else if (databaseMatchObjAry != 0) {
                    full = databaseMatchObjAry[0].full;
                    unmatchedList.push(mainAbbreList_filtered[i] + " = " + full);
                    databaseRefedList.push(mainAbbreList_filtered[i] + " = " + full);
               } else if (mainAbbreList_filtered[i].match(/[A-Z]+\/[A-Z]+/)) {
                    let tmpAry = mainAbbreList_filtered[i].split("/");
                    for (let k = 0; k < tmpAry.length; k++) {
                         if (mainAbbreList_filtered.indexOf(tmpAry[k]) === -1) {
                              mainAbbreList_filtered.push(tmpAry[k]);
                         }
                    }
                    mainAbbreList_filtered.splice(i, 1);
                    i--;
                    continue;
               } else {
                    unmatchedList.push(mainAbbreList_filtered[i] + " = ");
                    full = "_______________";
               }

               mainAbbreList_matched.push(mainAbbreList_filtered[i] + " = " + full);
          }

          let suspectList_filtered = [];
          for (let i = 0; i < suspectList.length; i++) {
               if (suspectList_filtered.indexOf(suspectList[i]) === -1) {
                    suspectList_filtered.push(suspectList[i]);
               }
          }

          let unusedRefedList = [];
          for (let i = 0; i < excistedAbbreList_ObjedAry.length; i++) {
               if (excistedAbbreList_ObjedAry[i].abbre.match(" ")) {
                    let tmpRematchAry = [];
                    let tmpSplitAbbreAry = excistedAbbreList_ObjedAry[i].abbre.toString().split(" ");
                    let rematchFail = false;
                    for (let k = 0; k < tmpSplitAbbreAry.length; k++) {
                         // alert(mainAbbreList_filtered.indexOf(tmpSplitAbbreAry[k]));
                         if (mainAbbreList_filtered.indexOf(tmpSplitAbbreAry[k]) == -1) {
                              rematchFail = true;
                              continue;
                         } else {
                              tmpRematchAry.push[k];
                         }
                    }
                    if (rematchFail) {
                         continue;
                    }
                    mainAbbreList_matched.push(excistedAbbreList_ObjedAry[i].abbre + " = " + excistedAbbreList_ObjedAry[i].full);
                    tmpRematchAry.sort();
                    for (let k = 0; k < tmpRematchAry[i]; k++) {
                         suspectList_filtered.splice(tmpRematchAry.pop(), 1);
                    }
               } else if (suspectList_filtered.indexOf(excistedAbbreList_ObjedAry[i].abbre) != -1) {
                    mainAbbreList_matched.push(excistedAbbreList_ObjedAry[i].abbre + " = " + excistedAbbreList_ObjedAry[i].full);
                    suspectList_filtered.splice(suspectList_filtered.indexOf(excistedAbbreList_ObjedAry[i].abbre), 1);
               } else {
                    unusedRefedList.push(excistedAbbreList_ObjedAry[i].abbre + " = " + excistedAbbreList_ObjedAry[i].full);
               }
          }

          let mergedMatchedList = mainAbbreList_matched.join("; ").toString();
          let mergedUnmatchedList = unmatchedList.join("; ").toString();
          let mergedUnusedRefedList = unusedRefedList.join("; ").toString();
          let mergedDatabaseRefedList = databaseRefedList.join("; ").toString();
          let mergedSuspectList = suspectList_filtered.join("; ").toString();
          let newAbbreToUpdateList = newAbbreToUpdateAry.join("; ").toString();

          let saveString = mergedMatchedList + "\r\r縮寫表清單:\r" + mergedMatchedList.replace(/; /g, "\r") + "\r\r";
          if (mergedUnmatchedList != "") {
               saveString += "\r\r待補上的縮寫:\r" + mergedUnmatchedList.replace(/; /g, "\r");
          }
          if (mergedUnusedRefedList != "") {
               saveString += "\r\r未使用的縮寫:\r" + mergedUnusedRefedList.replace(/; /g, "\r");
          }
          if (mergedDatabaseRefedList != "") {
               saveString += "\r\r\r\r取自資料庫的縮寫:\r" + mergedDatabaseRefedList.replace(/; /g, "\r");
          }
          if (mergedSuspectList != "") {
               saveString += "\r\r\r\r疑似縮寫字:\r" + mergedSuspectList;
          }
          if (newAbbreToUpdateList != "") {
               saveString += "\r\r\r\r資料庫新增:\r" + newAbbreToUpdateList;
          }

          // abbreTextFile.write(saveString);
          // await console.log("filtered\n" + allEngWords);
          // console.log("removed\n" + removedWords);
          document.getElementById("outcome").innerText = saveString;
     });
}
async function sayHello() {
     try {
          // Set coercion type to text since
          let coercionType = { coercionType: Office.CoercionType.Text };

          // clear current selection
          let outcome = "";
          // await Office.context.document.getResourceByIndexAsync(
          //      1,

          //      function (result) {
          //           // callback
          //           outcome = result.value;
          //           console.log(outcome);
          //           document.getElementById("outcome").innerText = outcome;
          //           //   write('Selected data is: ' + dataValue);
          //      }
          // );
          // await PowerPoint.run(async (context) => {
          //      let slides = context.presentation.getSelectedSlides();
          //      slides.load("items");
          //      await context.sync();
          //      slides.items.map((slide) => {
          //           console.log(slide.id);
          //      });
          //      let shapes = context.presentation.getSelectedShapes();

          //      let shapeCount = shapes.getCount();
          //      shapes.load("items");
          //      await context.sync();
          //      shapes.items.map((shape) => {
          //           shape.fill.setSolidColor("red");
          //           document.getElementById("outcome").innerText = shape.id;
          //      });
          //      await context.sync();
          // });
          await Office.context.document.getSelectedDataAsync(
               "text", // coercionType
               {
                    valueFormat: "unformatted", // valueFormat
                    filterType: "all",
               }, // filterType
               function (result) {
                    // callback
                    outcome = result.value;
                    console.log(outcome);
                    document.getElementById("outcome").innerText = outcome;
                    //   write('Selected data is: ' + dataValue);
               }
          );
     } catch (err) {
          console.log(err, err.line);
     }

     // Set text in selection to 'Hello world!'
     // await Office.context.document.setSelectedDataAsync("Hello world!", coercionType);
}
async function sortIgnoreUpperCase(ary) {
     await ary.sort();
     await ary.sort(function compare(a, b) {
          a.toUpperCase() - b.toUpperCase();
     });
     // await console.log(ary);
}
function readAbbreCsvFile() {
     try {
          let xhr = new XMLHttpRequest();
          xhr.open("GET", "/js/database/abbreListDatabase.csv", false);
          xhr.onload = function () {
               let inputTxt = xhr.responseText;

               let basicAry = inputTxt.replace(/[\r\n]+/g, "\n").split("\n");
               abbreDatabaseAry = [];
               // alert(basicAry);
               basicAry.forEach(function (row, i) {
                    let tmpObject = new Object();
                    tmpObject.abbre = splitCsvRow(row)[0].toString();
                    tmpObject.full = splitCsvRow(row)[1].toString();
                    abbreDatabaseAry.push(tmpObject);
               });
          };
          // reader.readAsText(file);
          xhr.send();
     } catch (err) {
          alert(err.line + "\n" + err);
     }
}
