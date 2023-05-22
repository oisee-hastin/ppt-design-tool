var script = document.currentScript;
var fullUrl = script.src;
var rootPath = fullUrl.toString().replace("file:///", "").replace("main.js", "") + "../../../Ai/";
// var csInterface = new CSInterface();
var colorList = [];

var mainAbbreList_matched = [];

let abbreDatabaseAry;
let abbreCsvFilePath = fullUrl.toString().replace("file:///", "").replace("main.js", "") + "database/abbreListDatabase.csv";
abbreCsvFilePath = decodeURI(abbreCsvFilePath).toString();
function readAbbreCsvFile() {
     try {
          let xhr = new XMLHttpRequest();
          xhr.open("GET", abbreCsvFilePath, false);
          xhr.onload = function () {
               let inputTxt = xhr.responseText;

               let basicAry = inputTxt.replace(/[\r\n]+/g, "\n").split("\n");
               abbreDatabaseAry = [];
               basicAry.forEach(function (row, i) {
                    let tmpObject = new Object();
                    tmpObject.abbre = splitCsvRow(row)[0].toString();
                    tmpObject.full = splitCsvRow(row)[1].toString();
                    abbreDatabaseAry.push(tmpObject);
               });
          };

          xhr.send();
     } catch (err) {
          alert(err.line + "\n" + err);
     }
}

function splitCsvRow(textdata) {
     let csvColAry = [];
     let quoted = false;
     let curContent = "";
     csvColAry.push(curContent);
     for (let k = 0; k < textdata.length; k++) {
          if (textdata[k] == '"') {
               quoted = !quoted;
          } else if (textdata[k] == "," && !quoted) {
               csvColAry.push("");
          } else {
               csvColAry[csvColAry.length - 1] += textdata[k];
          }
     }
     return csvColAry;
}

function collectAbbreScript() {
     if (event.shiftKey == 1) {
          csInterface.evalScript('$.evalFile("' + rootPath + 'Text Tools/sortAbbre.jsx")');
     } else {
          try {
               readAbbreCsvFile();
               let allTxtContent;
               PowerPoint.run(function (context) {
                    // 獲取當前幻燈片和當前選定的幻燈片範圍
                    var slide = context.presentation.getActiveSlide();
                    var selection = context.selection;

                    // 如果選擇的是整個幻燈片，則設置幻燈片範圍為全部
                    if (selection.slides.length > 1) {
                         selection.load("slides");
                         return context.sync().then(function () {
                              slide = selection.slides.getItemAt(0);
                         });
                    }

                    // 獲取當前幻燈片的文本框，並遍歷其中的文本內容
                    var shapes = slide.shapes;
                    shapes.load("items");
                    return context.sync().then(function () {
                         var text = "";
                         for (var i = 0; i < shapes.items.length; i++) {
                              var shape = shapes.items[i];
                              if (shape instanceof PowerPoint.TextFrame) {
                                   var paragraphs = shape.textRange.paragraphs;
                                   paragraphs.load("text");
                                   context.sync().then(function () {
                                        for (var j = 0; j < paragraphs.items.length; j++) {
                                             text += paragraphs.items[j].text + "\n";
                                        }
                                   });
                              }
                         }
                         allTxtContent = text;
                         console.log(text);
                    });
               });

               var allEngWords = allTxtContent.contents.match(/[A-Za-z0-9αβγμ][A-Za-z0-9αβγμ\-\/]*[A-Za-z0-9αβγμ]+/g);

               if (allTxtContent.contents.match(/\b([a-z]\.)+\b/g) != null) {
                    var tmpAry = allTxtContent.contents.match(/\b([a-z]\.)+\b/g);
                    if (allEngWords == null) {
                         allEngWords = [];
                    }
                    for (var k = 0; k < tmpAry.length; k++) {
                         allEngWords.push(tmpAry[k]);
                    }
               }
               var removedWords = [];

               if (allEngWords != null) {
                    for (var k = 0; k < allEngWords.length; k++) {
                         // if (allEngWords[k].match(/[A-Z]?[a-z\-]+/)) {
                         //     if (allEngWords[k].length == allEngWords[k].match(/[A-Z]?[a-z\-]+/).toString().length) {
                         //         removedWords.push(allEngWords[k].toString())
                         //         allEngWords.splice(k, 1)
                         //         k--
                         //         continue
                         //     }
                         // }
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

                         // if (!allEngWords[k].match(/[A-Z]/)) {
                         //     allEngWords.splice(k, 1)
                         //     k--
                         //     continue
                         // }
                         // if (allEngWords[k].match(/[A-Z]/g).length <= allEngWords[k].length / 2) {
                         //     allEngWords.splice(k, 1)
                         //     k--
                         //     continue
                         // }
                    }
               }

               if (allEngWords != null) {
                    allAbbs = allAbbs.concat(allEngWords).sort(function compare(a, b) {
                         a.toUpperCase() - b.toUpperCase();
                    });
               }
               if (removedWords != null) {
                    allRemoved = allRemoved.concat(removedWords).sort(function compare(a, b) {
                         a.toUpperCase() - b.toUpperCase();
                    });
               }

               // csInterface.evalScript('$.evalFile("' + rootPath + 'Text Tools/collectAbbreviation.jsx")', function (result) {
               try {
                    // let excistedAbbreList = result.split("#split#")[0].split(/[\s ]*;[\s ]*/);
                    // let mainAbbreList = result.split("#split#")[1].split(", ");
                    // let suspectList = result.split("#split#")[2].split(", ");
                    let excistedAbbreList = "";
                    let mainAbbreList = allEngWords;
                    let suspectList = removedWords;
                    // alert(abbreList);
                    // alert(abbreList.length);
                    let excistedAbbreList_ObjedAry = [];
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
                    // excistedAbbreList_ObjedAry.sort(function compare(a, b) {
                    //      a.abbre.toUpperCase() - b.abbre.toUpperCase();
                    // });
                    //
                    let mainAbbreList_filtered = [];
                    for (let i = 0; i < mainAbbreList.length; i++) {
                         if (mainAbbreList_filtered.indexOf(mainAbbreList[i]) === -1) {
                              mainAbbreList_filtered.push(mainAbbreList[i]);
                         }
                    }
                    // mainAbbreList_filtered.sort(function compare(a, b) {
                    //      if (a.toLowerCase() > b.toLowerCase()) {
                    //           return 1;
                    //      } else {
                    //           return -1;
                    //      }
                    // });
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
                                        return obj.full.toLowerCase() == exsistedMatchObj.full.toLowerCase();
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
                    // mainAbbreList_matched.sort(function compare(a, b) {
                    //      if (a.toLowerCase() > b.toLowerCase()) {
                    //           return 1;
                    //      } else {
                    //           return -1;
                    //      }
                    // });
                    // unusedRefedList.sort(function compare(a, b) {
                    //      if (a.toLowerCase() > b.toLowerCase()) {
                    //           return 1;
                    //      } else {
                    //           return -1;
                    //      }
                    // });
                    // //

                    // suspectList_filtered.sort(function compare(a, b) {
                    //      if (a.toLowerCase() > b.toLowerCase()) {
                    //           return 1;
                    //      } else {
                    //           return -1;
                    //      }
                    // });

                    let arraysToSort = [
                         mainAbbreList_matched,
                         unusedRefedList,
                         suspectList_filtered,
                         unmatchedList,
                         databaseRefedList,
                         newAbbreToUpdateAry,
                    ];

                    arraysToSort.forEach((element) => {
                         element.sort(function compare(a, b) {
                              if (a.toLowerCase() > b.toLowerCase()) {
                                   return 1;
                              } else {
                                   return -1;
                              }
                         });
                    });
                    let mergedMatchedList = "'" + mainAbbreList_matched.join("; ").toString() + "'";
                    let mergedUnmatchedList = "'" + unmatchedList.join("; ").toString() + "'";
                    let mergedUnusedRefedList = "'" + unusedRefedList.join("; ").toString() + "'";
                    let mergedDatabaseRefedList = "'" + databaseRefedList.join("; ").toString() + "'";
                    let mergedSuspectList = "'" + suspectList_filtered.join("; ").toString() + "'";
                    let newAbbreToUpdateList = "'" + newAbbreToUpdateAry.join("; ").toString() + "'";
                    let test = "AA";
                    document.getElementById("outcome").innerText = mergedMatchedList;
                    // csInterface.evalScript(
                    //      "var colletedOutcome = " +
                    //           mergedMatchedList +
                    //           "; var unmatcheddAbbre = " +
                    //           mergedUnmatchedList +
                    //           "; var unusedAbbre = " +
                    //           mergedUnusedRefedList +
                    //           "; var databasedAbbre = " +
                    //           mergedDatabaseRefedList +
                    //           "; var suspectAbbre = " +
                    //           mergedDatabaseRefedList +
                    //           "; var suspectAbbre = " +
                    //           mergedSuspectList +
                    //           "; var newAbbre = " +
                    //           newAbbreToUpdateList +
                    //           ';$.evalFile("' +
                    //           rootPath +
                    //           'Text Tools/saveCollectAbbreviation.jsx")'
                    // );
                    // alert(abbreCsvFilePath);
                    if (newAbbreToUpdateAry.length > 0) {
                         csInterface.evalScript(
                              "var newAbbre = " +
                                   newAbbreToUpdateList +
                                   ';$.evalFile("' +
                                   fullUrl.toString().replace("file:///", "").replace("main.js", "") +
                                   'database/updateAbbreDatabase.jsx")'
                         );
                    }
               } catch (err) {
                    alert(err.line + "\n" + err, err.line);
               }
               // });
          } catch (err) {
               alert(err, err.line);
          }
     }
     $("button").tooltip("hide");
}
