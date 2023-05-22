try {
    var abbreTextFile = File(File($.fileName).path + "/abbreListDatabase.csv")
    //     var abbreTextFile = File("G:/共用雲端硬碟/6_資源/17 插件/Script Development/Extension/oisee.sctipts'/js/database/abbreListDatabase.csv");
    var date = new Date()
    var fromatedDate = date.getFullYear() + "/" + (date.getMonth() + 1) + "/" + date.getDate()
    newAbbre = newAbbre.replace(/ = /g, ',').replace(/; /g, "," + fromatedDate + '\n');
    abbreTextFile.open('a');
    abbreTextFile.encoding = 'BIG5';
    abbreTextFile.write("\n" + newAbbre + "," + fromatedDate)
    abbreTextFile.close();
} catch (err) {
    alert(err, err.line)
}


//     abbreTextFile.open("w");
//     //TODO:確定UTF-8紀錄簡體字不會變亂碼
//     abbreTextFile.encoding = "UTF-8";

//     var saveString = colletedOutcome + "\r\r縮寫表清單:\r" + colletedOutcome.replace(/; /g, "\r") + "\r\r"
//     if (unmatcheddAbbre != "") {
//         saveString += "\r\r待補上的縮寫:\r" + unmatcheddAbbre.replace(/; /g, "\r")
//     }
//     if (unusedAbbre != "") {
//         saveString += "\r\r未使用的縮寫:\r" + unusedAbbre.replace(/; /g, "\r")
//     }
//     if (databasedAbbre != "") {
//         saveString += "\r\r\r\r取自資料庫的縮寫:\r" + databasedAbbre.replace(/; /g, "\r")
//     }
//     if (suspectAbbre != "") {
//         saveString += "\r\r\r\r疑似縮寫字:\r" + suspectAbbre
//     }
//     if (newAbbre != "") {
//         saveString += "\r\r\r\r資料庫新增:\r" + newAbbre
//     }

//     abbreTextFile.write(saveString);
//     abbreTextFile.close();
//     abbreTextFile.execute();
// } catch (err) {
//     alert(err, err.line)
// }