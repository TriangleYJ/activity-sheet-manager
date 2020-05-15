var d = new Date();
var yy = d.getFullYear();
var dd = d.getDate();
var mm = d.getMonth() + 1;


function setFloorData(floor0) {
    var time = 1;
    var place = 2;
    var content = 3;
    var students = 4;
    var teacher = 5;

    var config = {
        Key: "/Api Key/",
        databaseURL: "https://jshsus-6144b.firebaseio.com",
    };
    database = FirebaseApp.getDatabaseByUrl(config.databaseURL, config.Key);
    var data = database.getData("ssam/club");


    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[floor0 - 2]
    var sheet_pep = SpreadsheetApp.getActiveSpreadsheet().getSheets()[floor0 + 1]
    today = '최종 탐구 활동 신청서 명단 (' + mm + '월 ' + dd + '일) (' + floor0 + '층)';


    sheet.getRange(1, 2).setValue(today);

    // 기존 데이터 지우기
    var start = 3;
    var lcount = sheet.getRange(1, 5).getValue().split(" : ")[1];
    for (var i = 0; i <= lcount + 1; i++) {
        for (var j = 0; j < 5; j++) {
            sheet.getRange(i + 3, j + 1).setValue('');
        }
    }

    var pep = [];
    //새 데이터 집어 넣기 및 오래된 탐활서 삭제
    for (var key in data) {
        //Logger.log(data[key])
        var item = data[key];
        var da = item['date'].split("/");
        var floors = item['floor'].split(",");
        Logger.log(floors)
        if ((da[0] < yy) || (da[0] == yy && da[1] < mm) || (da[0] == yy && da[1] == mm && da[2] < dd)) { // 날짜 지난 거 지우기
            database.removeData("ssam/club/" + key); //<======= 잘못 만지면 과구리 전부 날라감. 조심하길
        } else if (da[0] == yy && da[1] == mm && da[2] == dd && floors.indexOf(floor0 + "") != -1 && item['teacher'] != "") { // 오늘자 && 층수 일치 && 선생님 이름 존재
            var time0 = item['time'];
            var place0 = item['place'];
            //항목들 작성
            sheet.getRange(start, time).setValue(item['print_time']);
            sheet.getRange(start, place).setValue(place0);
            sheet.getRange(start, content).setValue(item['desc']);
            //sheet.getRange(start, students).setValue(item['studentZ'] + ' ( 대표 학생 ) , ' + item['students']);
            sheet.getRange(start, teacher).setValue(item['teacher']);

            var stz = item['studentZ'];
            var sts = item['students'];
            var str = "";
            var nstz = rb(stz); //학번이름 of student Z
            var stsa = sts.split(",");


            // 시트 4,5,6에 들어갈 시간 정렬 부분 때문에 아래 코드가 복잡해짐.

            //탐활서.net에 학생 대표가 학생 이름에 중복으로 들어가든, 안들어가든 상관없이 한번만 뜨게 해주느라 같은 문단이 2번 들어감.

            if (sts != "") { // 대표 학생 외 나머지 학생들 시간 정리 코드
                stsa.forEach(function (i) {
                    nst = rb(i);
                    if (nst != nstz) { // 대표학생이 중복될 경우 -> 무시
                        str += ", " + sp(nst);

                        var pep_exist = false;
                        pep.forEach(function (j) {
                            if (ha(j[0]) == ha(nst)) { // 이미 학생이 있을 때
                                pep_exist = true;
                                if (time0.indexOf(5) != -1) j[1].push(place0);
                                if (time0.indexOf(6) != -1) j[2].push(place0);
                            }
                        });
                        if (!pep_exist && checkFloor(nst, floor0)) {//최초로 학생 등장할때
                            var time1 = [];
                            var time2 = [];
                            if (time0.indexOf(5) != -1) time1.push(place0);
                            if (time0.indexOf(6) != -1) time2.push(place0);
                            pep.push([nst, time1, time2]);
                        }

                    }

                });
            }

            var pep_exist = false;
            pep.forEach(function (j) { // 대표 학생 시간 정리 코드
                if (ha(j[0]) == ha(nstz)) {
                    pep_exist = true;
                    if (time0.indexOf(5) != -1) j[1].push(place0);
                    if (time0.indexOf(6) != -1) j[2].push(place0);
                }
            });
            if (!pep_exist && checkFloor(nstz, floor0)) { // 321
                var time1 = [];
                var time2 = [];
                if (time0.indexOf(5) != -1) time1.push(place0);
                if (time0.indexOf(6) != -1) time2.push(place0);
                pep.push([nstz, time1, time2]);
            }

            if (pep.length > 1) { // 학생 정렬
                pep.sort(function (a, b) {
                    return ha(a[0]) - ha(b[0]);
                });
            }

            str = sp(rb(stz)) + " ( 대표학생 ) " + str;

            sheet.getRange(start, students).setValue(str);
            start++;
        }

    }
    sstart = 3;

    while (true) { // 시간 정리 시트 작성
        if (sheet_pep.getRange(sstart, 1).getValue() == "") { // 지난거 지워버리기
            break;
        } else {
            sheet_pep.getRange(sstart, 1).setValue("");
            sheet_pep.getRange(sstart, 2).setValue("");
            sheet_pep.getRange(sstart, 3).setValue("");
        }
        sstart++;
    }

    sstart = 3; // 변수 재활용
    today = floor0 + '층 탐활서 명단 및 시간 (' + mm + '월 ' + dd + '일)';
    sheet_pep.getRange(1, 1).setValue(today);
    pep.forEach(function (j) { // 업데이트
        if (!(j[1].length == 0 && j[2].length == 0)) { // 아침 오후면학인 경우 추가하지 않음.
            sheet_pep.getRange(sstart, 1).setValue(j[0]);
            sheet_pep.getRange(sstart, 2).setValue(j[1].join(", "));
            sheet_pep.getRange(sstart, 3).setValue(j[2].join(", "));
            sstart++;
        }
    });


    //안내문
    sheet.getRange(start + 1, students).setValue('[업데이트 : ' + d.toLocaleString() + ']');
    sheet.getRange(start + 1, time).setValue('학생회 IT부');
    //sheet.getRange(start + 1, content).setValue('(n) : 아침 n면학\n [n] : 오후 n면학\n ⓝ : 저녁 n면학');

    count = (start - 3);
    sheet.getRange(1, 5).setValue('활동 수 : ' + count);

}

function init() {
    setFloorData(2);
    setFloorData(3);
    setFloorData(4);
}

function time_ck(s) { // not used
    var k = s.split(",");
    var arr = []
    for (var i in k) {
        switch (k[i]) {
            case "1":
                arr.push("(1)");
                break;
            case "2":
                arr.push("(2)");
                break;
            case "3":
                arr.push("[1]");
                break;
            case "4":
                arr.push("[2]");
                break;
            case "5":
                arr.push("①");
                break;
            case "6":
                arr.push("②");
                break;
            case "7":
                arr.push("③");
                break;
        }
    }
    return arr.join(", ");
}

function rb(s) { //  1234 홍길동 -> 1234홍길동
    return s.replace(/(\s*)/g, "")
}

function sp(s) { //3212홍길 -> 3212 홍길
    return s.substring(0, 4) + " " + s.substring(4, s.length);
}

function gr(s) { // 학년 가져오기
    return s.substring(0, 1) / 1;
}

function ha(s) { // 학번 가져오기
    return s.substring(0, 4) / 1;
}

function checkFloor(hakbunirum, floor){
    return (gr(hakbunirum) == 5 - floor); // 3학년 2층 2학년 3층, 1학년 4층 체계가 바뀌면 이 코드도 복잡하게 바뀌어 질 듯.
}




//pdf exporter
/**
 * @license MIT
 *
 * © 2019 xfanatical.com. All Rights Reserved.
 *
 */

function _getAsBlob(url, sheet, range) {
    var rangeParam = ''
    if (range) {
        rangeParam =
            '&r1=' + (range.getRow() - 1)
            + '&r2=' + range.getLastRow()
            + '&c1=' + (range.getColumn() - 1)
            + '&c2=' + range.getLastColumn()
    }
    var exportUrl = url.replace(/\/edit.*$/, '')
        + '/export?exportFormat=pdf&format=pdf'
        + '&size=A4'
        + '&portrait=false'
        + '&fitw=true'
        + '&top_margin=0.75'
        + '&bottom_margin=0.75'
        + '&left_margin=0.7'
        + '&right_margin=0.7'
        + '&sheetnames=false&printtitle=false'
        + '&pagenum=false'
        + '&gridlines=true'
        + '&fzr=FALSE'
        + '&gid=' + sheet.getSheetId()
        + rangeParam

    Logger.log('exportUrl=' + exportUrl)
    var response = UrlFetchApp.fetch(exportUrl, {
        headers: {
            Authorization: 'Bearer ' +  ScriptApp.getOAuthToken(),
        },
    })

    return response.getBlob()
}


function exportAllSheetsAsSeparatePDFs() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
    var files = []
    for(var i = 0; i < 3; i++){
        var sheet = spreadsheet.getSheets()[i]
        spreadsheet.setActiveSheet(sheet)
        var lcount = sheet.getRange(1, 5).getValue().split(" : ")[1];

        if(lcount != 0){ // 만약 그날 해당 층 탐활서가 없으면 pdf 로 기록이 안됨
            var blob = _getAsBlob(spreadsheet.getUrl(), sheet)
            var fileName = sheet.getName() + ' 탐활서 (' + (yy + "").substr(2,2) + '.' + mm + '.' + dd + '.)'
            blob = blob.setName(fileName)
            var dir = DriveApp.getFoldersByName("탐활서 기록").next();

            var children = dir.getFilesByName(fileName);
            if (children.hasNext()) {
                dir.removeFile(children.next());
                dir.createFile(blob);
            } else {
                dir.createFile(blob);
            }
        }
    }
}