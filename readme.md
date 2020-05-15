Activity-Sheet Manager
========
Powered by Google Spreadsheet and Google Apps script

Objects
-----------
1. 고등학교 내 탐구 활동 신청서(탐활서) 의 체계적인 관리가 필요
2. 과구리 (jshsus.kr)을 통해 접수된 탐활서 명단을 Spreadsheet로 정리

What will be in the spreadsheet?
--------------
1. Firebase (database) 에 저장된 당일 신청 명단 json data를 가져옴
2. 탐활서가 보여질 층에 따라 분리됨 (1, 2, 3 시트)
3. 탐활서에 있는 면학 시간, 장소, 내용, 전체학생, 지도교사가 한 행에 추가됨
4. 4, 5, 6 시트에는 해당 층에서 탐활서가 존재하는 학생들의 활동 장소와 시간을 정렬해 표시함.

How does it work?
------------
Google Apps Script에는 trigger라는 기능이 존재합니다. 예약된 시간마다 Script내 함수를 실행할 수 있도록 돕는 기능입니다. <br>
전체 탐활서를 업데이트하는 init함수는 10분마다 실행됩니다. 그리고, 날짜가 바뀐 뒤 init함수가 실행되면, 전날 존재했던 탐활서를 전부 지워버립니다.
<br> 그날 탐활서를 Google drive에 pdf로 저장하는 exportAllSheetsAsSeparatePDFs 함수  또한 매일 18시, 23시마다 실행됩니다. 

Screenshot
------------
![Test Image](https://raw.githubusercontent.com/TriangleYJ/activity-sheet-manager/master/sample.PNG)

TODO
----------
* 해당 코드는 2019년도에 작성되었습니다. 따라서 해당 층에 있는 학년들이 달라질 경우 checkFloor 함수 수정이 필요합니다. (현재 : 3학년 - 2층, 2학년 - 3층, 1학년 - 4층)
