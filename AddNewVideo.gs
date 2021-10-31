function AddNewVideo() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastvideoid = sheet.getRange('Result!A11').getValue()
  var lastrow = sheet.getRange('Result!B11').getValue()

  spreadsheet.getRange('Database!A1')
  .setFormula('=query(importjson("https://www.vlive.tv/globalv-web/vam-web/post/v1.0/board-3484/posts?appId=8c6cc7b45d2568fb668be6e05b6e5a3b&fields=author,createdAt,officialVideo,thumbnail,title,url&sortType=LATEST&limit=100000&gcc=KR&locale=ko_KR","/data/"),"select * limit 6")');
  // Database!A1에 함수 입력
  
  Utilities.sleep(6000);
  // 불러오는데 시간걸릴까봐 슬립

  if(sheet.getRange('Result!C12').getValue() === "YES") { // VideoID칸이 숫자+6자리 체크가 YES이면
    for (var i=6; i>1; i--){ // i가 6부터 1씩 줄어들면서 2까지 반복
      if(spreadsheet.getSheetByName('Result').getRange(i+9,11).getValue() !== "LIVE") { // 라이브중인거는 제외
        if(lastvideoid < spreadsheet.getSheetByName('Result').getRange(i+9,10).getValue()){ // 불러온 비디오가 브이앱목록의 Last VideoID보다 크면
          spreadsheet.getSheetByName('VLIVE').insertRowsAfter(lastrow, 1); // LastRow 밑에 1행 추가
          var lastrow = lastrow + 1 // 브이앱목록 마지막행 번호에 +1 (그래야 그 행에 새로운 비디오를 추가하니까)
          spreadsheet.getSheetByName('Result').getRange(i + ":" + i).copyTo(spreadsheet.getSheetByName('VLIVE').getRange(lastrow,1), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false) // 새로운 비디오를 목록에 추가
          Utilities.sleep(1000); // 너무 빨라서 오류날수있음
          spreadsheet.getSheetByName('VLIVE').getRange(lastrow,4).setRichTextValue(SpreadsheetApp.newRichTextValue()
          .setText(spreadsheet.getSheetByName('Result').getRange(i,4).getValue())
          .setTextStyle(0, 33, SpreadsheetApp.newTextStyle()
          .setForegroundColor('#1155cc')
          .setUnderline(true)
          .build())
          .build()); // 링크에 하이퍼링크 설정
          Utilities.sleep(1000); // 너무 빨라서 오류날수있음
          spreadsheet.getSheetByName('Result').getRange(i+9,10).copyTo(spreadsheet.getRange('Result!A11'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false)
          // Last VideoID 갱신
          Utilities.sleep(1000); // 너무 빨라서 오류날수있음
          spreadsheet.getRange('Result!B11').setValue(lastrow); // lastrow 갱신
          Utilities.sleep(1000); // 너무 빨라서 오류날수있음
          console.log(lastrow); // 잘 진행되는지 확인용
        }
     }
    }
  }
  Utilities.sleep(5000); // 너무 빨라서 오류날수있음
  spreadsheet.getRange('Database!A1').clear({contentsOnly: true, skipFilteredRows: true}); // Database!A1에 함수 제거

  spreadsheet.getRange('Result!D12').setValue('=now()'); // now함수 입력
  spreadsheet.getRange('Result!D12').copyTo(spreadsheet.getRange('Result!D11'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false) // now함수 결과값 복붙
  spreadsheet.getRange('Result!D12').copyTo(spreadsheet.getRange('Readme!A13'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false) // now함수 결과값 복붙
  spreadsheet.getRange('Result!D12').clear({contentsOnly: true, skipFilteredRows: true}); // now함수 제거
}