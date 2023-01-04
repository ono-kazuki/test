// function myFunction() {
//   let num;
//   num = 10;
//   console.log(num);
  
//   const x = 10, y = 5;
//   console.log(`${x}と${y}の和は${x + y}です。`);
  
//   const msg = 'Hello GAS!';
//   console.log(`GASによるメッセージ: ${msg}`);

//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   console.log(ss.getName());

//   const sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('シート1');
//   console.log(sheet1.getName());

//   // const range = sheet.getRange('A3:C4');
//   // console.log(range.getValues());

  
//   // let i = 2;
//   // while(sheet.getRange(i, 4).getValue()) {
//   //   i++;
//   // }

//   // console.log(sheet.getRange(i, 1).getValue());
//   // sheet.getRange(i, 4).setValue(true);
//   const sheet = SpreadsheetApp.getActiveSheet(); 
//   const lastRow = sheet.getLastRow();
//   for(let i = 2; i <= lastRow; i++) {
//     if(!sheet.getRange(i, 4).getValue()){ 
//       console.log(sheet.getRange(i, 1).getValue());
//       sheet.getRange(i, 4).setValue(true);
      
//       if(i >= lastRow) {
//         console.log(sheet.getRange(2, 4, lastRow - 1, 1).getA1Notation());
//         sheet.getRange(2, 4, lastRow - 1).clearContent();
//       }
//     }
//   }
// }

function myFunction() {
  const sheet = SpreadsheetApp.getActiveSheet(); 
  const lastRow = sheet.getLastRow();
  
  for(let i = 2; i <= lastRow; i++) {
    if(!sheet.getRange(i, 4).getValue()){ 
      const body = sheet.getRange(i, 1).getValue();
      sendMessage(body);
      sheet.getRange(i, 4).setValue(true);
      
      if(i >= lastRow) {
        sheet.getRange(2, 4, lastRow - 1).clearContent();
      }
      break;
    }
  }
}


function sendMessage(body){
  const cw = ChatWorkClient.factory({token: 'e34543f49aad3be1384e42c0edc3815e'});
  cw.sendMessageToMyChat(body);
}

function setScriptProperty() {
  PropertiesService.getScriptProperties().setProperty('CW_TOKEN', 'e34543f49aad3be1384e42c0edc3815e');
}
