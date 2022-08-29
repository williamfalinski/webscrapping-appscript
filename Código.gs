function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Menu Investimentos')
      .addItem('Atualizar cotações FIIS', 'etlFII')
      .addItem('Atualizar cotações Ações', 'TODO')
      .addToUi();
}


function etlFII(){
  //Executando extracao dos dados de FIIS
  //SHEET_NAME
  const ss = SpreadsheetApp.getActive().getSheetByName("FIIS");
  //RANGE dos codigos de FIIs
  var range = ss.getRange("D2:D")
  //Recuperando valores da tabela
  var values = range.getValues()
  let row = 2;
  for (const fii of values) {
    if (fii != '') {
      console.log("FII: " + fii + "Row: " + row)
      try {
        data = extractDataFII(fii)
        //Salvando o VALOR
        SpreadsheetApp.getActiveSheet().getRange('E'+row).setValue(data[0].replace(/[^\.,\d]/g, ''));
        //Salvando o P/VP
        SpreadsheetApp.getActiveSheet().getRange('F'+row).setValue(data[1].replace(/[^\.,\d]/g, ''));
        //Salvando o DY 12 meses
        SpreadsheetApp.getActiveSheet().getRange('G'+row).setValue(parseFloat(data[2].replace(/[^\.,\d]/g, '').replace(',', '.'))/100);
        //Salvando o DIVIDENDOS
        SpreadsheetApp.getActiveSheet().getRange('H'+row).setValue(data[3].replace(/[^\.,\d]/g, ''));
        Utilities.sleep(2000)
      }catch (e) {
        Logger.log(e);
        SpreadsheetApp.getUi().alert('Erro ao buscar dados de: '+fii+'\n'+e);
      }
    }
    row++;
  }
}

function extractDataFII(fii) {
  var url = 'https://investidor10.com.br/fiis/'+fii+'/';

  var websiteContent = UrlFetchApp.fetch(url).getContentText();

  //Extract com RegEx
  var valueRegExp = new RegExp(/Cotação<\/span>\n<\/div>\n<\/div>\n<div class="_card-body">\n<div>\n<span>(.+)<\/span>/m); 
  var pvpRegExp = new RegExp(/<span title="P\/VP">P\/VP<\/span>\n<\/div>\n<\/div>\n<div class="_card-body">\n<span>(.+)<\/span>/m); 
  var dy12RegExp = new RegExp(/DY \(12M\)<\/span>\n<\/div>\n<\/div>\n<div class="_card-body">\n<div>\n<span>(.+)<\/span>/m); 
  var dyieldRegExp = new RegExp(/ÚLTIMO RENDIMENTO\n<\/span>\n<div class="value">\n<span>\n(.+)\n<\/span>/m);
  
  var value = valueRegExp.exec(websiteContent);
  var pvp = pvpRegExp.exec(websiteContent);
  var dy12 = dy12RegExp.exec(websiteContent);
  var dyield = dyieldRegExp.exec(websiteContent);

  // Logger.log('VALOR: ' + value[1]);
  // Logger.log('P/VP: ' + pvp[1]);
  // Logger.log('DY(12): ' + dy12[1]);
  // Logger.log('DY: ' + dyield[1]);

  return [value[1], pvp[1], dy12[1], dyield[1]]
}
