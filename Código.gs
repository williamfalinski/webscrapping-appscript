function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Menu Investimentos')
      .addItem('Atualizar cotações Ações', 'etlAcoes')
      .addToUi();
}

function etlAcoes(){
  const error = []
  const expts = []
  //Executando extracao dos dados das acoes
  //SHEET_NAME
  const ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SHEET_NAME");
  //RANGE dos codigos de Acoes
  var range = ss.getRange("C2:C")
  //Recuperando valores da tabela
  var values = range.getValues()
  let row = 2;
  for (const tckt of values) {
    if (tckt != '') {
      console.log("Ação: " + tckt + " Row: " + row)
      try {
        data = extractDataTicket(tckt+'.SA')
        //Salvando o Valor
        // SpreadsheetApp.getActiveSheet().getRange('D'+row).setValue(data[0].replace(/[^\.,\d]/g, ''));
        data[0] !== null ? ss.getRange('D'+row).setValue(data[0].replace('.', ',')) : null;
        //Salvando o P/VP
        data[1] !== null ? ss.getRange('E'+row).setValue(data[1].replace('.', ',')) : null;
        //Salvando o P/L
        data[2] !== null ? ss.getRange('F'+row).setValue(data[2].replace('.', ',')) : null;
        //Salvando o VPA
        data[3] !== null ? ss.getRange('H'+row).setValue(data[3].replace('.', ',')) : null;
        //Salvando o LPA
        data[4] !== null ? ss.getRange('I'+row).setValue(data[4].replace('.', ',')) : null;
        //Salvando o DY
        data[5] !== null ? ss.getRange('Q'+row).setValue(data[5].replace('.', ',')+'%') : null;
      }catch (e) {
        error.push(tckt);
        expts.push(e);
      }
    }
    row++;
    Utilities.sleep(50);
  }
  if (error.length!==0){
    SpreadsheetApp.getUi().alert('Erro ao buscar dados de: '+error+'\n'+expts);
  }
}

function extractDataTicket(ticket) {
  var options = {
    headers: {
    'User-Agent'      : 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36', 
    'Accept'          : 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9', 
    'Accept-Language' : 'en-US,en;q=0.5',
    'DNT'             : '1',
    'Connection'      : 'close'
    }
  };
  var url = 'https://finance.yahoo.com/quote/'+ticket+'/key-statistics?p='+ticket;
  console.log(url)

  var websiteContent = UrlFetchApp.fetch(url, options).getContentText();

  //RegEx Value
  var valueRegExp = new RegExp(/<fin-streamer class="Fw\(b\) Fz\(36px\) Mb\(-4px\) D\(ib\)" data-symbol="(.+)" data-test="qsp-price" data-field="regularMarketPrice" data-trend="none" data-pricehint="2" value="([0-9]+.?[0-9]+)" active="">([0-9]+.?[0-9]+)/m); 

  //RegEx P/VPA
  var pvpRegExp = new RegExp(/<span>Price\/Book<\/span> <!-- -->\(mrq\)<sup aria-label=""><\/sup><div class="W\(3px\) Pos\(a\) Start\(100%\) T\(0\) H\(100%\) Bg\(\$pfColumnFakeShadowGradient\) Pe\(n\) Pend\(5px\)"><\/div><\/td><td class="Ta\(c\) Pstart\(10px\) Miw\(60px\) Miw\(80px\)--pnclg Bgc\(\$lv1BgColor\) fi-row:h_Bgc\(\$hoverBgColor\)">([0-9]+.?[0-9]+)/m); 

  //RegEx P/E
  var plRegExp = new RegExp(/<span>Trailing P\/E<\/span> <sup aria-label=""><\/sup><div class="W\(3px\) Pos\(a\) Start\(100%\) T\(0\) H\(100%\) Bg\(\$pfColumnFakeShadowGradient\) Pe\(n\) Pend\(5px\)"><\/div><\/td><td class="Ta\(c\) Pstart\(10px\) Miw\(60px\) Miw\(80px\)--pnclg Bgc\(\$lv1BgColor\) fi-row:h_Bgc\(\$hoverBgColor\)">([0-9]+.?[0-9]+)/m); 

  //RegEx Trailing DY 12
  var dy12RegExp = new RegExp(/<span>Trailing Annual Dividend Yield<\/span> <sup aria-label="Data derived from multiple sources or calculated by Yahoo Finance\.">3<\/sup><\/td><td class="Fw\(500\) Ta\(end\) Pstart\(10px\) Miw\(60px\)">([0-9]+.?[0-9]+)/m); 

  //RegEx VPA
  var vpaRegExp = new RegExp(/<span>Book Value Per Share<\/span> <!-- -->\(mrq\)<sup aria-label=""><\/sup><\/td><td class="Fw\(500\) Ta\(end\) Pstart\(10px\) Miw\(60px\)">([0-9]+.?[0-9]+)/m); 

  //RegEx LPA
  var lpaRegExp = new RegExp(/<span>Diluted EPS<\/span> <!-- -->\(ttm\)<sup aria-label=""><\/sup><\/td><td class="Fw\(500\) Ta\(end\) Pstart\(10px\) Miw\(60px\)">([0-9]+.?[0-9]+)/m); 

  
  var value = valueRegExp.exec(websiteContent);
  var pvp = pvpRegExp.exec(websiteContent);
  var pl = plRegExp.exec(websiteContent);
  var dy12 = dy12RegExp.exec(websiteContent);
  var vpa = vpaRegExp.exec(websiteContent);
  var lpa = lpaRegExp.exec(websiteContent);

  // Logger.log('Value: ' + value[2]);
  // Logger.log('P/VPA: ' + pvp[1]);
  // Logger.log('P/L: ' + pl[1]);
  // Logger.log('Trailing DY(12): ' + dy12[1]);
  // Logger.log('VPA: ' + vpa[1]);
  // Logger.log('LPA: ' + lpa[1]);

  // Logger.log( [
  //   value !== null ? value[2] : null, 
  //   pvp !== null ? pvp[1] : null, 
  //   pl !== null ? pl[1] : null,
  //   vpa !== null ? vpa[1] : null, 
  //   lpa !== null ? lpa[1] : null,
  //   dy12 !== null ? dy12[1] : null,
  //   ]);

  return [
    value !== null ? value[2] : null, 
    pvp !== null ? pvp[1] : null, 
    pl !== null ? pl[1] : null,
    vpa !== null ? vpa[1] : null, 
    lpa !== null ? lpa[1] : null,
    dy12 !== null ? dy12[1] : null,
    ]
}
