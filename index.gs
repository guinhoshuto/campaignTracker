var planilhaDestino = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1zDhXytK7l15-CJV4p4X5uAFLok4ajIFtf-1htb6cxAQ/edit#gid=0");
var token = "EAAeuwgLZBQfkBAMNbk5PKiflzKaVKrxBoJfWMg7apxnrn9bbZA5jhcK1b5uZCoPLmhvZBSbOajDBGSY0j19ZAzutfD1vIABianqHkfLksX9FpEQdRwALilHGrjHEKYmnZA8MzhAA3wM8AlDvbKcrPm7XMgbbZAvHfwZD";

function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('(=ↀωↀ=)')
      .addItem('Atualiza', 'configuraPlanilha')
      .addSeparator()
      

      .addToUi();
      

}

function configuraPlanilha(){
    var abaInteresses = false;
    var abaJob = false;
    var abas = planilhaDestino.getSheets();

    var config = planilhaDestino.getSheetByName('config');
    var numRows = config.getLastRow(); 
    var actIdsRange = onfig.getRange('A2:D' + numRows);
    var actIds = actIdsRange.getValues();
    var config = planilhaDestino.getSheetByName('config');
      


    Logger.log(numRows);
    for ( var i = 0 ; i < actIds.length ; i++ ){
        Logger.log(actIds[i]);
        if(planilhaDestino.getSheetByName(actIds[i][1])){

          preencheAcompanhamento(actIds[i][0],actIds[i][1],actIds[i][2],actIds[i][3]);
          Logger.log("entrou"); 
        } else {
          planilhaDestino.insertSheet().setName(actIds[i][1]);
          preencheAcompanhamento(actIds[i][0],actIds[i][1],actIds[i][2],actIds[i][3]);
        }
    }

    
    
    function preencheAcompanhamento(act_id,conta,mensal,threshold) {
    
      // var transactionsUrl = "https://graph.facebook.com/v3.0/" + act_id + "/transactions?access_token=" + token + "&time_start=1540267200&time_stop=1549594799";
      // var transactions = JSON.parse(UrlFetchApp.fetch(transactionsUrl)).data[0];
      // Logger.log(transactions);
      
    
      var accountInsightsUrl = "https://graph.facebook.com/v3.2/act_" + act_id + "/insights?fields=reach%2Cspend%2Cimpressions%2Ccpm&date_preset=this_month&access_token=" + token;
      var accountInsights = JSON.parse(UrlFetchApp.fetch(accountInsightsUrl)).data[0];
    
      var search = "https://graph.facebook.com/v3.2/act_" + act_id + "?fields=campaigns%7Bname%2Clifetime_budget%2Cadsets%7Bname%2Clifetime_budget%2Cinsights%7Bspend%2Creach%7D%7D%7D%2Cbalance&access_token=" + token;
      // var search = "https://graph.facebook.com/v3.2/act_" + act_id + "/campaigns?fields=name%2Clifetime_budget%2Cadsets%7Bname%2Clifetime_budget%2Cinsights%7Breach%2Cspend%7D%7D&access_token=" + token;
      Logger.log(search);
      var sheet = planilhaDestino.getSheetByName(conta);
      sheet.clear();
       
      var charts = sheet.getCharts();
      if(charts.length != 0){
        sheet.removeChart(charts[0]);
        sheet.removeChart(charts[1]);
      }
      
      Logger.log(search);
      var updated = new Date();
      
      var accountData = JSON.parse(UrlFetchApp.fetch(search));
      var resposta = accountData.campaigns;
      
      var gasto = accountInsights.spend.replace(".",",");
      mensal -= accountInsights.spend;

      var restante = mensal.toString().replace(".",",");
      Logger.log(restante + " oi");
      
      threshold -= accountData.balance/100;
      var restanteCobranca = threshold.toString().replace(".",",");
      
      sheet.appendRow(["Alcance",accountInsights.reach]);
      sheet.appendRow(["Impressões",accountInsights.impressions]);
      sheet.appendRow(["Gasto",gasto,restante]); 
      sheet.appendRow(["Cobrança",accountData.balance/100,restanteCobranca]);
      
      sheet.appendRow(["---"]);
      sheet.appendRow(["---"]);
      sheet.appendRow(["---"]);
      sheet.appendRow(["---"]);
      
      var cobranca = sheet.newChart();
      cobranca.addRange(sheet.getRange("A4:C4"))
          .setChartType(Charts.ChartType.BAR)
          .setPosition(2, 2, 650, 430)
          .setOption('isStacked', 'percent') 
          .setOption('legend', '{position: "top", maxLines: 3}')
          .setOption('hAxis', 'minValue: 0,ticks: [0, .3, .6, .9, 1]')
          .setOption('title', 'Cobrança');
      sheet.insertChart(cobranca.build());
      
      var chart = sheet.newChart();
      chart.addRange(sheet.getRange("A3:C3"))
          .setChartType(Charts.ChartType.BAR)
          .setPosition(2, 2, 650, 0)
          .setOption('isStacked', 'percent') 
          .setOption('legend', '{position: "top", maxLines: 3}')
          .setOption('hAxis', 'minValue: 0,ticks: [0, .3, .6, .9, 1]')
          .setOption('title', 'Gasto');
      sheet.insertChart(chart.build());

      var options_fullStacked = {
        isStacked: 'percent',
        height: 300,
        legend: {position: 'top', maxLines: 3},
        hAxis: {
          minValue: 0,
          ticks: [0, .3, .6, .9, 1]
        }
      };


      for(var i = 0; i < resposta.data.length; i++){
        sheet.appendRow([resposta.data[i].name]);
        for(var j = 0; j < resposta.data[i].adsets.data.length; j++){
          
          
          if(typeof resposta.data[i].adsets.data[j].insights !== 'undefined'){
            
          }
        
          Logger.log(sheet.appendRow(["",
                                      resposta.data[i].adsets.data[j].name,
                                      "R$ " + resposta.data[i].adsets.data[j].lifetime_budget/100,

                                      
                                     ])
          );
        }
      }

//      for(var i = 0; i < resultados.length; i++){
//        sheet.appendRow([resultados[i].id, resultados[i].name,resultados[i].type,resultados[i].audience_size,updated,query]);
//      }
    }
    
    // if(!abaInteresses) {
      
    //   var abaInteressesNova = planilhaDestino.insertSheet();
    //   abaInteressesNova.setName("Busca");
    //   Logger.log(abaInteressesNova.appendRow(["ID","Nome","Tipo","Tamanho","Atualizado em","Busca"]));
    // }
    
    // if(!abaJob) {
      
    //   var abaJobNova = planilhaDestino.insertSheet();
    //   abaJobNova.setName("Jobtitle");
    //   Logger.log(abaJobNova.appendRow(["ID","Nome","Tipo","Tamanho","Atualizado em","Busca"]));
    // }
  }
  
