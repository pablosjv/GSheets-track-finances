function getBanks() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Connection");

  mainSheet.getRange("J1:J1000").clear();

  var country = mainSheet.getRange("B31").getValue();
  var token = mainSheet.getRange("B24").getValue();

  var url = "https://ob.nordigen.com/api/aspsps/?country="+country;
  var headers = {
             "headers":{"accept": "application/json",
                        "Authorization": "Token " + token}
             };

  var response = UrlFetchApp.fetch(url, headers);
  var json = response.getContentText();
  var data = JSON.parse(json);

  for (var i in data) {
  mainSheet.getRange(Number(i)+1,10).setValue([data[i].name]);
  }
  
}

function createLink() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Connection");

  var bank = mainSheet.getRange("B40").getValue();
  var country = mainSheet.getRange("B31").getValue();
  var token = mainSheet.getRange("B24").getValue();

  var url = "https://ob.nordigen.com/api/aspsps/?country="+country;
  var headers = {
             "headers":{"accept": "application/json",
                        "Authorization": "Token " + token}
             };

  var response = UrlFetchApp.fetch(url, headers);
  var json = response.getContentText();
  var data = JSON.parse(json);

  for (var j in data) {
    if (data[j].name == bank) {
      var aspsp_id = data[j].id;
    }
  }

  var random_number = Math.random() * 100000000

  var myHeaders = {"accept": "application/json",
                   "Content-Type": "application/json",
                   "Authorization": "Token " + token
                   }

  var SS = SpreadsheetApp.getActiveSpreadsheet();
  var ss = SS.getActiveSheet();
  var redirect_link = '';
  redirect_link += SS.getUrl();
  redirect_link += '#gid=';
  redirect_link += ss.getSheetId(); 

  var raw = JSON.stringify({"redirect":redirect_link,"reference":random_number,"enduser_id":random_number});
  var type = "application/json";

  var requestOptions = {
    'method': 'POST',
    'headers': myHeaders,
    'payload': raw
  };

  var response = UrlFetchApp.fetch("https://ob.nordigen.com/api/requisitions/", requestOptions);
  var json = response.getContentText();
  var requisition_id = JSON.parse(json).id;

  var myHeaders = {"accept": "application/json",
                   "Content-Type": "application/json",
                   "Authorization": "Token " + token}

  var raw = JSON.stringify({"aspsp_id":aspsp_id});
  var type = "application/json";

  var requestOptions = {
    'method': 'POST',
    'headers': myHeaders,
    'payload': raw
  };

  var response = UrlFetchApp.fetch("https://ob.nordigen.com/api/requisitions/" + requisition_id + "/links/", requestOptions);
  var json = response.getContentText();

  var link = JSON.parse(json).initiate;

  mainSheet.getRange(50,2).setValue([link]);
  mainSheet.getRange(1,12).setValue([requisition_id]);
  
}

function getTransactions() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("Connection");
  var transactionsSheet = ss.getSheetByName("Transactions");

  transactionsSheet.getRange("A2:A1000").clearContent();
  transactionsSheet.getRange("B2:B1000").clearContent();
  transactionsSheet.getRange("C2:C1000").clearContent();

  var token = mainSheet.getRange("B24").getValue();
  var requisition_id = mainSheet.getRange("L1").getValue();

  var url = "https://ob.nordigen.com/api/requisitions/" + requisition_id + "/";
  var headers = {
             "headers":{"accept": "application/json",
                        "Authorization": "Token " + token}
             };

  var response = UrlFetchApp.fetch(url, headers);
  var json = response.getContentText();
  var accounts = JSON.parse(json).accounts;
  
  row_counter = 2

  for (var i in accounts) {

      var account_id = accounts[i]

      var url = "https://ob.nordigen.com/api/accounts/" + account_id + "/transactions/";
      var headers = {
                "headers":{"accept": "application/json",
                            "Authorization": "Token " + token}
                };

      var response = UrlFetchApp.fetch(url, headers);
      var json = response.getContentText();
      var transactions = JSON.parse(json).transactions.booked;

      for (var i in transactions) {

        transactionsSheet.getRange(row_counter,1).setValue([transactions[i].valueDate]);
        transactionsSheet.getRange(row_counter,2).setValue([transactions[i].remittanceInformationUnstructured]);
        transactionsSheet.getRange(row_counter,3).setValue([transactions[i].transactionAmount.amount]);

        row_counter += 1

      }

  }

}