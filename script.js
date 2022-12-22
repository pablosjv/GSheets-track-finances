const ss = SpreadsheetApp.getActiveSpreadsheet();
const connectionSheet = ss.getSheetByName("Connection");

const tokenSheetCell = "F25"
const refreshSheetCell = "G25"
const secretIdSheetCell = "B25"
const secretKeySheetCell = "C25"

const bankSheetCellRange = "J4:J100"

function getNordigenToken(secretId, secretKey) {
    var requestOptions = {
        'method': 'POST',
        'headers': {
            'Accept': 'application/json',
            'Content-Type': 'application/json'
        },
        'payload': JSON.stringify({
            "secret_id": secretId,
            "secret_key": secretKey
        }),
        'muteHttpExceptions': true
    };

    var response = UrlFetchApp.fetch("https://ob.nordigen.com/api/v2/token/new/", requestOptions);
    if (response.getResponseCode() != 200)
    {
        Logger.log(response.getResponseCode());
        Logger.log(response.getContentText());
        throw new Error('Login failed! Check the credentials');
    }
    var json = response.getContentText();

    return {
        'token': JSON.parse(json).access,
        'refresh': JSON.parse(json).refresh,
    };
}

function getNordigenTokenRefresh(refresh) {
    var requestOptions = {
        'method': 'POST',
        'headers': {
            'Accept': 'application/json',
            'Content-Type': 'application/json'
        },
        'payload': JSON.stringify({
            "refresh": refresh,
        }),
        'muteHttpExceptions': true
    };

    var response = UrlFetchApp.fetch("https://ob.nordigen.com/api/v2/token/refresh/", requestOptions);
    if (response.getResponseCode() != 200)
    {
        Logger.log(response.getResponseCode());
        throw new Error('Refresh failed! New token might need be issued');
    }
    var json = response.getContentText();

    return JSON.parse(json).access;
}

function getTokenSheet() {
    const secretId = connectionSheet.getRange(secretIdSheetCell).getValue();
    const secretKey = connectionSheet.getRange(secretKeySheetCell).getValue();
    Logger.log('Secret ID: ' + secretId)
    Logger.log('Secret Key: ' + secretKey)
    const { token, refresh } = getNordigenToken(secretId, secretKey);
    Logger.log(refresh);
    connectionSheet.getRange(tokenSheetCell).setValue(token);
    connectionSheet.getRange(refreshSheetCell).setValue(refresh);

    return {
        'token': token,
        'refresh': refresh,
    };
}

function refreshTokenSheet() {
    var token, refresh;
    refresh = connectionSheet.getRange(refreshSheetCell).getValue();
    try
    {
        token = getNordigenTokenRefresh(refresh);
    } catch (error)
    {
        ({ token, refresh } = getTokenSheet());
    }
    connectionSheet.getRange(tokenSheetCell).setValue(token);
    connectionSheet.getRange(refreshSheetCell).setValue(refresh);

    return {
        'token': token,
        'refresh': refresh,
    };
}

function requestNordigenWithSheet(url, method = 'get', payload = {}) {
    let token;

    token = connectionSheet.getRange(tokenSheetCell).getValue();
    Logger.log(token);

    if (token === '')
    {
        Logger.log("Empty token! Refreshing...");
        ({ token } = getTokenSheet());
    }

    var requestOptions = {
        'method': method,
        'headers': {
            'Accept': 'application/json',
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + token
        },
        // 'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
    };

    if (method.toUpperCase() != 'GET') {
      requestOptions.payload = JSON.stringify(payload);
    }

    Logger.log(requestOptions)

    var response = UrlFetchApp.fetch(url, requestOptions);
    if (response.getResponseCode() == 401)
    {
        Logger.log("Invalid token! Trying to refresh");
        ({ token } = refreshTokenSheet());
        requestOptions.headers.Authorization = 'Bearer ' + token
        response = UrlFetchApp.fetch(url, requestOptions);
        if (response.getResponseCode() != 200)
        {
            Logger.log(response.getResponseCode());
            Logger.log(response.getContentText());
            throw new Error('Response error');
        }
    }
    else if (response.getResponseCode() != 200)
    {
        Logger.log(response.getResponseCode());
        Logger.log(response.getContentText());
        throw new Error('Response error');
    }
    var json = response.getContentText();

    return JSON.parse(json);
}

function getBanks() {
    Logger.log('Start getBanks')

    connectionSheet.getRange(bankSheetCellRange).clear();

    var country = connectionSheet.getRange("B32").getValue();
    Logger.log(country)

    var data = requestNordigenWithSheet('https://ob.nordigen.com/api/v2/institutions/?country=' + country)

    for (var i in data)
    {
        connectionSheet.getRange(Number(i) + 1, 10).setValue([data[i].name]);
    }

}

function createLink() {

    var bank = connectionSheet.getRange("B40").getValue();
    var country = connectionSheet.getRange("B31").getValue();
    var token = connectionSheet.getRange("B24").getValue();

    var url = "https://ob.nordigen.com/api/aspsps/?country=" + country;
    var headers = {
        "headers": {
            "accept": "application/json",
            "Authorization": "Token " + token
        }
    };

    var response = UrlFetchApp.fetch(url, headers);
    var json = response.getContentText();
    var data = JSON.parse(json);

    for (var j in data)
    {
        if (data[j].name == bank)
        {
            var aspsp_id = data[j].id;
        }
    }

    var random_number = Math.random() * 100000000

    var myHeaders = {
        "accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": "Token " + token
    }

    var SS = SpreadsheetApp.getActiveSpreadsheet();
    var ss = SS.getActiveSheet();
    var redirect_link = '';
    redirect_link += SS.getUrl();
    redirect_link += '#gid=';
    redirect_link += ss.getSheetId();

    var raw = JSON.stringify({ "redirect": redirect_link, "reference": random_number, "enduser_id": random_number });
    var type = "application/json";

    var requestOptions = {
        'method': 'POST',
        'headers': myHeaders,
        'payload': raw
    };

    var response = UrlFetchApp.fetch("https://ob.nordigen.com/api/requisitions/", requestOptions);
    var json = response.getContentText();
    var requisition_id = JSON.parse(json).id;

    var myHeaders = {
        "accept": "application/json",
        "Content-Type": "application/json",
        "Authorization": "Token " + token
    }

    var raw = JSON.stringify({ "aspsp_id": aspsp_id });
    var type = "application/json";

    var requestOptions = {
        'method': 'POST',
        'headers': myHeaders,
        'payload': raw
    };

    var response = UrlFetchApp.fetch("https://ob.nordigen.com/api/requisitions/" + requisition_id + "/links/", requestOptions);
    var json = response.getContentText();

    var link = JSON.parse(json).initiate;

    connectionSheet.getRange(50, 2).setValue([link]);
    connectionSheet.getRange(1, 12).setValue([requisition_id]);

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
        "headers": {
            "accept": "application/json",
            "Authorization": "Token " + token
        }
    };

    var response = UrlFetchApp.fetch(url, headers);
    var json = response.getContentText();
    var accounts = JSON.parse(json).accounts;

    row_counter = 2

    for (var i in accounts)
    {

        var account_id = accounts[i]

        var url = "https://ob.nordigen.com/api/accounts/" + account_id + "/transactions/";
        var headers = {
            "headers": {
                "accept": "application/json",
                "Authorization": "Token " + token
            }
        };

        var response = UrlFetchApp.fetch(url, headers);
        var json = response.getContentText();
        var transactions = JSON.parse(json).transactions.booked;

        for (var i in transactions)
        {

            transactionsSheet.getRange(row_counter, 1).setValue([transactions[i].valueDate]);
            transactionsSheet.getRange(row_counter, 2).setValue([transactions[i].remittanceInformationUnstructured]);
            transactionsSheet.getRange(row_counter, 3).setValue([transactions[i].transactionAmount.amount]);

            // if(transactions[i].hasOwnProperty('debtorName')){
            //   transactionsSheet.getRange(row_counter,4).setValue([transactions[i].debtorName]);
            // }

            row_counter += 1

            // break

        }

    }
}
