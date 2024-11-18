function PrimeiraAutenticacao() {
  // Abrir a planilha atual
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Acessar a aba chamada "Dados"
  var sheet_dados = spreadsheet.getSheetByName("PrimeiraAuth");
  
  // Ler o valor na célula A1 da aba "Dados"
  var client_id = sheet_dados.getRange("A2").getValue();
  var client_secret = sheet_dados.getRange("B2").getValue();
  var code = sheet_dados.getRange("C2").getValue();

  // Codificar em Base64
  var encodedCredentials = Utilities.base64Encode(client_id + ":" + client_secret);

  // Definir a URL e os parâmetros para a requisição de refresh
  var apiUrl = "https://www.bling.com.br/Api/v3/oauth/token";
  
  var headers = {
    "Authorization" : "Basic "+ encodedCredentials,
     "Content-Type": "application/x-www-form-urlencoded"
  };
  var payload = "grant_type=authorization_code&code=" + code;

  // Opções da requisição
  var options = {
    "method": "post",
    "headers": headers,
    "payload": payload,
    "muteHttpExceptions": true
  };

  // Fazer a requisição e tratar a resposta
  var response = UrlFetchApp.fetch(apiUrl, options);
  var json = response.getContentText();
  var data = JSON.parse(json);

  // Atualizar os tokens armazenados nas células A2 e B2
  sheet_dados.getRange("D2").setValue(data.refresh_token);
  sheet_dados.getRange("E2").setValue(data.access_token);
}
