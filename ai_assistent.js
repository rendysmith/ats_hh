function callGeminiApi(prompt) {
  var datas = PropertiesService.getScriptProperties();

  var apiKey = datas.getProperty('gemini_token'); 
  var model = datas.getProperty('gemini_model'); 

  // 2. Construct the URL
  var url = "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent?key=" + apiKey;

  // 3. Construct the Request Body (Payload)
  var requestBody = {
    "contents": [{
      "parts": [{"text": prompt}]
    }]
  };
  var payload = JSON.stringify(requestBody); // Convert JS object to JSON string

  // 4. Construct the Request Options
  var options = {
    'method': 'post',
    'contentType': 'application/json',
    // 'headers': { 'Content-Type': 'application/json' }, // Alternative to contentType
    'payload': payload,
    'muteHttpExceptions': true // Important to handle potential errors gracefully
  };

  // 5. Make the API Call
  var response = UrlFetchApp.fetch(url, options);

  // 6. Process the Response
  var responseCode = response.getResponseCode();
  var responseBody = response.getContentText();

  if (responseCode === 200) {
    var jsonResponse = JSON.parse(responseBody);
    //Logger.log(jsonResponse);

    // Extract the actual generated text (need to inspect API response structure)
    // Assuming structure: response.candidates[0].content.parts[0].text
    var generatedText = jsonResponse.candidates[0].content.parts[0].text;
    Logger.log("Success!");
    Logger.log(generatedText);
    // Potentially return the text or do something else with it
    return generatedText;
  } else {
    Logger.log("Error: " + responseCode);
    Logger.log(responseBody);
    // Handle the error (e.g., return an error message, throw an exception)
    return "Error calling Gemini API: " + responseBody;
  }
}

function aiTest() {
  var result = callGeminiApi('Сколько лететь до луны??');
  Logger.log(result);
}