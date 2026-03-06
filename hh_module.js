function validAccessToken() {
  var scriptProperties = PropertiesService.getScriptProperties();

  var access_token = scriptProperties.getProperty('access_token');
  
  const url = 'https://api.hh.ru/me';

  // Выполняем GET-запрос
  var response = UrlFetchApp.fetch(url, {
    method: 'GET',
      headers: {
      "Authorization": "Bearer " + access_token, // токен работодателя
      "User-Agent": "YourApp/1.0 (anku@sidorinlab.ru)" // обязательный контакт
    },
    muteHttpExceptions: true // Не выбрасывать исключения при ошибках HTTP
  });

  var RCode = response.getResponseCode();
  Logger.log(RCode); 

  // Получаем JSON-ответ
  var rJson = JSON.parse(response.getContentText());
  Logger.log(rJson);
  Logger.log(rJson.email);

}

function refreshHhAccessToken() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var refreshToken = scriptProperties.getProperty('refresh_token');
  var clientId = scriptProperties.getProperty('client_id');
  var clientSecret = scriptProperties.getProperty('client_secret'); ;

  Logger.log(refreshToken);

  const tokenUrl = 'https://api.hh.ru/token';
  const payload = {
    'grant_type': 'refresh_token',
    'refresh_token': refreshToken,
    'client_id': clientId,
    'client_secret': clientSecret
  };

  const options = {
    'method': 'post',
    'contentType': 'application/x-www-form-urlencoded',
    'payload': payload
  };

  try {
    const response = UrlFetchApp.fetch(tokenUrl, options);
    const result = JSON.parse(response.getContentText());

    // Проверяем наличие новых токенов в ответе
    if (result.access_token && result.refresh_token) {
      // Сохраняем новую пару токенов в свойствах скрипта
      scriptProperties.setProperty('access_token', result.access_token);
      scriptProperties.setProperty('refresh_token', result.refresh_token);

      const currentTime = new Date().getTime();
      scriptProperties.setProperty('expires_in', currentTime + Number(result.expires_in));

      Logger.log('Токены успешно обновлены и сохранены.');
      return result;
    } else {
      Logger.log('Ошибка при обновлении токенов: ' + JSON.stringify(result));
      return null;
    }
  } catch (e) {
    Logger.log('Произошла ошибка при запросе: ' + e.toString());
    return null;
  }
}

function getResumeId(resumeId) {
  // Получаем сервис свойств скрипта.
  var scriptProperties = PropertiesService.getScriptProperties();
  // Получаем значение по ключу 'access_token'.
  var access_token = scriptProperties.getProperty('access_token');

  var url = "https://api.hh.ru/resumes/" + resumeId;

  try {
    // Выполняем GET-запрос
    var response = UrlFetchApp.fetch(url, {
      method: 'GET',
        headers: {
        "Authorization": "Bearer " + access_token, // токен работодателя
        "User-Agent": "YourApp/1.0 (anku@sidorinlab.ru)" // обязательный контакт
      },
      muteHttpExceptions: true // Не выбрасывать исключения при ошибках HTTP
    });

    // Получаем JSON-ответ
    var rJson = JSON.parse(response.getContentText());
    
    // Выводим JSON в консоль (аналог pprint)
    //Logger.log(JSON.stringify(rJson, null, 2));

    // Выводим код ответа
    //Logger.log('Code: ' + response.getResponseCode());

    // Проверяем, если код ответа 403
    if (response.getResponseCode() == 403) {
      Logger.log('- Error: ' + JSON.stringify(rJson.errors));
    }
    return rJson;

  } catch (e) {
    Logger.log('Error occurred: ' + e.message);
    return null;
  }
}

function getNegotiations(vacancyId) {
  Logger.log('***************** getNegotiations *****************')
  // Получаем сервис свойств скрипта.
  var scriptProperties = PropertiesService.getScriptProperties();
  // Получаем значение по ключу 'access_token'.
  var access_token = scriptProperties.getProperty('access_token');

  //var url = "https://api.hh.ru/negotiations?vacancy_id=" + vacancyId + "&per_page=50";
  var url = "https://api.hh.ru/negotiations/response?vacancy_id=" + vacancyId + "&per_page=50";

  try {
    // Выполняем GET-запрос
    var response = UrlFetchApp.fetch(url, {
      method: 'GET',
        headers: {
        "Authorization": "Bearer " + access_token, // токен работодателя
        "User-Agent": "YourApp/1.0 (anku@sidorinlab.ru)" // обязательный контакт
      },
      muteHttpExceptions: true // Не выбрасывать исключения при ошибках HTTP
    });

    // Получаем JSON-ответ
    var rJson = JSON.parse(response.getContentText());
    Logger.log(rJson);

    var RCode = response.getResponseCode()
    Logger.log('RCode: ' + RCode)

    if (RCode != 200) {
      var errorsArray = rJson.errors[0];
      Logger.log(errorsArray);
      var ErrorValue = errorsArray['value'];
      var ErrorType = errorsArray['type'];

      if (RCode == 400 || RCode == 403) {
        Logger.log('Error HH');

        if (ErrorValue == 'token_expired') {
          Logger.log('Время жизни access_token завершилось, необходимо получить');
          refreshHhAccessToken();
          Logger.log('access_token обновлен.')
          return null
        }
      }

      else if (RCode == 404) {
        Logger.log('- Error 404');
        Logger.log(errorsArray);
        return {['items']: []};
      }
    }

    return rJson;

  } catch (e) {
    Logger.log('Error occurred: ' + e.message);
    return null;
  }
}

function getVacansyId(vacancyId) {
  /*
   * Возвращает подробную информацию по указанной вакансии
   * @param {string} vacancyId - ID вакансии
   * @return {object} - JSON-ответ от API
   */
  var url = 'https://api.hh.ru/vacancies/' + vacancyId;

  try {
    // Выполняем GET-запрос
    var response = UrlFetchApp.fetch(url, {
      method: 'GET',
      muteHttpExceptions: true // Не выбрасывать исключения при ошибках HTTP
    });

    // Получаем JSON-ответ
    var rJson = JSON.parse(response.getContentText());
    
    // Выводим JSON в консоль (аналог pprint)
    //Logger.log(JSON.stringify(rJson, null, 2));

    // Выводим код ответа
    //Logger.log('Code: ' + response.getResponseCode());

    // Проверяем, если код ответа 403
    var RCode = response.getResponseCode()
    if (response.getResponseCode() == 403) {
      Logger.log('- Error: ' + RCode  + " " + JSON.stringify(rJson.errors));
    }
    return rJson;

  } catch (e) {
    Logger.log('Error occurred: ' + e.message);
    return null;
  }
}

function getCode() {
    // Получаем сервис свойств скрипта.
  var scriptProperties = PropertiesService.getScriptProperties();
  // Получаем значение по ключу 'token_name'.
  //var tokenName = scriptProperties.getProperty('token_name');

  var urlContent = "https://hh.ru/oauth/authorize?response_type=code&client_id=%s&redirect_uri=%s"

  var client_id = scriptProperties.getProperty('client_id'); 
  var redirectUri = scriptProperties.getProperty('redirect_uri'); 
  var url = Utilities.formatString(urlContent, client_id, redirectUri)
  Logger.log(url);
  return url;
}

function getAccessToken() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var clientId = scriptProperties.getProperty('client_id');
  var clientSecret = scriptProperties.getProperty('client_secret');
  var redirect_uri = scriptProperties.getProperty('redirect_uri');
  var authCode = "GUB0CQ83EAOPGLVRGTVAOIUS7I1QP6UMMJ1H69CREA0BUTN38IVI0I8D8HLEVNAE";

  const url = 'https://api.hh.ru/token';
  
  // const payload = {
  //   'grant_type': 'authorization_code',
  //   'client_id': clientId,
  //   'client_secret': clientSecret,
  //   'code': authCode
  // };

  const payload =
  `grant_type=authorization_code` +
  `&client_id=${clientId}` +
  `&client_secret=${clientSecret}` +
  `&code=${authCode}` +
  `&redirect_uri=${encodeURIComponent(redirect_uri)}`;

  const options = {
    'method': 'post',
    'contentType': 'application/x-www-form-urlencoded',
    'headers': {
      'User-Agent': "YourApp/1.0 (anku@sidorinlab.ru)"
    },
    'payload': payload
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const statusCode = response.getResponseCode();
    const responseBody = response.getContentText();
    const rJson = JSON.parse(responseBody);
    
    Logger.log('Status Code: ' + statusCode);
    Logger.log('Response: ' + responseBody);
    Logger.log('Json: ' + rJson);
    
    if (statusCode === 400) {
      if (rJson.error_description === 'code has already been used') {
        Logger.log('Код уже был использован (его можно использовать только один раз).');

      } else if (rJson.error_description === 'error_description') {
        // У вас в Python-коде опечатка, здесь должно быть 'code has expired' или аналогичная ошибка.
        Logger.log('Код истёк.');
      }
    }
    
    scriptProperties.setProperty('access_token', result.access_token);
    scriptProperties.setProperty('refresh_token', result.refresh_token);
    
  
    return rJson;
  } catch (e) {
    Logger.log('An error occurred: ' + e.toString());
    return null;
  }
}

function getVacIds() {
  //Получаем все ID доступных вакансий и вносим с таблицу если они отсутствуют.

  var scriptProperties = PropertiesService.getScriptProperties();
  // Получаем значение по ключу 'access_token'.
  var access_token = scriptProperties.getProperty('access_token');

  var headers = {
      "Authorization": "Bearer " + access_token, // токен работодателя
      "User-Agent": "YourApp/1.0 (anku@sidorinlab.ru)" // обязательный контакт
    }

  var url = 'https://api.hh.ru/me';

  // Выполняем GET-запрос
  var response = UrlFetchApp.fetch(url, {
    method: 'GET',
      headers: headers,
    muteHttpExceptions: true // Не выбрасывать исключения при ошибках HTTP
  });

  var rJson = JSON.parse(response.getContentText());
  var RCode = response.getResponseCode();

  Logger.log(rJson);

  var employer_id = rJson['employer']['id']

  Logger.log(employer_id);

  var url_id = 'https://api.hh.ru/vacancies?employer_id=' + employer_id;
  var response = UrlFetchApp.fetch(url_id, {
    method: 'GET',
    muteHttpExceptions: true // Не выбрасывать исключения при ошибках HTTP
  });

  var rJson_id = JSON.parse(response.getContentText());
  Logger.log(rJson_id);

  var vacansyIds = []

  for (var i of rJson_id['items']) {
    vacansyIds.push(i['id'])
  }

  Logger.log(vacansyIds);

  if (vacansyIds.length == 0) {
    return
  }

  var datasS = getData('prompts');
  var sheet = datasS[0];
  var data = datasS[1];
  Logger.log(data);

  var ids = [];

  for (let i = 1; i < data.length; i++) {
    ids.push(data[i][0]);
  }

  if (ids.length == 0) {
    return
  }

  Logger.log(ids);

  for (let vacId of vacansyIds) {
    if (!ids.includes(vacId)) {
      appendDataToColumn(sheet, 'vacancy_id', vacId);
    }
  }

  
















}


function test_hh() {
  getNegotiations("128172358");
}