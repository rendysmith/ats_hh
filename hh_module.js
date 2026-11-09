/**
 * Выполняет авторизованный GET-запрос к API HH.ru с автоматическим рефрешем токена при 403 token_expired.
 * @param {string} url - URL для запроса
 * @return {object|null} - JSON-ответ или null при ошибке
 */
function hhFetch(url) {
  var scriptProperties = PropertiesService.getScriptProperties();
  var access_token = scriptProperties.getProperty('access_token');

  function doFetch(token) {
    return UrlFetchApp.fetch(url, {
      method: 'GET',
      headers: {
        "Authorization": "Bearer " + token,
        "User-Agent": "YourApp/1.0 (anku@sidorinlab.ru)"
      },
      muteHttpExceptions: true
    });
  }

  var response = doFetch(access_token);
  var RCode = response.getResponseCode();
  var rJson = JSON.parse(response.getContentText());

  // Token expired -> refresh -> retry once
  if (RCode === 403) {
    var err = rJson.errors?.[0];
    if (err?.value === 'token_expired') {
      Logger.log('Access token expired, refreshing...');
      var refreshResult = refreshHhAccessToken();
      if (refreshResult) {
        var newToken = scriptProperties.getProperty('access_token');
        response = doFetch(newToken);
        RCode = response.getResponseCode();
        rJson = JSON.parse(response.getContentText());
      }
    }
  }

  if (RCode !== 200) {
    Logger.log('HH API error: ' + RCode + ' ' + JSON.stringify(rJson));
    return null;
  }

  return rJson;
}

function validAccessToken() {
  var result = hhFetch('https://api.hh.ru/me');
  if (result) {
    Logger.log(result);
    Logger.log(result.email);
  }
  return result;
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
  var url = "https://api.hh.ru/resumes/" + resumeId;

  try {
    var rJson = hhFetch(url);

    if (!rJson) {
      return null;
    }

    return rJson;

  } catch (e) {
    Logger.log('Error occurred: ' + e.message);
    return null;
  }
}

function getNegotiations(vacancyId) {
  Logger.log('***************** getNegotiations *****************')
  var url = "https://api.hh.ru/negotiations/response?vacancy_id=" + vacancyId + "&per_page=50";

  try {
    var rJson = hhFetch(url);

    if (!rJson) {
      return null;
    }

    if (rJson.errors) {
      var err = rJson.errors[0];
      if (err?.value === 'token_expired') {
        // Already retried in hhFetch, if still failing return null
        return null;
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
    var rJson = hhFetch(url);

    if (!rJson) {
      return null;
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
  var authCode = scriptProperties.getProperty('authorization_code');

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

    if (statusCode === 200 && rJson.access_token && rJson.refresh_token) {
      scriptProperties.setProperty('access_token', rJson.access_token);
      scriptProperties.setProperty('refresh_token', rJson.refresh_token);

      const currentTime = new Date().getTime();
      scriptProperties.setProperty('expires_in', currentTime + Number(rJson.expires_in));

      Logger.log('Токены успешно получены и сохранены.');
    }

    return rJson;
  } catch (e) {
    Logger.log('An error occurred: ' + e.toString());
    return null;
  }
}

function getVacIds() {
  //Получаем все ID доступных вакансий и вносим с таблицу если они отсутствуют.

  var me = hhFetch('https://api.hh.ru/me');
  if (!me) {
    return;
  }

  var employer_id = me['employer']['id'];
  Logger.log(employer_id);

  var vacancies = hhFetch('https://api.hh.ru/vacancies?employer_id=' + employer_id);
  if (!vacancies) {
    return;
  }
  Logger.log(vacancies);

  var vacansyIds = [];
  for (var i of vacancies['items']) {
    vacansyIds.push(i['id']);
  }

  Logger.log(vacansyIds);

  if (vacansyIds.length == 0) {
    return;
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
    return;
  }

  Logger.log(ids);

  for (let vacId of vacansyIds) {
    if (!ids.includes(vacId)) {
      appendDataToColumn(sheet, 'vacancy_id', vacId);
    }
  }
}

