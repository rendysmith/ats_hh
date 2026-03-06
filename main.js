function getCurrentFormattedDate() {
  const date = new Date(); // Gets the current date and time
  const timeZone = Session.getScriptTimeZone(); // Gets the script's time zone
  const formattedDate = Utilities.formatDate(date, timeZone, "HH.mm dd.MM.yyyy");
  
  Logger.log(formattedDate); // Prints the result to the log
  
  return formattedDate;
}

function checkSetting() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = spreadsheet.getSheetByName('prompts');
  var datas = sheet.getDataRange().getValues();

  var data_list = datas[0];
  var vac_posit = data_list.indexOf('vacancy_id')
  var pos_posit = data_list.indexOf('position')

  for (let i = 1; i < datas.length; i++) {
    //Logger.log(i);
    var vacancy_id = datas[i][vac_posit];
    var position = datas[i][pos_posit];

    // if (position) {
    //   continue;
    // }

    var vacancy_data = getVacansyId(vacancy_id);

    var name = vacancy_data['name'];

    var experience = vacancy_data['experience']['name'];

    var description_html = vacancy_data['description'];
    var description = removeHtmlTags(description_html);

    // Logger.log(description_html);
    // Logger.log(description);
    // Utilities.sleep(50000);

    var schedule = vacancy_data['schedule']['name'];	
    var employment = vacancy_data['employment']['name'];	

    if (vacancy_data['salary'] == null) {
      description	= description + "\n" + experience + "\n" + schedule + "\n" + employment + "\n";

    } else {
      var salary_from = vacancy_data['salary']['from'];
      var salary_to = vacancy_data['salary']['to'];
      var salary_currency = vacancy_data['salary']['currency'];
      var salary_gross = vacancy_data['salary']['gross'];

      var salary_range_from = vacancy_data['salary_range']['from'];
      var salary_range_to = vacancy_data['salary_range']['to'];
      var salary_range_currency = vacancy_data['salary_range']['currency'];
      var salary_range_gross = vacancy_data['salary_range']['gross'];

      description	= description + "\n" + experience + "\n" + schedule + "\n" + employment + "\nsalary: " + salary_from + " - " + salary_to + " " + salary_currency + ", Gross: " + salary_gross + "\nsalary_range: " + salary_range_from + " - " + salary_range_to + " " + salary_range_currency + ", Gross: " + salary_range_gross + "\n";

    }

    writeDataByColumnNameAndRowNumber(sheet, "position", i + 1, name);
    writeDataByColumnNameAndRowNumber(sheet, "skills", i + 1, description);
  }
}

function getResume() {
  Logger.log('***************** getResume ********************')
  //Функция для сбора резюме и размещению их каждую в свою таблицу

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('prompts');
  var datas_p = sheet.getDataRange().getValues();

  var data_list = datas_p[0];
  var vac_posit = data_list.indexOf('vacancy_id')
  var pos_posit = data_list.indexOf('position')

  for (let i = 1; i < datas_p.length; i++) {
    var vacancy_id = datas_p[i][vac_posit];
    var position = datas_p[i][pos_posit];

    if (!vacancy_id || !position) { //#если ячейка с вакансией или должностью пустая - пропустить
      continue;
    }

    Logger.log(vacancy_id + " " + position)

    // Check if a sheet with the name 'position' already exists.
    var positionSheet = spreadsheet.getSheetByName(position);

    var list_urls = [];
    // If the sheet doesn't exist, create it.
    if (!positionSheet) {
      var positionSheet = spreadsheet.insertSheet(position);

    } else {
      var positionData = positionSheet.getDataRange().getValues();
      var positionDataList = positionData[0];
      var positResumeUrl = positionDataList.indexOf('resume_url')

      for (let i = 1; i < positionData.length; i++) {
        list_urls.push(positionData[i][positResumeUrl]);
      }
    }

    //Logger.log("Url list: " + list_urls);
    //Utilities.sleep(10000)

    //функция собирает отклики
    var datas_v = getNegotiations(vacancy_id); 
    //Logger.log(datas_v);

    if (datas_v == null) {
      return
    }
    
    const typeValue = datas_v?.[0]?.type;

    if (typeValue !== undefined) {
      var error = datas_v[0]['type'] + ", " + getCurrentFormattedDate();
      writeDataByColumnNameAndRowNumber(sheet, 'status_api', i + 1, error)
      continue

    } else if (datas_v['items'].length == 0) {
      writeDataByColumnNameAndRowNumber(sheet, 'status_api', i + 1, 'Items = 0, >' + getCurrentFormattedDate())
      continue
    }

    Logger.log("Vacansy len: " + datas_v['items'].length)

    //Logger.log(datas['items'])

    for (var data of datas_v['items']) {
      //Logger.log(data);
      if (data["resume"] == null) {
        continue
      }

      //Logger.log(data['resume']);
      //Utilities.sleep(50000);

      var last_name = data['resume']['last_name'];
      var first_name = data['resume']['first_name'];
      var name = last_name + " " + first_name

      var title = data['resume']['title']; //Должность например: 'Тестировщик'
      
      var salary;
      try {
        salary = data['resume']['salary']['amount'] + " " + data['resume']['salary']['currency'];
      } catch(e) {
        salary = null; // или "" вместо null
      }

      //var alternate_url = data['resume']['alternate_url'];

      var resume_id = data['resume']['id'];
      var resume_url = "https://hh.ru/resume/" + resume_id

      if (list_urls.includes(resume_url)) {
        continue
      }

      var experience_json = getResumeId(resume_id);

      experience = ''
      for (var exp of experience_json['experience']) {
        var company = exp['company'];
        var position = exp['position'];
        var description = exp['description'];

        experience = experience + company + "\n" + position + "\n" + description + "\n\n"

      }

      var datasJson = {"title": title, "name": name, "salary": salary, "experience": experience, "resume_url": resume_url, "perc_match": '', 'summary': ''};
      appendRowByColumnNames(positionSheet, datasJson);

      Utilities.sleep(2000)
    }
    writeDataByColumnNameAndRowNumber(sheet, 'status_api', i + 1, "OK! " + getCurrentFormattedDate())
  }
}

function aiAnalyst() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();

  var sheetPrompts = spreadsheet.getSheetByName('prompts');
  var datasPrompts = sheetPrompts.getDataRange().getValues();

  var posPosition = datasPrompts[0].indexOf('position');
  var posSkills = datasPrompts[0].indexOf('skills');
  var posPerc = datasPrompts[0].indexOf('perc');
  Logger.log(posPosition + " " + posSkills);

  for (const sheet of sheets) {
    var sheetName = sheet.getName();
    if (sheetName == 'prompts') {
      continue
    }

    Logger.log("SheetName: " + sheetName)

    for (let i = 1; i < datasPrompts.length; i++) {
      Logger.log(datasPrompts[i][posPosition]);
      if (datasPrompts[i][posPosition] == sheetName) { //Поиск параметров по имени
        var skillsPrompts = datasPrompts[i][posSkills]; //Описание вакансии
        var percMax = datasPrompts[i][posPerc]; //максимальный процент
        break;
      }
    }

    Logger.log(skillsPrompts);

    var datasVacansy = sheet.getDataRange().getValues();
    var datasVacansyList = datasVacansy[0];

    var posExperience = datasVacansyList.indexOf('experience');
    var posSalary = datasVacansyList.indexOf('salary');
    var posPercMatch = datasVacansyList.indexOf('perc_match');

    for (let i = 1; i < datasVacansy.length; i++) {
      var perMatch = datasVacansy[i][posPercMatch]; //если ячейка perc_match пустая, пропускаем.
      //Logger.log(perMatch);

      if (perMatch) {
        Logger.log(i + " " + posPercMatch + ": Ячейка уже заполнена.")

        //Скрывать вакансии с низким процентом
        if (perMatch < percMax) {
          sheet.hideRows(i + 1); 
        }

        continue;
      } 

      var experience = datasVacansy[i][posExperience] + "\nSalary: " + datasVacansy[i][posSalary];
      //Logger.log(experience);
 
      var prompt = Utilities.formatString(promptS, skillsPrompts, experience)
      var result = callGeminiApi(prompt);

      Logger.log(typeof result);

      if (typeof result === "string") {
      result = JSON.parse(result);
      }

      Logger.log(result);

      writeDataByColumnNameAndRowNumber(sheet, 'perc_match', i + 1, result[0])
      writeDataByColumnNameAndRowNumber(sheet, 'summary', i + 1, result[1])

      Utilities.sleep(60000);

    }





    //var datas = sheet.getDataRange().getValues();
    //Logger.log(datas);
    




  }
}
