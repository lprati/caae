/**
 * Uma função executada durante a abertura do documento Spreadsheet.
 * 
 * Insere um menu personalizado que permite a execução das outras funcionalidades.
 */
function onOpen() {
  var menu = [{name: 'Configurar entrevistas', functionName: 'configuraEntrevistas_'},
              {name: 'Ajuda', functionName: 'displayHelp_'},
              {name: 'Mostrar propriedades', functionName: 'displayProps_'}
             ];
  SpreadsheetApp.getActive().addMenu('Entrevistas', menu);
}

function displayProps_() {
  for (var prop in PropertiesService.getScriptProperties().getProperties()) {
    Browser.msgBox(prop + ' ' + PropertiesService.getScriptProperties().getProperty(prop))
  }
  
}

/**
 * Exibe uma mensagem de ajuda.
 * 
 * Deve fornecer ao usuário material instrutivo sobre a utilização do aplicativo.
 */ 
function displayHelp_() {
  Browser.msgBox('*Cries in gs language*');
}


/**
 * Este função acessa as informações de horários inscritas no documento Spreadsheet
 * para configurar o formulário Forms (2) e gerar os eventos na agenda Calendar (3)
 */
function configuraEntrevistas_() {
  if (PropertiesService.getScriptProperties().getProperty('cal_id')) {
    
    Browser.msgBox('O Formulário de Agendamento já foi criado anteriormente!');
    var yes_no = Browser.Buttons.YES_NO;
    var button_response = Browser.msgBox('Deseja limpar as variáveis internas do script?', '', yes_no);
    if (button_response == 'yes') {
      var props = PropertiesService.getScriptProperties()
      var executou = props.getProperty('executou')
      Browser.msgBox(executou);
      PropertiesService.getScriptProperties().deleteAllProperties();
      
      Browser.msgBox('O script foi resetado. Por favor, atualize o navegador.')
    }
  }
  else {
    
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName('Horários');
    var range = sheet.getDataRange();
    var values = range.getValues();
    
    var form = FormApp.getActiveForm();
    
    criaAgenda_(values, range);
    criaFormulario_(ss, values);
    ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
    // ScriptApp.newTrigger('onFormSubmit').forForm(form).onFormSubmit().create()
    // ss.removeMenu('Entrevistas');
  }
}



/**
 * Creates a Google Calendar with events for each conference session in the
 * spreadsheet, then writes the event IDs to the spreadsheet for future use.
 * @param {Array<string[]>} values Cell values for the spreadsheet range.
 * @param {Range} range A spreadsheet range that contains conference data.
 */
function criaAgenda_(values, range) {
  
  var props = PropertiesService.getScriptProperties();
  
  // Cria uma nova agenda
  var cal = CalendarApp.createCalendar('Entrevistas CAAE');
  props.setProperty('cal_id', cal.getId());
  
  // Percorre cada linha da tabela
  for (var i = 1; i < values.length; i++) {
    
    // Extrai as informações da entrevista
    var entrevista = values[i];
    var nome_entrevistador  = entrevista[0];
    var email_entrevistador = entrevista[1];
    var data_entrevista     = entrevista[2];
    var horario_inicial     = entrevista[3];  
    
    // Configura horário de inicio e término
    var inicio_do_evento = joinDateAndTime_(data_entrevista, horario_inicial);
    var fim_do_evento = new Date(inicio_do_evento.getTime());
    var minutes = fim_do_evento.getMinutes();
    minutes += 30;
    fim_do_evento.setMinutes(minutes);
    
    // Título e local
    var titulo_do_evento = 'Entrevista CAAE com ' + nome_entrevistador;
    var local_do_evento = 'Alojamento USP São Carlos, Terceiro Andar';
    
    // Propriedades do evento
    var options = {location: local_do_evento, sendInvites: true};
    var event = cal.createEvent(titulo_do_evento, inicio_do_evento, fim_do_evento, options)
        .setGuestsCanSeeGuests(true);
    entrevista[4] = event.getId();
  }
  range.setValues(values);


}


/**
 * Creates a Google Form that allows respondents to select which conference
 * sessions they would like to attend, grouped by date and start time.
 *
 * @param {Spreadsheet} ss The spreadsheet that contains the conference data.
 * @param {Array<String[]>} values Cell values for the spreadsheet range.
 */
function criaFormulario_(ss, values) {

  var props = PropertiesService.getScriptProperties();
  
  // Percorre a planilha colhendo as informações das entrevistas disponíveis
  var schedule = {};
  for (var i = 1; i < values.length; i++) {
    var session = values[i];
    var name = session[0];
    var day = session[2].toLocaleDateString();
    var time = session[3].toLocaleTimeString();
    time = time.replace('0s BRST', '');

    if (!schedule[name]) {
      schedule[name] = {};
    }
    if (!schedule[name][day]) {
      schedule[name][day] = {};
    }
    
    if (!schedule[name][day][time]) {
      schedule[name][day][time] = day + ' ' + time  + ' - ' + name; 
    }    
    
  }

  // Cria formulário e salva id para alterar posteriormente
  var form = FormApp.create('Inscrição CAAE - Entrevista');
  var form_id = form.getId();
  props.setProperty("form_id", form_id);
  
  // Somente uma resposta por usuário
  // <---------- HABILITAR DEPOIS DOS TESTES!!!!! ---------->
  // form.setLimitOneResponsePerUser(true);
  
  // --> provavelmente desnecessário
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId()); 
  
  form.addTextItem().setTitle('Nome').setRequired(true);
  form.addTextItem().setTitle('Email').setRequired(true);
  form.addSectionHeaderItem().setTitle('Selecione abaixo um horário para sua entrevista');
  
  // Cria item para resposta
  var item = form.addListItem();
  var list_id = item.getId();
  props.setProperty("list_id", list_id);
  
  // Popula lista de horários
  var choices = []; 
  for (var name in schedule) {    
    for (var day in schedule[name]) {
      for (var time in schedule[name][day]) {
        choices.push(schedule[name][day][time]);  
      }
    }
  }
 
  item.setTitle('Horários disponíveis');
  item.setChoiceValues(choices);
  item.setRequired(true);
  
}

/**
 * Creates a Google Form that allows respondents to select which conference
 * sessions they would like to attend, grouped by date and start time.
 *
 * @param {Spreadsheet} ss The spreadsheet that contains the conference data.
 * @param {Array<String[]>} values Cell values for the spreadsheet range.
 */
function atualizaFormulario_(ss, values) {
  
  var props = PropertiesService.getScriptProperties();
  var form = FormApp.openById(props.getProperty('form_id'));

  var item_id = props.getProperty('list_id') 
  var lista_entrevistas = form.getItemById(item_id).asListItem();
  
  
  var old_choices = lista_entrevistas.getChoices();
  var old_len = old_choices.length;
  
  // Percorre a planilha colhendo as informações das entrevistas PARA REMOVER  
  var new_choices = [];
  for (var i = 1; i < values.length; i++) {
    var session = values[i];
    
    var name = session[0];
    var day = session[2].toLocaleDateString();
    var time = session[3].toLocaleTimeString();
    time = time.replace('0s BRST', '');
    
    var choice = day + ' ' + time  + ' - ' + name; 
    if (session[5]) {
           
    }
    else {
      new_choices.push(choice);
    }
  }
  

  
  var new_len = new_choices.length;
  lista_entrevistas.setChoiceValues(new_choices);
  
  props.setProperty('old_len', old_len);
  props.setProperty('new_len', new_len);
}

/**
 * A trigger-driven function that sends out calendar invitations and a
 * personalized Google Docs itinerary after a user responds to the form.
 *
 * @param {Object} e The event parameter for form submission to a spreadsheet;
 *     see https://developers.google.com/apps-script/understanding_events
 */
function onFormSubmit(e) {

  var user = {name: e.namedValues['Nome'][0], email: e.namedValues['Email'][0]};
  
  // Grab the session data again so that we can match it to the user's choices.
  var ss = SpreadsheetApp.getActive();
  var range = ss.getSheetByName('Horários').getDataRange();
  var values = range.getValues();
  
  for (var i = 1; i < values.length; i++) {
    var session = values[i];
    
    var name = session[0];
    var day = session[2].toLocaleDateString();
    var time = session[3].toLocaleTimeString();
    time = time.replace('0s BRST', '');

    var title = 'Horários disponíveis';
    var choice = day + ' ' + time  + ' - ' + name;
    
    // For every selection in the response, find the matching timeslot and title
    // in the spreadsheet and add the session data to the response array.
    if (e.namedValues[title] && e.namedValues[title] == choice) {
      session[5] = user.email;
      range.setValues(values);
    }
  }
  
  
  atualizaFormulario_(ss, values);
  //sendDoc_(user, response);
}

/**
 * Add the user as a guest for every session he or she selected.
 * @param {object} user An object that contains the user's name and email.
 * @param {Array<String[]>} response An array of data for the user's session choices.
 */
function sendInvites_(user, response) {
  
  var id = PropertiesService.getScriptProperties().getProperty('cal_id');
  var cal = CalendarApp.getCalendarById(id);

  for (var i = 0; i < response.length; i++) {
    cal.getEventSeriesById(response[i][4]).addGuest(user.email);
      
  }
}

/**
 * Creates a single Date object from separate date and time cells.
 *
 * @param {Date} date A Date object from which to extract the date.
 * @param {Date} time A Date object from which to extract the time.
 * @return {Date} A Date object representing the combined date and time.
 */
function joinDateAndTime_(date, time) {
  date = new Date(date);
  date.setHours(time.getHours());
  date.setMinutes(time.getMinutes());
  return date;
}

/**
 * Create and share a personalized Google Doc that shows the user's itinerary.
 * @param {object} user An object that contains the user's name and email.
 * @param {Array<string[]>} response An array of data for the user's session choices.
 *
function sendDoc_(user, response) {
  var doc = DocumentApp.create('Conference Itinerary for ' + user.name)
      .addEditor(user.email);
  var body = doc.getBody();
  var table = [['Session', 'Date', 'Time', 'Location']];
  for (var i = 0; i < response.length; i++) {
    table.push([response[i][0], response[i][1].toLocaleDateString(),
        response[i][2].toLocaleTimeString(), response[i][4]]);
  }
  body.insertParagraph(0, doc.getName())
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  table = body.appendTable(table);
  table.getRow(0).editAsText().setBold(true);
  doc.saveAndClose();

  // Email a link to the Doc as well as a PDF copy.
  MailApp.sendEmail({
    to: user.email,
    subject: doc.getName(),
    body: 'Thanks for registering! Here\'s your itinerary: ' + doc.getUrl(),
    attachments: doc.getAs(MimeType.PDF)
  });
}
*/