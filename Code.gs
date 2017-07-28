var ADDON_TITLE = 'Auto Forward Add-On';

function onOpen(e) {
  FormApp.getUi()
      .createAddonMenu()
      .addItem('Setup', 'showSetup')
      .addItem('About', 'showAbout')
      .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSetup() {
  var ui = HtmlService.createHtmlOutputFromFile('Setup')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle(ADDON_TITLE);
  //FormApp.getUi().showModalDialog(ui, 'Setup ' + ADDON_TITLE);
  FormApp.getUi().showSidebar(ui);
}

function showAbout() {
  var ui = HtmlService.createHtmlOutputFromFile('About')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(420)
      .setHeight(270);
  FormApp.getUi().showModalDialog(ui, 'About ' + ADDON_TITLE);
}

function saveSettings(settings) {
  var form = FormApp.getActiveForm();
  var triggers = ScriptApp.getUserTriggers(form);

  // determine if the addon is active
  settings.active = (settings.emailField !== '' && settings.subject !== '' && settings.body !== '');
  
  // save the settings to google service
  PropertiesService.getDocumentProperties().setProperties(settings);
  
  // find the current trigger if set
  var trigger = null;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getEventType() == ScriptApp.EventType.ON_FORM_SUBMIT && triggers[i].getHandlerFunction() === 'respondToFormSubmit') {
      trigger = triggers[i];
    }
  }
  
  // add or remove trigger if needed
  if (settings.active === true && trigger === null) {
    // not found but add-on is active, create
    var trigger = ScriptApp.newTrigger('respondToFormSubmit')
      .forForm(form)
      .onFormSubmit()
      .create();
  } else if (settings.active === false && trigger !== null) {
    // found but add-on is deactivated, remove
    ScriptApp.deleteTrigger(trigger);
  }
}

function getSettings() {
  var settings = PropertiesService.getDocumentProperties().getProperties();
  settings.active = settings.active || false;

  // create a list of all the questions/form items for placeholder insertion in the setup page
  var formQuestions = FormApp.getActiveForm().getItems(); // FormApp.ItemType.TEXT
  settings.formQuestions = [];
  for (var i = 0; i < formQuestions.length; i++) {
    settings.formQuestions.push({
      title: formQuestions[i].getTitle(),
      id: formQuestions[i].getId(),
      type: formQuestions[i].getType()
    });
  }
  
  // add map of form types
  settings.formTypes = {
    'TEXT': FormApp.ItemType.TEXT
  };
  
  return settings;
}

function respondToFormSubmit(e) {
  var settings = PropertiesService.getDocumentProperties();
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  // check if authorization is oke
  if (authInfo.getAuthorizationStatus() == ScriptApp.AuthorizationStatus.REQUIRED) {
    sendReauthorizationRequest();
    return;
  }
  // check if addon is setup
  if (settings.getProperty('active') == false) {
    return;
  }
  // check mail quota
  if (MailApp.getRemainingDailyQuota() == 0) {
    return;
  }

  // get form
  var formResponse = e.response;
  var form = FormApp.getActiveForm();
  
  // info for e-mail
  var toAddress = settings.getProperty('emailField');
  var subject = settings.getProperty('subject');
  var ccAddresses = settings.getProperty('ccAddress');
  var message = settings.getProperty('body');
  var extraOptions = {name: form.getTitle()};
  
  var numberofitems = form.getItems().length;
  Logger.log("" + numberofitems + "")
  // replace placeholders in subject, cc, message
  for (var i = 0; i < numberofitems; i ++) {
    var items = form.getItems();
    var item = items[i];
    if (item.getType() != 'SECTION_HEADER' && item.getType() != 'IMAGE' && item.getType()!= 'PAGE_BREAK') {
      if (formResponse.getResponseForItem(item) == null) {
        subject = subject.replace('{{' + item.getId() + ':' + item.getTitle() + '}}', '');
        ccAddresses = ccAddresses.replace('{{' + item.getId() + ':' + item.getTitle() + '}}', '');
        message = message.replace('{{' + item.getId() + ':' + item.getTitle() + '}}', item.getTitle() + ": " + '');
      } else {
        subject = subject.replace('{{' + item.getId() + ':' + item.getTitle() + '}}', formResponse.getResponseForItem(item).getResponse());
        ccAddresses = ccAddresses.replace('{{' + item.getId() + ':' + item.getTitle() + '}}', formResponse.getResponseForItem(item).getResponse());
        message = message.replace('{{' + item.getId() + ':' + item.getTitle() + '}}', item.getTitle() + ": " + formResponse.getResponseForItem(item).getResponse());
      }
    }
  }
  
  // attach cc if set
  if (ccAddresses !== null && ccAddresses !== '') {
    extraOptions.cc = ccAddresses;
  }
  
  // and mail
  MailApp.sendEmail(toAddress, subject, message, extraOptions);
}

function sendReauthorizationRequest() {
  var settings = PropertiesService.getDocumentProperties();
  var authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL);
  var lastAuthEmailDate = settings.getProperty('lastAuthEmailDate');
  var today = new Date().toDateString();
  if (lastAuthEmailDate != today) {
    if (MailApp.getRemainingDailyQuota() > 0) {
      var template = HtmlService.createTemplateFromFile('AuthorizationEmail');
      template.addon_title = ADDON_TITLE;
      var message = template.evaluate();
      MailApp.sendEmail(Session.getEffectiveUser().getEmail(), 'Authorization Required for the ' + ADDON_TITLE + ' add-on', message.getContent(), { name: ADDON_TITLE, htmlBody: message.getContent() });
    }
    settings.setProperty('lastAuthEmailDate', today);
  }
}
