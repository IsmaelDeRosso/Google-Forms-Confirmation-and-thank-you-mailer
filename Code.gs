/**
 * Copyright (c) 2016 Markei.nl
 * Permission is hereby granted, free of charge, to any person obtaining a copy 
 * of this software and associated documentation files (the "Software"), to 
 * deal in the Software without restriction, including without limitation the 
 * rights to use, copy, modify, merge, publish, distribute, sublicense, and/or 
 * sell copies of the Software, and to permit persons to whom the Software is 
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in 
 * all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING 
 * FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
 * DEALINGS IN THE SOFTWARE.
 */

/**
 * @OnlyCurrentDoc
 */

/**
 * A global constant String holding the title of the add-on. This is
 * used to identify the add-on in the notification emails.
 */
var ADDON_TITLE = 'Confirmation and thank you mailer';

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
    Logger.log('sendReauthorizationRequest');
    sendReauthorizationRequest();
    return;
  }
  // check if addon is setup
  if (settings.getProperty('active') == false) {
    Logger.log('not active');
    return;
  }
  // check mail quota
  if (MailApp.getRemainingDailyQuota() == 0) {
    Logger.log('no remaing daily mail quota');
    return;
  }

  // get form
  var formResponse = e.response;
  var form = FormApp.getActiveForm();
  
  // info for e-mail
  var toAddress = formResponse.getResponseForItem(form.getItemById(settings.getProperty('emailField'))).getResponse();
  var subject = settings.getProperty('subject');
  var ccAddresses = settings.getProperty('ccAddress');
  var message = settings.getProperty('body');
  var extraOptions = {name: form.getTitle()};
  
  // replace placeholders in subject, cc, message
  for (var i = 0; i < form.getItems().length; i ++) {
    var formItem = form.getItems()[i];
    subject = subject.replace('{{' + formItem.getId() + ':' + formItem.getTitle() + '}}', formResponse.getResponseForItem(formItem).getResponse());
    ccAddresses = ccAddresses.replace('{{' + formItem.getId() + ':' + formItem.getTitle() + '}}', formResponse.getResponseForItem(formItem).getResponse());
    message = message.replace('{{' + formItem.getId() + ':' + formItem.getTitle() + '}}', formResponse.getResponseForItem(formItem).getResponse());
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
