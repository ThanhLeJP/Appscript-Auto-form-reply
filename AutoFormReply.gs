/**
 * Gシートにてメニューボタン生成
 */
function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('返信ドラフト生成')
      .addItem('Enable auto draft replies', 'installTrigger')
      .addToUi();
}

/**
 * フォーム送信される際、トリッガー定義
 */
function installTrigger() {
  ScriptApp.newTrigger('onFormSubmit')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onFormSubmit()
      .create();
}

/**
 * フォームSubmitされる度、下書き生成
 *
 * @param {Object} event - Form submit event
 */
function onFormSubmit(e) {
  let responses = e.namedValues;

  // parse form response data
  let timestamp = responses.Timestamp[0];
  let email = responses['Email address'][0].trim();

  // create email body
  let emailBody = createEmailBody(responses);

  // create draft email
  createDraft(timestamp, email, emailBody);
}

/**
 * Creates email body and includes feedback from Google Form.
 *
 * @param {string} responses - The form response data
 * @return {string} - The email body as an HTML string
 */
function createEmailBody(responses) {
  // parse form response data
  let name = responses.Name[0].trim();
  let industry = responses['What industry do you work in?'][0];
  let source = responses['How did you find out about this course?'][0];
  let rating = responses['On a scale of 1 - 5 how would you rate this course?'][0];
  let productFeedback = responses['What could be different to make it a 5 rating?'][0];
  let otherFeedback = responses['Any other feedback?'][0];

  // create email body
  let htmlBody = 'こんにちは ' + name + ',<br><br>' +
    'アンケートのご回答ありがとうございます。.<br><br>' +
      'コース改善には参考させて頂きます。<br><br>' +
          'Thanks,<br>' +
            'BBAチーム<br><br>' +
              '****************************************************************<br><br>' +
                '<i>Your feedback:<br><br>' +
                  'What industry do you work in?<br><br>' +
                    industry + '<br><br>' +
                      'How did you find out about this course?<br><br>' +
                        source + '<br><br>' +
                          'On a scale of 1 - 5 how would you rate this course?<br><br>' +
                            rating + '<br><br>' +
                              'What could be different to make it a 5 rating?<br><br>' +
                                productFeedback + '<br><br>' +
                                  'Any other feedback?<br><br>' +
                                    otherFeedback + '<br><br></i>';

  return htmlBody;
}

/**
 * 回答内容から該当項目にて下書き内容生成
 *
 * @param {string} timestamp Timestamp for the form response
 * @param {string} email Email address from the form response
 * @param {string} emailBody The email body as an HTML string
 */
function createDraft(timestamp, email, emailBody) {
  Logger.log('draft email create process started');

  // create subject line
  let subjectLine = 'Thanks for your course feedback! ' + timestamp;

  // create draft email
  GmailApp.createDraft(
      email,
      subjectLine,
      '',
      {
        htmlBody: emailBody,
      }
  );
}
