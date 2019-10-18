/**
 * Slackにメッセージを送信
 *
 * @param channelName チェンネル名 publicチャンネルの場合は#を頭につける
 * @param userName ユーザー名
 * @param iconName アイコン名 カスタム絵文字を利用可能 e.g. :memo:
 * @param message メッセージ
 */
function sendMessageToSlack(channelName, userName, iconName, message) {
    const slackWebHookUrl = 'https://hooks.slack.com/services/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';

    const jsonData =
        {
            "channel": channelName || "#random",
            "username": userName || "Bot",
            "icon_emoji": iconName || ":robot_face:",
            "text": message || "no message"
        };
    const payload = JSON.stringify(jsonData);

    const options =
        {
            "method": "post",
            "contentType": "application/json",
            "payload": payload
        };
    UrlFetchApp.fetch(slackWebHookUrl, options);
}

/**
 * スプレッドシートからSlack通知設定を読み込んで、Slackに通知
 */
function readSlackSettingAndSendToSlack() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(SheetName.SLACK_SETTING);
    //シートの中身を読みこみ
    const lastRow = sheet.getLastRow();
    for (var i = SLACK_SETTING_SHEET_CONTENT_START_INDEX; i <= lastRow; i++) {
        const projectName = sheet.getRange(i, SlackSettingSheetColumn.PROJECT_NAME).getValue();
        const dueDate = sheet.getRange(i, SlackSettingSheetColumn.DUE_DATE).getValue();
        const slackChannelName = sheet.getRange(i, SlackSettingSheetColumn.SLACK_CHANNEL_NAME).getValue();
        const userName = sheet.getRange(i, SlackSettingSheetColumn.USER_NAME).getValue();
        const iconName = sheet.getRange(i, SlackSettingSheetColumn.ICON_NAME).getValue();
        const message = sheet.getRange(i, SlackSettingSheetColumn.MESSAGE).getValue();
        sendMessageToSlack(slackChannelName, userName, iconName, message);
    }
}

var SheetName = {
    SLACK_SETTING: 'SlackSetting'
};

//gasのバグによりconstでエラーになるので、varを使用
var SLACK_SETTING_SHEET_CONTENT_START_INDEX = 2;

//gasのバグによりconstでエラーになるので、varを使用
var SlackSettingSheetColumn = {
    PROJECT_NAME: 1,
    DUE_DATE: 2,
    SLACK_CHANNEL_NAME: 3,
    USER_NAME: 4,
    ICON_NAME: 5,
    MESSAGE: 6
};

