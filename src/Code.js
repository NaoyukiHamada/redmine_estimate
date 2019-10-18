var SheetName = {
    SLACK_SETTING: 'SlackSetting',
    MESSAGE_TEMPLATE: 'MessageTemplate'
};

//gasのバグによりconstでエラーになるので、varを使用
var SLACK_SETTING_SHEET_CONTENT_START_INDEX = 2;

//gasのバグによりconstでエラーになるので、varを使用
var SlackSettingSheetColumn = {
    ID: 1,
    PROJECT_NAME: 2,
    VERSION: 3,
    DUE_DATE: 4,
    SLACK_CHANNEL_NAME: 5,
    USER_NAME: 6,
    ICON_NAME: 7,
    MESSAGE_TEMPLATE_ID: 8,
    NOTIFICATION_ON_OFF: 9
};

var MessageTemplateSheetColumn = {
    ID: 1,
    MESSAGE: 2
}

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
    const slackSettingSheet = spreadsheet.getSheetByName(SheetName.SLACK_SETTING);
    //シートの中身を読みこみ
    const lastRow = slackSettingSheet.getLastRow();
    for (var i = SLACK_SETTING_SHEET_CONTENT_START_INDEX; i <= lastRow; i++) {
        const id = slackSettingSheet.getRange(i, SlackSettingSheetColumn.ID).getValue();
        const projectName = slackSettingSheet.getRange(i, SlackSettingSheetColumn.PROJECT_NAME).getValue();
        const version = slackSettingSheet.getRange(i, SlackSettingSheetColumn.VERSION).getValue();
        const dueDate = slackSettingSheet.getRange(i, SlackSettingSheetColumn.DUE_DATE).getValue();
        const slackChannelName = slackSettingSheet.getRange(i, SlackSettingSheetColumn.SLACK_CHANNEL_NAME).getValue();
        const userName = slackSettingSheet.getRange(i, SlackSettingSheetColumn.USER_NAME).getValue();
        const iconName = slackSettingSheet.getRange(i, SlackSettingSheetColumn.ICON_NAME).getValue();
        const messageTemplateId = slackSettingSheet.getRange(i, SlackSettingSheetColumn.MESSAGE_TEMPLATE_ID).getValue();
        const shouldSendToSlack = slackSettingSheet.getRange(i, SlackSettingSheetColumn.NOTIFICATION_ON_OFF).getValue();
        const message = createMessage(spreadsheet, messageTemplateId, version, dueDate);
        if (shouldSendToSlack) {
            sendMessageToSlack(slackChannelName, userName, iconName, message);
        }
    }
}

/**
 * MessageTemplateIdからメッセージを作成
 * @param spreadsheet
 * @param messageTemplateId
 * @param version
 * @param dueDate
 * @returns {string}
 */
function createMessage(spreadsheet, messageTemplateId, version, dueDate) {
    const messageTemplateSheet = spreadsheet.getSheetByName(SheetName.MESSAGE_TEMPLATE);
    const row = findRow(messageTemplateSheet, messageTemplateId, MessageTemplateSheetColumn.ID);
    var message = messageTemplateSheet.getRange(row, MessageTemplateSheetColumn.MESSAGE).getValue();
    switch (messageTemplateId) {
        case 1: {
            message = createEstimateReportMessage(version, dueDate, message);
            break;
        }
        default : {
            break;
        }
    }
    return message
}

/**
 * 進捗定期報告用のメッセージ作成
 * @param version
 * @param dueDate
 * @param message
 */
function createEstimateReportMessage(version, dueDate, message) {
    const formattedDueDate = Utilities.formatDate(dueDate, 'JST', 'yyyy/MM/dd');
    return Utilities.formatString(message, version, 23, formattedDueDate, 30, 7, "3日超過", "2日超過", "1日超過", "1日余剰", "2日余剰");
}

/**
 * 特定の値が何行目にあるかを探す
 *
 * @param sheet
 * @param val 検索対象の値
 * @param col 検索列
 * @returns {number} 検索対象の値の行番号
 */
function findRow(sheet, val, col) {
    const dat = sheet.getDataRange().getValues();
    for (var i = 0; i < dat.length; i++) {
        if (dat[i][col - 1] === val) {
            return i + 1;
        }
    }
    return -1;
}


