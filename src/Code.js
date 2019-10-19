/**
 * シート名用Enum
 * @type {{MESSAGE_TEMPLATE: string, SLACK_SETTING: string}}
 */
//gasのバグによりconstでエラーになるので、varを使用
var SheetName = {
    SLACK_SETTING: 'SlackSetting',
    MESSAGE_TEMPLATE: 'MessageTemplate'
};

/**
 * SlackSettingシートのコンテンツの行始め
 * @type {number}
 */
//gasのバグによりconstでエラーになるので、varを使用
var SLACK_SETTING_SHEET_CONTENT_START_INDEX = 2;

/**
 * SlackSettingシートの各カラムEnum
 * @type {{NOTIFICATION_ON_OFF: number, DUE_DATE: number, SLACK_CHANNEL_NAME: number, ICON_NAME: number, VERSION: number, ID: number, USER_NAME: number, MESSAGE_TEMPLATE_ID: number, RED_MINE_QUERY_ID: number, PROJECT_NAME: number}}
 */
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
    RED_MINE_QUERY_ID: 9,
    NOTIFICATION_ON_OFF: 10
};

/**
 * MessageTemplateシートの各カラムEnum
 * @type {{MESSAGE: number, ID: number}}
 */
//gasのバグによりconstでエラーになるので、varを使用
var MessageTemplateSheetColumn = {
    ID: 1,
    MESSAGE: 2
};

/**
 * 1日の実働時間の基準
 * @type {number} 時間
 */
//gasのバグによりconstでエラーになるので、varを使用
var ACTUAL_WORKING_HOURS = 7;

/**
 * Slackにメッセージを送信
 *
 * @param channelName チェンネル名 publicチャンネルの場合は#を頭につける
 * @param userName ユーザー名
 * @param iconName アイコン名 カスタム絵文字を利用可能 e.g. :memo:
 * @param message メッセージ
 */
function sendMessageToSlack(channelName, userName, iconName, message) {
    const slackSendMessageUrl = 'https://slack.com/api/chat.postMessage';

    const payload =
        {
            "token": PropertiesService.getScriptProperties().getProperty('slack_access_token'),
            "channel": channelName || "#random",
            "username": userName || "Bot",
            "icon_emoji": iconName || ":robot_face:",
            "text": message || "no message"
        };

    const options =
        {
            "method": "post",
            "contentType": "application/x-www-form-urlencoded",
            "payload": payload
        };

    UrlFetchApp.fetch(slackSendMessageUrl, options);
}

/**
 * 特定のSlack設定に応じて、もしくは全てのSlack設定に応じて、Slackに通知
 * @param slackSettingId
 */
function specifySettingOrAllSettingReadSlackSettingAndSendToSlack(slackSettingId) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const slackSettingSheet = spreadsheet.getSheetByName(SheetName.SLACK_SETTING);
    if (slackSettingId != null) {
        //特定のslack通知設定を読みこんで、Slackに通知
        const row = findRow(slackSettingSheet, slackSettingId, SlackSettingSheetColumn.ID)
        readSlackSettingAndSendToSlack(spreadsheet, slackSettingSheet, row)
    } else {
        //全てのslack通知設定を読みこんで、Slackに通知
        const lastRow = slackSettingSheet.getLastRow();
        for (var i = SLACK_SETTING_SHEET_CONTENT_START_INDEX; i <= lastRow; i++) {
            readSlackSettingAndSendToSlack(spreadsheet, slackSettingSheet, i)
        }
    }
}

/**
 * スプレッドシートからSlack通知設定を読み込んで、Slackに通知
 */
function readSlackSettingAndSendToSlack(spreadsheet, slackSettingSheet, row) {
    const id = slackSettingSheet.getRange(row, SlackSettingSheetColumn.ID).getValue();
    const projectName = slackSettingSheet.getRange(row, SlackSettingSheetColumn.PROJECT_NAME).getValue();
    const version = slackSettingSheet.getRange(row, SlackSettingSheetColumn.VERSION).getValue();
    const dueDate = slackSettingSheet.getRange(row, SlackSettingSheetColumn.DUE_DATE).getValue();
    const slackChannelName = slackSettingSheet.getRange(row, SlackSettingSheetColumn.SLACK_CHANNEL_NAME).getValue();
    const userName = slackSettingSheet.getRange(row, SlackSettingSheetColumn.USER_NAME).getValue();
    const iconName = slackSettingSheet.getRange(row, SlackSettingSheetColumn.ICON_NAME).getValue();
    const messageTemplateId = slackSettingSheet.getRange(row, SlackSettingSheetColumn.MESSAGE_TEMPLATE_ID).getValue();
    const redMineQueryId = slackSettingSheet.getRange(row, SlackSettingSheetColumn.RED_MINE_QUERY_ID).getValue();
    const shouldSendToSlack = slackSettingSheet.getRange(row, SlackSettingSheetColumn.NOTIFICATION_ON_OFF).getValue();
    const message = createMessage(spreadsheet, messageTemplateId, version, dueDate, redMineQueryId);
    if (shouldSendToSlack) {
        sendMessageToSlack(slackChannelName, userName, iconName, message);
    }
}

/**
 * MessageTemplateIdからメッセージを作成
 * @param spreadsheet
 * @param messageTemplateId
 * @param version
 * @param dueDate
 * @param redMineQueryId
 * @returns {string}
 */
function createMessage(spreadsheet, messageTemplateId, version, dueDate, redMineQueryId) {
    const messageTemplateSheet = spreadsheet.getSheetByName(SheetName.MESSAGE_TEMPLATE);
    const row = findRow(messageTemplateSheet, messageTemplateId, MessageTemplateSheetColumn.ID);
    var message = messageTemplateSheet.getRange(row, MessageTemplateSheetColumn.MESSAGE).getValue();
    switch (messageTemplateId) {
        case 1: {
            message = createEstimateReportMessage(version, dueDate, message, redMineQueryId);
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
 * @param redMineQueryId
 */
function createEstimateReportMessage(version, dueDate, message, redMineQueryId) {
    const formattedDueDate = Utilities.formatDate(dueDate, 'JST', 'yyyy/MM/dd');
    const totalEstimateTime = getTotalEstimateTimeFromRedMine(redMineQueryId);
    //人日
    const manDay = totalEstimateTime / ACTUAL_WORKING_HOURS;
    return Utilities.formatString(message, version, 23, formattedDueDate, manDay, ACTUAL_WORKING_HOURS, "3日超過", "2日超過", "1日超過", "1日余剰", "2日余剰");
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

/**
 * Redmineから指定のカスタムクエリの全チケットの合計予定工数を算出
 * チケットの進捗率に応じて合計予定工数を削減
 * RedmineAPIの制限で一度に取得できるチケットは100件まで
 *
 * @param redMineQueryId redmineで作成したカスタムクエリのID(カスタムクエリを開いた時のURLに記載)
 * @returns {number} 合計予定工数
 */
function getTotalEstimateTimeFromRedMine(redMineQueryId) {
    const baseUrl = 'https://www2195ue.sakura.ne.jp/redmine/';
    const path = 'issues.json';
    const token = PropertiesService.getScriptProperties().getProperty('red_mine_token');
    const limit = 100;
    //何ページ読み込むべきかを算出
    const requestUrlForIssueTotalCount = baseUrl + path + '?query_id=' + redMineQueryId + '&key=' + token;
    const responseJsonForIssueTotalCount = UrlFetchApp.fetch(requestUrlForIssueTotalCount).getContentText();
    const responseMapForIssueTotalCount = JSON.parse(responseJsonForIssueTotalCount);
    const totalCount = responseMapForIssueTotalCount.total_count;
    var pageCount;
    if (totalCount < 100) {
        pageCount = 1;
    } else {
        pageCount = Math.ceil(totalCount / limit);
    }

    //チケット全ての合計予定工数を算出
    var totalEstimate = 0;
    for (var i = 1; i <= pageCount; i++) {
        var requestUrl = baseUrl + path + '?query_id=' + redMineQueryId + '&key=' + token + '&limit=' + limit + '&page=' + i;
        var responseJson = UrlFetchApp.fetch(requestUrl).getContentText();
        var responseMap = JSON.parse(responseJson);
        responseMap.issues.forEach(function (issue) {
            var estimateTime = issue.estimated_hours || 0;
            //現在の進捗率に応じて予定工数を削減
            var estimateTimeByProgress = estimateTime * (100 - issue.done_ratio) / 100;
            totalEstimate += estimateTimeByProgress;
        });
    }
    return totalEstimate
}


