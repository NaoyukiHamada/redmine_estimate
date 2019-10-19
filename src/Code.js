/**
 * シート名用Enum
 * @type {{MESSAGE_TEMPLATE: string, SLACK_SETTING: string}}
 */
//gasのバグによりconstでエラーになるので、varを使用
var SheetName = {
    SLACK_SETTING: 'SlackSetting',
    MESSAGE_TEMPLATE: 'MessageTemplate',
    SEND_TEST: "テスト配信"
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
    TITLE: 8,
    MESSAGE_TEMPLATE_ID: 9,
    RED_MINE_QUERY_ID: 10,
    ACTUAL_WORKING_HOURS: 11,
    NOTIFICATION_ON_OFF: 12
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
var DEFAULT_ACTUAL_WORKING_HOURS = 7;

/**
 * Slackにメッセージを送信
 *
 * @param channelName チェンネル名 publicチャンネルの場合は#を頭につける
 * @param userName ユーザー名
 * @param iconName アイコン名 カスタム絵文字を利用可能 e.g. :memo:
 * @param title
 * @param message メッセージ
 * @param isTest テスト配信時はtrue
 */
function sendMessageToSlack(channelName, userName, iconName, title, message, isTest) {
    // const slackSendMessageUrl = 'https://slack.com/api/chat.postMessage';
    const slackSendMessageUrl = 'https://hooks.slack.com/services/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';

    if (title != null && title.length > 0) {
        message = '*' + title + '*\n\n' + message;
    }

    if (isTest != null && isTest) {
        message = '*テスト配信*\n\n' + message;
    }

    const payload =
        {
            // "token": PropertiesService.getScriptProperties().getProperty('slack_access_token'),
            "channel": channelName || "#random",
            "username": userName || "Bot",
            "icon_emoji": iconName || ":robot_face:",
            "text": message || "no message"
        };


    const options =
        {
            "method": "post",
            // "contentType": "application/x-www-form-urlencoded",
            // "payload": JSON.payload
            "payload": JSON.stringify(payload)
        };

    UrlFetchApp.fetch(slackSendMessageUrl, options);
}

/**
 * 特定のSlack設定に応じて、もしくは全てのSlack設定に応じて、Slackに通知
 * @param slackSettingId 特定の設定に応じて通知を行いたいときに入力。全てに通知する時はnull
 * @param isTest テスト配信時はtrue
 */
function sendToSlackBySpecifyOrAllSlackSettings(slackSettingId, isTest) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const slackSettingSheet = spreadsheet.getSheetByName(SheetName.SLACK_SETTING);
    //GASのトリガーで定期実行した時に引数に不正な値が入るので、intかどうかをチェック
    if (slackSettingId != null && slackSettingId === parseInt(slackSettingId, 10)) {
        //特定のslack通知設定を読みこんで、Slackに通知
        const row = findRow(slackSettingSheet, slackSettingId, SlackSettingSheetColumn.ID);
        readSlackSettingAndSendToSlack(spreadsheet, slackSettingSheet, row, isTest)
    } else {
        //全てのslack通知設定を読みこんで、Slackに通知
        const lastRow = slackSettingSheet.getLastRow();
        for (var i = SLACK_SETTING_SHEET_CONTENT_START_INDEX; i <= lastRow; i++) {
            readSlackSettingAndSendToSlack(spreadsheet, slackSettingSheet, i, isTest)
        }
    }
}

/**
 * スプレッドシートからSlack通知設定を読み込んで、Slackに通知
 */
function readSlackSettingAndSendToSlack(spreadsheet, slackSettingSheet, row, isTest) {
    const id = slackSettingSheet.getRange(row, SlackSettingSheetColumn.ID).getValue();
    const projectName = slackSettingSheet.getRange(row, SlackSettingSheetColumn.PROJECT_NAME).getValue();
    const version = slackSettingSheet.getRange(row, SlackSettingSheetColumn.VERSION).getValue();
    const dueDate = slackSettingSheet.getRange(row, SlackSettingSheetColumn.DUE_DATE).getValue();
    const slackChannelName = slackSettingSheet.getRange(row, SlackSettingSheetColumn.SLACK_CHANNEL_NAME).getValue();
    const userName = slackSettingSheet.getRange(row, SlackSettingSheetColumn.USER_NAME).getValue();
    const iconName = slackSettingSheet.getRange(row, SlackSettingSheetColumn.ICON_NAME).getValue();
    const title = slackSettingSheet.getRange(row, SlackSettingSheetColumn.TITLE).getValue();
    const messageTemplateId = slackSettingSheet.getRange(row, SlackSettingSheetColumn.MESSAGE_TEMPLATE_ID).getValue();
    const redMineQueryId = slackSettingSheet.getRange(row, SlackSettingSheetColumn.RED_MINE_QUERY_ID).getValue();
    const actualWorkingHours = slackSettingSheet.getRange(row, SlackSettingSheetColumn.ACTUAL_WORKING_HOURS).getValue();
    const shouldSendToSlack = slackSettingSheet.getRange(row, SlackSettingSheetColumn.NOTIFICATION_ON_OFF).getValue();
    const message = createMessage(spreadsheet, messageTemplateId, version, dueDate, redMineQueryId, actualWorkingHours);
    if (shouldSendToSlack) {
        sendMessageToSlack(slackChannelName, userName, iconName, title, message, isTest);
    }
}

/**
 * MessageTemplateIdからメッセージを作成
 * @param spreadsheet
 * @param messageTemplateId
 * @param version
 * @param dueDate
 * @param redMineQueryId
 * @param actualWorkingHours １日の稼働時間
 * @returns {string}
 */
function createMessage(spreadsheet, messageTemplateId, version, dueDate, redMineQueryId, actualWorkingHours) {
    const messageTemplateSheet = spreadsheet.getSheetByName(SheetName.MESSAGE_TEMPLATE);
    const row = findRow(messageTemplateSheet, messageTemplateId, MessageTemplateSheetColumn.ID);
    var message = messageTemplateSheet.getRange(row, MessageTemplateSheetColumn.MESSAGE).getValue();
    switch (messageTemplateId) {
        case 1: {
            message = createEstimateReportMessage(version, dueDate, message, redMineQueryId, actualWorkingHours);
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
 * @param actualWorkingHours 1日の稼働時間
 */
function createEstimateReportMessage(version, dueDate, message, redMineQueryId, actualWorkingHours) {
    const formattedDueDate = Utilities.formatDate(dueDate, 'JST', 'yyyy/MM/dd');
    const totalEstimateTime = getTotalEstimateTimeFromRedMine(redMineQueryId);
    const workingHours = actualWorkingHours || DEFAULT_ACTUAL_WORKING_HOURS;
    //人日を計算
    const manDay = totalEstimateTime / workingHours;
    return Utilities.formatString(message, version, formattedDueDate, 23, manDay, workingHours, "3日超過", "2日超過", "1日超過", "1日余剰", "2日余剰");
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

/**
 * 設定したSlack通知設定が適切に動作するかをテスト配信
 */
function debugSendToSlack() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sendTestSheet = spreadsheet.getSheetByName(SheetName.SEND_TEST);
    const targetId = sendTestSheet.getRange(2, 1).getValue();
    sendToSlackBySpecifyOrAllSlackSettings(targetId, true);
}


