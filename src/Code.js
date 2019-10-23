/**
 * シート名用Enum
 * @type {{MESSAGE_TEMPLATE: string, SLACK_SETTING: string}}
 */
//gasのバグによりconstでエラーになるので、varを使用
var SheetName = {
    SLACK_SETTING: 'SlackSetting',
    MESSAGE_TEMPLATE: 'MessageTemplate',
    SEND_TEST: "テスト配信",
    HOLIDAY: '休日'
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
    MANUAL_TOTAL_ESTIMATE_TIME: 12,
    NOTIFICATION_ON_OFF: 13
};

/**
 * 休日シートの各カラムEnum
 * @type {{HOLIDAY: number}}
 */
//gasのバグによりconstでエラーになるので、varを使用
var HolidaySheetColumn = {
    HOLIDAY: 1,
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
 * 休日シートのコンテンツの行始め
 * @type {number}
 */
//gasのバグによりconstでエラーになるので、varを使用
var HOLIDAY_SHEET_CONTENT_START_INDEX = 2;

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
    const slackSendMessageUrl = PropertiesService.getScriptProperties().getProperty('slack_access_token');

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
    const message = createMessage(spreadsheet, id, messageTemplateId, version, dueDate, redMineQueryId, actualWorkingHours);
    if (shouldSendToSlack) {
        sendMessageToSlack(slackChannelName, userName, iconName, title, message, isTest);
    }
}

/**
 * MessageTemplateIdからメッセージを作成
 * @param spreadsheet
 * @param slackSettingId
 * @param messageTemplateId
 * @param version
 * @param dueDate
 * @param redMineQueryId
 * @param actualWorkingHours １日の稼働時間
 * @returns {string}
 */
function createMessage(spreadsheet, slackSettingId, messageTemplateId, version, dueDate, redMineQueryId, actualWorkingHours) {
    const messageTemplateSheet = spreadsheet.getSheetByName(SheetName.MESSAGE_TEMPLATE);
    const row = findRow(messageTemplateSheet, messageTemplateId, MessageTemplateSheetColumn.ID);
    var message = messageTemplateSheet.getRange(row, MessageTemplateSheetColumn.MESSAGE).getValue();
    switch (messageTemplateId) {
        case 1: {
            message = createEstimateReportMessage(spreadsheet, slackSettingId, version, dueDate, message, redMineQueryId, actualWorkingHours);
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
 * @param spreadsheet
 * @param slackSettingId
 * @param version
 * @param dueDate
 * @param message
 * @param redMineQueryId
 * @param actualWorkingHours 1日の稼働時間
 */
function createEstimateReportMessage(spreadsheet, slackSettingId, version, dueDate, message, redMineQueryId, actualWorkingHours) {
    const formattedDueDate = Utilities.formatDate(dueDate, 'JST', 'yyyy/MM/dd');
    const totalEstimateTime = getTotalEstimateTimeFromRedMine(spreadsheet, slackSettingId, redMineQueryId);
    const workingHours = actualWorkingHours || DEFAULT_ACTUAL_WORKING_HOURS;
    const actualWorkingDay = getActualWorkDay(null, dueDate);
    //人日を計算
    const manDay = totalEstimateTime / workingHours;
    //超過もしくは余剰分を計算
    var overDayForOne = manDay - actualWorkingDay;
    var overDayForOnePointFive = manDay / 1.5 - actualWorkingDay;
    var overDayForTwo = manDay / 2 - actualWorkingDay;
    var overDayForTwoPointFive = manDay / 2.5 - actualWorkingDay;
    var overDayForThree = manDay / 3 - actualWorkingDay;

    /**
     * 工数超過参考値用のメッセージ作成
     *
     * @param overDay
     * @returns {string}
     */
    function createOverDayMessage(overDay) {
        if (overDay > 0) {
            return overDay.toFixed(1) + "日超過"
        } else {
            return overDay.toFixed(1) + "日"
        }
    }

    return Utilities
        .formatString(
            message,
            version,
            formattedDueDate,
            actualWorkingDay,
            manDay,
            workingHours,
            totalEstimateTime,
            createOverDayMessage(overDayForOne),
            createOverDayMessage(overDayForOnePointFive),
            createOverDayMessage(overDayForTwo),
            createOverDayMessage(overDayForTwoPointFive),
            createOverDayMessage(overDayForThree)
        );
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
 * @param spreadsheet
 * @param slackSettingId
 * @param redMineQueryId redmineで作成したカスタムクエリのID(カスタムクエリを開いた時のURLに記載)
 * @returns {number} 合計予定工数
 */
function getTotalEstimateTimeFromRedMine(spreadsheet, slackSettingId, redMineQueryId) {
    //合計見積もり時間が手動で入力されている場合はそれを取得
    var slackSettingSheet = spreadsheet.getSheetByName(SheetName.SLACK_SETTING);
    var row = findRow(slackSettingSheet, slackSettingId, SlackSettingSheetColumn.ID);
    var manualTotalEstimateTime = slackSettingSheet.getRange(row, SlackSettingSheetColumn.MANUAL_TOTAL_ESTIMATE_TIME).getValue();
    if (manualTotalEstimateTime !== -1) {
        return manualTotalEstimateTime;
    }

    //Redmineのカスタムクエリから合計見積もり時間を取得
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

/**
 * 指定した日付の間の稼動日数を取得(開始日を含む稼働日数)
 * startDayを指定しない場合は現在の日付との間で計算
 * @param startDay 開始日 nullable
 * @param endDay 締切日
 * @returns {number} 稼働日数
 */
function getActualWorkDay(startDay, endDay) {
    var targetDay;
    if (startDay == null) {
        var today = new Date();
        //今日の0時を取得するために一度format
        targetDay = new Date(Utilities.formatDate(today, 'JST', 'yyyy/MM/dd'));
    } else {
        targetDay = new Date(Utilities.formatDate(startDay, 'JST', 'yyyy/MM/dd'));
    }

    //締切日になるまで、日付を進めて、稼働日数をカウント
    var workDayCount = 0;
    var otherHolidays = getOtherHolidays();
    do {
        if (!isOff(targetDay, otherHolidays)) {
            workDayCount++
        }
        targetDay = goToNextDay(targetDay);
    } while (targetDay.getTime() <= endDay.getTime());
    return workDayCount;
}

/**
 * 休みか判定
 * @param targetDay 判定する日
 * @param otherHolidays その他の休み 毎回getRange.getValueで読み込むと、「読み込みすぎだ」と怒られるので、引数で取得
 * @returns {boolean} 休みか
 */
function isOff(targetDay, otherHolidays) {
    /**
     * 祝日か
     * @returns {boolean} 終日か
     */
    function isHoliday(targetDay) {
        const calendars = CalendarApp.getCalendarsByName('日本の祝日');
        const count = calendars[0].getEventsForDay(targetDay).length;
        return count === 1;
    }

    /**
     * 週末か
     * @returns {boolean}
     */
    function isWeekend(targetDay) {
        const day = targetDay.getDay();
        return (day === 6) || (day === 0);
    }


    /**
     * 土日祝以外の休日か
     * @param targetDay
     * @param otherHolidays 土日祝以外の休みリスト
     * @returns {boolean} 土日祝以外の休日か
     */
    function isOtherOff(targetDay, otherHolidays) {
        var isOtherOff = false;
        for (var i = 0; i < otherHolidays.length; i++) {
            var holiday = otherHolidays[i];
            if (targetDay.getTime() === holiday.getTime()) {
                isOtherOff = true;
            }
        }
        return isOtherOff;
    }

    return isHoliday(targetDay) || isWeekend(targetDay) || isOtherOff(targetDay, otherHolidays);
}

/**
 * 土日祝以外の休みを休日シートから取得
 * @returns {[]} 土日祝以外の休みリスト
 */
function getOtherHolidays() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const holidaySheet = spreadsheet.getSheetByName(SheetName.HOLIDAY);
    const lastRow = holidaySheet.getLastRow();
    var otherHolidays = [];
    for (var i = HOLIDAY_SHEET_CONTENT_START_INDEX; i <= lastRow; i++) {
        var holiday = holidaySheet.getRange(i, HolidaySheetColumn.HOLIDAY).getValue();
        otherHolidays.push(holiday);
    }
    return otherHolidays;
}

/**
 * 指定の日の次の日を取得
 * @param targetDay
 * @returns {Date}
 */
function goToNextDay(targetDay) {
    //次の日をミリ秒で取得
    var date = targetDay.setDate(targetDay.getDate() + 1);
    //ミリ秒からDateオブジェクトを作成
    date = new Date(date);
    return date;
}


