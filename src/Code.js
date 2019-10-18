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
