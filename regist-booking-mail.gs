/**
 * TODO
 * ラベル名とラベルを貼るメールの設定
 * カレンダーID
 * カレンダーの登録内容
 * メールの運用
 * サーフィンとヨガの使い分け
 */

/** カレンダー登録関数 */
function importEmailsToCalendar() {
  // 定数
  const LABEL_NAME = "予約メール"; // Gmailのラベル名
  const CALENDER_ID = "ryuji18270608@gmail.com"; // カレンダーのID
  const LINE_CODE = '\n';
  // ログファイルの準備
  const logFile = prepareLogFile(prepareLogFolder());
  // 登録件数カウンタ
  let createCount = 0;

  try {
    Logger.log("カレンダー登録関数 開始");
    // 1.Gmailを検索
    const threads = GmailApp.search("label:" + LABEL_NAME + " is:unread");
    Logger.log("Gmailの検索結果: %s件", threads.length);
    // 2.対象が存在する
    if (!threads) return;

    for (let i = 0; i < threads.length; i++) {
      const thread = threads[i];
      const messages = thread.getMessages();
      
      for (var j = 0; j < messages.length; j++) {
        const message = messages[j];
        const subject = message.getSubject();
        const bodyArray = message.getPlainBody().split("\r\n");
        // 3.メール本文の解析
        // 取得対象となる項目を取得する
        const bookingNumber = extractTargetString(bodyArray, "予約番号："); // 予約番号
        const date = extractTargetString(bodyArray, "日時："); // 日時
        const name = extractTargetString(bodyArray, "氏名："); // 氏名
        const plan = extractTargetString(bodyArray, "プラン名（コース名）："); // プラン
        const numberOfPeople = "人数:" + extractTargetString(bodyArray, "大人・小人"); // 人数
        const descriptionArray = [bookingNumber, name, plan, numberOfPeople]; // カレンダーの説明文字列作成用の配列
        
        // カレンダーへ登録する項目を作成
        const title = "Activity Japan"; // タイトル
        const startTime = extractTimeString(plan); // プラン名から開始時間を取得
        const startDate = createDateObject(date, startTime.substring(0, 2), startTime.substring(3, 5)); // 開始時間
        const endDate = new Date(startDate); // 終了時間
        endDate.setHours(endDate.getHours() + 3);
        const description = descriptionArray.join(LINE_CODE); // 説明

        // 4.Googleカレンダーに登録
        const targetEvents = CalendarApp.getCalendarById(CALENDER_ID).getEvents(startDate, endDate);
        const alreadyCreateEvent = targetEvents.find((event) => event.getDescription().includes(bookingNumber));
        if (alreadyCreateEvent) {
          Logger.log("%sは既にカレンダーに登録されています", bookingNumber);
        } else {
          const createdEvent = CalendarApp.getCalendarById(CALENDER_ID).createEvent(title, startDate, endDate, {
            description: description
          });
          // サーフィンとヨガで作成イベントの色を出し分ける
          plan.includes("サーフィン") ? createdEvent.setColor(CalendarApp.EventColor.BLUE) : createdEvent.setColor(CalendarApp.EventColor.RED)
          createCount++; // 登録件数のカウンタを加算
          Logger.log("%sをカレンダーへ登録しました", bookingNumber);
        }
      }
      // 5.メールを既読にする
      thread.markRead();
    }
  } catch(e) {
    writeToLogFile("処理中にエラーが発生しました", logFile);
  } finally {
    Logger.log("%s件の予定を登録しました", createCount);
    Logger.log("カレンダー登録関数 終了");
    writeToLogFile(Logger.getLog(), logFile);
  }
}

/** ログフォルダを取得または作成する関数 */
function prepareLogFolder() {
  const folders = DriveApp.getFoldersByName('テストログ'); // フォルダの名前を指定して検索
  const currentDate = getCurrentDateAsString(); // 現在日の取得 YYYYMMDD
  let folder;
  // ログフォルダが存在しない場合は新規作成を行う
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder('テストログ');
  }
  return folder;
}

/** ログファイルを取得または作成する関数 */
function prepareLogFile(folder) {
  const currentDate = getCurrentDateAsString(); // 現在日の取得 YYYYMMDD
  const files = folder.getFilesByName("log_${currentDate}.txt");
  let logFile;
  // ログフォルダが存在しない場合は新規作成を行う
  if (files.hasNext()) {
    logFile = files.next();
  } else {
    logFile = folder.createFile(`log_${currentDate}.txt`, ""); // ファイル名を指定
  }
  return logFile;
}

/** 現在日付をYYYYMMDD形式の文字列で取得する関数 */
function getCurrentDateAsString() {
  const currentDate = new Date();
  // 年、月、日を取得
  const year = currentDate.getFullYear();
  const month = (currentDate.getMonth() + 1).toString().padStart(2, '0'); // 月は0から始まるため+1
  const day = currentDate.getDate().toString().padStart(2, '0');
  // YYYYMMDD形式の文字列を組み立て
  return year + month + day;
}

/** 第一引数の配列から第二引数の文字列を抽出・成形して返却する関数 */
function extractTargetString(array, targetString) {
  // 配列から対象の文字列を含む要素の取得
  const stringElement = array.find((e) => e.includes(targetString));
  // 対象が存在する場合は、不要な文字列を削除して返却する
  return stringElement ? stringElement.substring(stringElement.indexOf(targetString)) : ""
}

/** ログファイルにログを書き込む関数 */
function writeToLogFile(logText, file) {
  var fileContent = file.getBlob().getDataAsString() + logText + '\n';
  file.setContent(fileContent);
}

/** 時間の文字列を抽出する関数 */
function extractTimeString(inputString) {
  // 正規表現を使用して時間に関する文字列を抽出
  const timeStrings = inputString.match(/\b\d{1,2}:\d{2}(\s?[APap][Mm])?\b/g);
  
  return timeStrings[0] || "";
}

/** 文字列からDateオブジェクトを作成する関数 */
function createDateObject(date, hours, minutes) {
  const year = date.substring(3, 7); // 年
  const month = Number(date.substring(8, 10)) - 1; // 月 ※Dateオブジェクトの仕様上、-1する
  const day = date.substring(11, 13); // 日
  
  return new Date(year, month, day, hours, minutes);
}
