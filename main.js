/********************************
 * 
 * カレンダー連携全体に関係する関数
 * 
 ********************************/

/**
 * スプレッドシート表示の際に呼出し
 */
function onOpen() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  //スプレッドシートのメニューにカスタムメニュー「カレンダー連携 > 実行」を作成
  let subMenus = [];
  subMenus.push({
    name: "実行",
    functionName: "createSchedule"  //実行で呼び出す関数を指定
  });

  ss.addMenu("カレンダー連携", subMenus);
}

/**
 * 予定を作成する
 */
function createSchedule() {
 
  // 予定の最終行を取得
  let lastRow = EVENT_SHEET.getLastRow();
  
  //予定の一覧を取得
  let contents = EVENT_SHEET.getRange(TOP_ROW, 1, EVENT_SHEET.getLastRow(), LAST_COL).getValues();
 
  // googleカレンダーの取得
  let calender = CalendarApp.getCalendarById("")
 
  //順に予定を作成（今回は正しい値が来ることを想定）
  for (i = 0; i <= lastRow - TOP_ROW; i++) {

    let status = contents[i][STATUS_CELL_NUM];
    //チェック済みは飛ばす
    if (
      status == true
    ) {
      continue;
    }
 
    // 日付
    let day = new Date(contents[i][DAY_CELL_NUM]);

    // 開始時刻と終了時刻
    let timeStr = contents[i][TIME_CELL_NUM];
    let time = timeStr.split('-');

    let startTime = time[0];
    let endTime = time[1];

    // ゲスト
    let guest = contents[i][GUEST_CELL_NUM]

    let staff = null

    if (guest == ''){
      staff = 
    }else if(guest == ''){
      staff = 
    }else if(guest == ''){
      staff = 
    }else if(guest == ''){
      staff = 
    }else{
      staff = ""
    }

    // タイトル
    let title = guest + "担当：" + contents[i][TITLE_CELL_NUM];

    // オプション
    let options = {
        timezone: 'Asia/Tokyo',
        guests: staff,IWASHITA,
        sendInvites: false
    }

    try {
      // 終了が無ければ90分で設定
      if (endTime == '') {
      //予定を作成
        // 開始日時をフォーマット
        let startDate = new Date(day);
        let startHour = startTime.split(':')[0];
        let startMinute = startTime.split(':')[1];
        startDate.setHours(startHour);
        startDate.setMinutes(startMinute);
        // 終了日時をフォーマット
        let endDate = new Date(day);
        startDate.setHours(startHour);
        startDate.setMinutes(startMinute + 60);
        // 予定を作成
        calender.createEvent(
          title,
          startDate,
          endDate,
          options
        );
        
      // 開始終了時間があれば範囲で設定
      } else {
        // 開始日時をフォーマット
        let startDate = new Date(day);
        let startHour = startTime.split(':')[0];
        let startMinute = startTime.split(':')[1];

        startDate.setHours(startHour);
        startDate.setMinutes(startMinute);

        // 終了日時をフォーマット
        let endDate = new Date(day);
        let endHour = endTime.split(':')[0];
        let endMinute = endTime.split(':')[1];

        endDate.setHours(endHour);
        endDate.setMinutes(endMinute);

        // 予定を作成
        calender.createEvent(
          title,
          startDate,
          endDate,
          options
        );
      }
 
      //無事に予定が作成されたらチェックする
      EVENT_SHEET.getRange(TOP_ROW + i, 1).check();
 
    // エラーの場合（今回はログ出力のみ）
    } catch(e) {
      Logger.log(e);
    }
    
  }
  // ブラウザへ完了通知
  Browser.msgBox("カレンダー連携完了しました");
}
