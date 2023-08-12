/********************************
 * 
 * プロジェクトで使う定数の格納場所
 * 
 ********************************/

// 連携するアカウント
const G_ACCOUNT = "";  // ★★ここに連携するカレンダーのアドレスを入れる

// 招待するアカウント
const  = "";
const  = "";
const  = "";
const  = "";
  
// シート情報
const URL = '';
const EVENT_SHEET = SpreadsheetApp.openByUrl(URL).getSheetByName('');

// 読み取り範囲（表の始まり行と終わり列）
const TOP_ROW = 2;
const LAST_COL = 5;
 
// 0始まりで列を指定しておく
const STATUS_CELL_NUM = 0;
const DAY_CELL_NUM = 1;
const TIME_CELL_NUM = 2;
const GUEST_CELL_NUM= 3;
const TITLE_CELL_NUM= 4;