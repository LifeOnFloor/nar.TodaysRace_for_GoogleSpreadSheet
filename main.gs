/**　使い方
 * 1.スクリプトプロパティに「folderId:'スプレッドシートの保管フォルダーのID'」を追加してください。
 * 2.必要なら、initProperties関数内の定数を変更してください。
 * 3.initProperties関数を実行するか、トリガーを設定してください。
 */

/** 注意
 * 1分以内に終わらなければ自動的に終了1分後に再起動します。(2分にするとGASの制限時間6分を超えてエラーになることが稀にあったため)
 * 画面左の「トリガー」ページで再起動用のトリガーが設定されていることが確認できます。
 * 
 * トリガーは自動的に削除されます。手動で削除する必要はありません。
 */





/**
 *  トリガー
 */

function initProperties() {
  /**
   * 定期トリガーを設定するための関数
   * 必要なら、thisDate, startThisTrack, startThisRaceを変更してください
   */
  const thisDate = 0;                                        // 0：当日、1:次の日を取得
  const startThisTrack = 1;                                  // 1:競馬場リストの1番目からスタート
  const startThisRace = 1;                                   // 1:1Rからスタート

  let track_name_array = todays_track_name(thisDate);
  let yyyy = track_name_array[0];                            // 今日の西暦
  let mmdd = track_name_array[1];                            // 今日の月日
  
  let folder_date_id = dir_folder(yyyy, mmdd).getId();       // スプレッドシートの保存フォルダー

  let properties = PropertiesService.getScriptProperties();
  let init_properties = {
    'track':startThisTrack + 1,
    'startRaceIndex':startThisRace,
    'trackNameArray':JSON.stringify(track_name_array),
    'folderDateId':folder_date_id
  };
  properties.setProperties(init_properties);

  main();
}

function deleteTrigger(functionName) {
  /**
   * @param functionName - 削除したいトリガーが設定されている関数の名前
   */
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() == functionName) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}


/**
 * メイン関数
 */

function main() {
  /**
   * スプレッドシートへの書き込みを行う関数です。
   * 実行時間が1分を超えると一時中断します。
   * 中断前に状態を保存し、1分後に再開できるようトリガーを設置します。
   */
  const properties = PropertiesService.getScriptProperties();
  const trackStart = Number(properties.getProperty('track'));
  let startRaceIndex = Number(properties.getProperty('startRaceIndex'));
  const trackNameArray = JSON.parse(properties.getProperty('trackNameArray'));
  const yyyy = trackNameArray[0];
  const mmdd = trackNameArray[1];
  const folderDateId = properties.getProperty('folderDateId');
  const folderDate = DriveApp.getFolderById(folderDateId);
  const startTime = new Date().getTime();
  const isLastRace = (raceIndex) => raceIndex === 12;
  const isLastTrack = (track) => track === trackNameArray.length - 1;
  const isTimeout = (startTime) => (new Date().getTime() - startTime) / (1000 * 60) >= 1;

  for (let track = trackStart; track < trackNameArray.length; track++) {
    const trackName = trackNameArray[track];
    const trackId = track_name_to_number(trackName);

    const spd = create_spreadsheet(folderDate, trackName);
    let raceIndex = startRaceIndex > 12 ? 1 : startRaceIndex;

    for (; raceIndex < 13; raceIndex++) {                       // 1レースごとに新しいシートを挿入して書き込み
      let r = `0${raceIndex}`.slice(-2);
      let race_id = `${yyyy}${trackId}${mmdd}${r}`;

      Logger.log([race_id, trackName, raceIndex]);              // 実行ログの出力

      let race_data_info = race(race_id);                       // 出走馬の過去戦績とレース距離
      let data = race_data_info[0];
      const race_distance = race_data_info[1].replace(' ', '');
      const schedule = race_data_info[2];

      let sheet = insertOrOverwriteSheet(spd, r, raceIndex);    // 新しいシートを挿入
      
      sheet.getRange(1, 1, data.length, 20).setValues(data);    // シートに過去戦績を書き込み

      let sheetForInfo = insertOrUpdateSheet(spd, 'info', 0);   // シートにレース距離を書き込み
      if (raceIndex==1) {
        sheetForInfo.getRange(1, 1, 1, 3).setValues([['レース番号', '距離', '発走時刻']]);
      }
      sheetForInfo.getRange(raceIndex+1, 1 ,1, 3)
        .setValues([[raceIndex, race_distance, schedule]]);

      if (raceIndex==1) {                                        // 不要なシートを削除
        if (spd.getSheetByName('シート1')) {
          let unusedSheet = spd.getSheetByName('シート1');
          spd.deleteSheet(unusedSheet);
        }
      }
      
      preprocessingFilterSheet(sheet, race_distance, trackName, yyyy, mmdd);  // シートの前処理
      SpreadsheetApp.flush();
      if (!isDataExist(sheet)) {continue}
      preprocessingSortSheet(sheet);

      if (isTimeout(startTime)) {                                 // タイムアウトしたときのためにトリガーを設定
        if (isLastRace(raceIndex)) {
          if (isLastTrack(track)) {
            deleteTrigger('main');                                // 最後に不要なトリガーを削除
            return
          }
          const nextTrack = track + 1;
          const nextRaceIndex = 1;
          properties.setProperty('startRaceIndex', nextRaceIndex);
          properties.setProperty('track', nextTrack);
        } else {
          const nextTrack = track;
          const nextRaceIndex = raceIndex + 1;
          properties.setProperty('startRaceIndex', nextRaceIndex);
          properties.setProperty('track', nextTrack);
        }
        deleteTrigger('main');
        ScriptApp.newTrigger('main')
          .timeBased()
          .after(1 * 60 * 1000)
          .create();
        return
      }
    }
    startRaceIndex = 1;
  }
  deleteTrigger('main');                                          // 最後に不要なトリガーを削除
}



/**
 * シート操作
 */

function insertOrOverwriteSheet(spreadsheet, sheetName, index) {
  /**
   * @param {object} spreadsheet - スプレッドシートのオブジェクト
   * @param {str} sheetName - 挿入（上書き）するシート名
   * @param {int} index - 挿入するシートのインデックス番号
   * @returns シートのオブジェクト
   */
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();                                                // すでに同名のシートがある場合は上書きする
  } else {
    sheet = spreadsheet.insertSheet(sheetName, index);            // 新しいシートを挿入
  }
  return sheet;
}
function insertOrUpdateSheet(spreadsheet, sheetName, index) {
  /**
   * @param {object} spreadsheet - スプレッドシートのオブジェクト
   * @param {str} sheetName - 挿入（更新）するシート名
   * @param {int} index - 挿入するシートのインデックス番号
   * @returns シートのオブジェクト
   */
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) {
    return sheet;
  } else {
    return spreadsheet.insertSheet(sheetName, index);
  }
}

function preprocessingFilterSheet(sheet, race_distance, place_name, yyyy, mmdd) {
  /**
   * 開催カラム：3
   * 日付カラム：5
   * タイムカラム：7
   * 距離カラム：9
   * 上がりカラム：14
   * @param {object} sheet - シートのオブジェクト
   * @param {int} race_distance - レース距離
   * @param {str} place_name - 競馬場の名前
   */
  const date = new Date(parseInt(yyyy), parseInt(mmdd.slice(0, 2)), parseInt(mmdd.slice(-2)));
  filterByValue(sheet, 7, whenTextDoesNotContain('&nbsp;'));
  filterByValue(sheet, 14, whenTextDoesNotContain('&nbsp;'));
  filterByValue(sheet, 5, whenFormulaSatisfied(date));
  filterByValue(sheet, 9, whenTextEqualTo(race_distance));
  filterByValue(sheet, 3, whenTextEqualTo(place_name));
}

function preprocessingSortSheet(sheet) {
  /**
   * 日付カラム：5
   * タイムカラム：7
   * @param {object} sheet - シートのオブジェクト
   */
  sortSheetByValue(sheet, 7, true);
  sortSheetByValue(sheet, 5, false);
}

function isDataExist(sheet) {
  const lr = sheet.getLastRow();
  if (lr < 2) {
  return false;
  }
  const data = sheet.getRange(2, 1, lr - 1, 1)
    .getValues()
    .filter((v, i) => !sheet.isRowHiddenByFilter(i + 2));
  return data.length > 0;
}

function sortSheetByValue(sheet, column, ascending) {
  sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn())
    .sort({
      column: column,
      ascending: ascending
  });
}

function filterByValue(sheet, column, criteria) {
  let filter = sheet.getFilter();
  if (!filter) {
    filter = sheet.getDataRange().createFilter();
  }
  sheet.getFilter().setColumnFilterCriteria(column, criteria);
}
function whenTextDoesNotContain(value) {
  return SpreadsheetApp.newFilterCriteria().whenTextDoesNotContain(value).build();
}
function whenTextEqualTo(value) {
  return SpreadsheetApp.newFilterCriteria().whenTextEqualTo(value).build();
}
function whenFormulaSatisfied(date) {
  const endDate = date;
  const endyyyy = endDate.getFullYear();
  const endmm = endDate.getMonth();
  const enddd = endDate.getDate();
  date.setDate(-240);
  const startDate = date;
  const startyyyy = startDate.getFullYear();
  const startmm = startDate.getMonth();
  const startdd = startDate.getDate();

  return SpreadsheetApp.newFilterCriteria().whenFormulaSatisfied(`=AND(E:E > DATE(${startyyyy}, ${startmm}, ${startdd}), E:E < DATE(${endyyyy}, ${endmm}, ${enddd}))`)
}

/**
 * フォルダー操作
 */

function dir_folder(yyyy, mmdd) {
  /**
   * スプレッドシートを保存するフォルダーを作成する関数です。
   * create_folder関数を用いています。
   * 
   * @yyyy - 4桁西暦年
   * @mmdd - 4桁月日
   * @returns フォルダーオブジェクト
   */
  let root = DriveApp.getRootFolder();
  let folder_root = create_folder(root, 'RACE');
  let folder_year = create_folder(folder_root, yyyy);
  let folder_date = create_folder(folder_year, mmdd);

  return folder_date
}

function create_spreadsheet(parent_folder, spd_name) {
  /**
   * 親フォルダの直下に指定した名前のファイルを作る関数です。
   * 同名ファイルがある場合は新しいファイルを作りません。
   *
   * @param parent_folder - 親フォルダのオブジェクト
   * @param spd_name - 作りたいスプレッドシートの名前
   * @returns：指定した名前のスプレッドシート
   */
  
  const existingFiles = parent_folder.getFilesByName(spd_name);
  if (existingFiles.hasNext()) {                                 // 存在する場合
    const existingFile = existingFiles.next();
    return SpreadsheetApp.openById(existingFile.getId());
  } else {                                                       // 存在しない場合
    const newFile = SpreadsheetApp.create(spd_name);
    const newDriveFile = DriveApp.getFileById(newFile.getId());
    parent_folder.addFile(newDriveFile);
    return newFile;
  }
}

function create_folder(parent_folder,folder_name) {
  /**
   * 親フォルダの直下に指定した名前のフォルダを作る関数です。
   * 同名フォルダーがある場合は新しいフォルダーを作りません。
   * 
   * @param {GoogleAppsScript.Drive.Folder} parentFolder - 親フォルダ
   * @param {string} folderName - 作成するフォルダの名前
   * @returns {GoogleAppsScript.Drive.Folder} - 作成されたフォルダ
   */
  let folderIterator = parent_folder.getFoldersByName(folder_name);
  let targetFolder;
  if (folderIterator.hasNext()) {
    targetFolder = folderIterator.next();
  } else {
    targetFolder = parent_folder.createFolder(folder_name);
  }
  return targetFolder
}


/**
 * スクレイピング
 */

function todays_track_name(df_date) {
  /**
   * 今日の日付と開催場所を取得します。
   * @results 配列[yyyy, mmdd, track_name]
   */
  let date = new Date();
  date.setDate(date.getDate() + df_date);
  let yyyy = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy');
  let mmdd = Utilities.formatDate(date, 'Asia/Tokyo', 'MMdd');
  
  let source = UrlFetchApp.fetch(                                                 // ネット競馬から開催場所を取得
    `https://nar.netkeiba.com/?kaisai_date=${yyyy}${mmdd}&rf=race_list`
  ).getContentText('euc-jp');
  let track_name_array = Parser.data(source).from('<h3>').to('</h3>').iterate();

  let race_name_array = [yyyy, mmdd];                                             // 取得した開催場所と日付をまとめてリストに保存
  for (i=0; i<track_name_array.length; i++){
    let track_name = track_name_array[i];
    if (track_name.includes('R') | track_name.includes('帯広(ば)')) continue;
    race_name_array.push(track_name);
  }
  return race_name_array
}

function race(race_id) {
 /**
  * 出馬表から出走馬のリストを取得し、戦績を取得します。
  * @param race_id - レースID
  * @returns レースに出走するすべての馬のすべての戦績をまとめたデータ（ヘッダー付き）
  */
  const race_url = `https://nar.netkeiba.com/race/shutuba.html?race_id=${race_id}`;
  const source = UrlFetchApp.fetch(race_url).getContentText('euc-jp');

  const race_info = Parser.data(source)
    .from('<div class="RaceData01">').to('</div>').build();
  // 距離とコースタイプ
  const distance = Parser.data(race_info)
    .from('<span>').to('m</span>').build();
  // 発送時刻
  const myRegexp = /\d+:\d+/;
  const schedule = race_info.match(myRegexp);

  const horse_list_array = Parser.data(source)
    .from('span class="HorseName"').to('</span>').iterate();
  const data = allHorsesResults(horse_list_array);

  const header = [...sheet_header(), ...data];
  const sorted_header = sort_header(header);

  return [sorted_header, distance, schedule];
}

function allHorsesResults(horse_list_array) {
  /**
   * レースに出走する馬の名前とidと戦績を取得し、すべての出走馬のデータをまとめます。
   * @param {Array} horse_list_array - 出走馬のリスト
   * @returns {Array} すべての出走馬の名前とidと戦績がまとめられた二次元配列
   */
  const allHorsesResults = [];
  
  for (let i = 1; i < horse_list_array.length; i++) {
    const horse_list = horse_list_array[i];
    const horse = Parser.data(horse_list)
      .from('<a>').to('</a>').build();

    const horseName = Parser.data(horse)
      .from('title=').to('id').build().toString().slice(1,-2);
    const horseId = Parser.data(horse)
      .from('id=').to('>').build().toString().slice(9,-1);

    const data = horseResult(i, horseName, horseId);
    if (Array.isArray(data)) {
      allHorsesResults.push(...data);
    }
  }
  return allHorsesResults;
}

function horseResult(horse_index, horse_name, horse_id) {
  /** 
   * netkkeiba.comから競走馬の過去戦績をスクレイピングする関数です。
   * @param {int} horse_index - 出走馬の馬番
   * @param {str} horse_name - 出走馬の名前
   * @param {int} horse_id - 出走馬のコード
   * @returns {Array} すべての過去戦績がまとめられた二次元配列　||　過去戦績がなければfalse
   */
  const source = UrlFetchApp.fetch(
    `https://db.netkeiba.com/horse/${horse_id}`
  ).getContentText('euc-jp');

  if (source.includes('競走データがありません')) {
    return false;
  }

  const table = Parser.data(source).from('の競走戦績').to('</table>').build();
  const tbody = Parser.data(table).from('<tbody').to('</tbody>').build();
  const trs = Parser.data(tbody).from('<tr').to('</tr>').iterate();

  const new_table = trs.map((tr) => {
    const tds = Parser.data(tr).from('<td').to('</td>').iterate();

    const new_row_key = [
      '日付',
      '開催',
      '天気',
      'R',
      'レース名',
      '映像',
      '頭数',
      '枠番',
      '馬番',
      'オッズ',
      '人気',
      '着順',
      '騎手',
      '斤量',
      '距離',
      '馬場',
      '馬場指数',
      'タイム',
      '着差',
      'タイム指数',
      '通過',
      'ペース',
      '上り',
      '馬体重',
      '厩舎コメント',
      '備考',
      '勝ち馬(2着馬)',
      '賞金',
      'horse_id'
    ];

    const new_row = tds.reduce((acc, td, i) => {
      let data;

      switch (i) {
        case 0:
        case 12:
        case 26:
          data = Parser.data(td).from('">').to('</').build();
          break;
        case 1:
          data = Parser.data(td).from('">').to('</').build();
          data = exportStringByTrackName(data);
          break;
        case 3:
        case 6:
        case 7:
        case 8:
        case 9:
        case 18:
          data = td.slice(19);
          break;
        case 17:
          data = td.slice(19);
          data = to_seconds(data);
          break;
        case 4:
          data = Parser.data(td.slice(9)).from('">').to('</').build();
          break;
        case 5:
        case 16:
        case 19:
        case 24:
        case 25:
          data = '有料';
          break;
        case 20:
          data = convertStringToArray(td.slice(td.indexOf('>')+1));
          break;
        case 10:
        case 11:
        case 22:
          data = td.slice(td.indexOf('>')+1);
          break;
        default:
          data = td.slice(1);
      }

      return { ...acc, [new_row_key[i]]: data , 'horse_id': horse_id};
    }, { horse_index, horse_name });

    return new_row;
  });

  return new_table;
}


/**
 * スクレイピング結果の出力用
 */

function sort_header(header) {
  /**
   * sortedHeaderを編集することで、自分の見やすいようにテーブルのカラムを選択し並び替えることができます。
   * ただし、カラム番号依存の関数があるので、編集するときは以下の関数も一緒に編集してください。
   * preprocessingFilterSheet, preprocessingSortSheet, calculateRatings
   * @param {Array} header - テーブルのカラム名の配列
   * @returns {Array} 必要な項目を抽出し、見やすい順序に並び替えたカラム名の配列
   */
  let sortedHeader = header.map(record => [
    record['horse_name'],
    record['騎手'],
    record['開催'],
    record['馬場'], 
    record['日付'],    //0
    record['horse_index'],
    record['タイム'],  //2
    record['馬番'],    //3
    record['距離'],    //4
    record['通過'][0],
    record['通過'][1], //6
    record['通過'][2],
    record['通過'][3], //8
    record['上り'],
    record['ペース'],   //10
    record['着順'],
    record['着差'],     //12 
    record['頭数'],     //13
    record['R'],
    record['horse_id']  //15
  ]);
  return sortedHeader;
}

function sheet_header() {
  /**
   * @returns テーブルのヘッダー
   */
  let header = [{
    'horse_index':'馬番',
    'horse_name':'馬名',
    '日付':'日付',
    '開催':'開催',
    '天気':'天気',
    'R':'R',
    'レース名':'レース名',
    '映像':'映像',
    '頭数':'頭数',
    '枠番':'枠番',
    '馬番':'過去馬番',
    'オッズ':'オッズ',
    '人気':'人気',
    '着順':'着順',
    '騎手':'騎手',
    '斤量':'斤量',
    '距離':'距離',
    '馬場':'馬場',
    '馬場指数':'馬場指数',
    'タイム':'タイム',
    '着差':'着差',
    'タイム指数':'タイム指数',
    '通過':['通過1','通過2','通過3','通過4'],
    'ペース':'ペース',
    '上り':'上り',
    '馬体重':'馬体重',
    '厩舎コメント':'厩舎コメント',
    '備考':'備考',
    '勝ち馬(2着馬)':'勝ち馬(2着馬)',
    '賞金':'賞金',
    'horse_id': 'id'
  }];
  return header
}


/**
 * データ変換
 */

function to_seconds(timeString) {
  /**
   * @param timeString - m:ss.msの形式の文字列
   * @returns トータル秒数
   */
  const [minutes, seconds] = timeString.split(":");
  const total_seconds = (Number(minutes) * 60 + Number(seconds));
  return total_seconds;
}

function format_date(dateString) {
  /**
   * @param dateString - 日付の文字列
   * @returns 日付の文字列yyyyMMdd
   */
  const raw_date = new Date(dateString);
  const date = String(raw_date.getFullYear())
    +String(('0'+raw_date.getMonth()).slice(-2))
    +String(('0'+raw_date.getDate()).slice(-2));
  return date;
}

function convertStringToArray(text) {
  /**
   * @param text - ‐を含む文字列（通過）‘2-3-3-4’
   * @returns ‐で区切った要素の配列[2,3,3,4]
   */
  var parts = text.split("-");
  var result = [
    (parts[0] ? parseInt(parts[0]) : 0),
    (parts[1] ? parseInt(parts[1]) : 0),
    (parts[2] ? parseInt(parts[2]) : 0),
    (parts[3] ? parseInt(parts[3]) : 0)
  ];
  return result;
}

function exportStringByTrackName(text) {
  /**
   * @param text - 1阪神12など、数字を含む文字列
   * @return 阪神など、数字を除外した文字列
   */
  const regex = /\W+/;
  return text.match(regex)[0];
}

function track_name_to_number(track_name) {
  /**
   * 開催場所の名前をコード番号に変換します。
   * 開催場所に対応するコード番号が存在しない場合は'00'を返します。
   */
  const track = {
    '札幌': '01', '函館': '02', '福島': '03',
    '新潟': '04', '東京': '05', '中山': '06',
    '中京': '07', '京都': '08', '阪神': '09',
    '小倉': '10',
    '門別': '30', '北見': '31', '岩見沢': '32',
    '帯広': '33', '旭川': '34', '盛岡': '35',
    '水沢': '36', '上山': '37', '三条': '38',
    '足利': '39',
    '宇都宮': '40', '高崎': '41', '浦和': '42',
    '船橋': '43', '大井': '44', '川崎': '45',
    '金沢': '46', '笠松': '47', '名古屋': '48',
    '園田': '50', '姫路': '51', '益田': '52',
    '福山': '53', '高知': '54', '佐賀': '55',
    '荒尾': '56', '中津': '57',
    '札幌（地方競馬）': '58', '函館（地方競馬）': '59',
    '新潟（地方競馬）': '60', '中京（地方競馬）': '61',
    '帯広（ば）': '83'
  };
  return track[track_name] || '00';
}
