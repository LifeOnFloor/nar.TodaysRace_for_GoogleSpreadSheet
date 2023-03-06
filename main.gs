function main() {
  /* スクレイピングの関数をまとめた関数です。
   * これを実行すると、スクレイピング、データの整形、スプレッドシートへの書き込みを行います。
   */

  // 今日の日付と開催場所
  let track_name_array = todays_track_name();
  let yyyy = track_name_array[0];
  let mmdd = track_name_array[1];
  
  // スプレッドシートの保存フォルダー
  let folder_date_id = dir_folder(yyyy, mmdd).getId();

  // スクリプトプロパティの初期化
  let properties = PropertiesService.getScriptProperties();
  let init_properties = {'track':2, 'track_name_array':JSON.stringify(track_name_array), 'folder_date_id':folder_date_id};
  properties.setProperties(init_properties);

  // スプレッドシートを作成
  write_to_sheet();
}

function write_to_sheet() {
  /* スプレッドシートへの書き込みを行う関数です。
   * この関数は実行時間が長いので、実行時間が4分を超えると一時中断します。
   * 中断前に状態を保存し、2分後に再開できるようトリガーを設置します。
   */
  
  let properties = PropertiesService.getScriptProperties();
  
  let track_start = Number(properties.getProperty('track'));
  let track_name_array = JSON.parse(properties.getProperty('track_name_array'));
  let yyyy = track_name_array[0];
  let mmdd = track_name_array[1];
  let folder_date_id = properties.getProperty('folder_date_id');
  let folder_date = DriveApp.getFolderById(folder_date_id);
  
  let start_time = new Date().getTime();

  Logger.log(track_name_array);

  for (track=track_start; track<track_name_array.length; track++) {
    let track_name = track_name_array[track];
    let track_id = track_name_to_number(track_name);

    let spd = SpreadsheetApp.create(track_name);
    let drive_spd = DriveApp.getFileById(spd.getId());
    folder_date.addFile(drive_spd);

    // 1レースごとに新しいシートを挿入して書き込み
    for (race_index=1; race_index<=12; race_index++) {
      let r = `00${race_index}`.slice(-2);
      let race_id = `${yyyy}${track_id}${mmdd}${r}`;

      // 出走馬の過去戦績とレース距離
      let race_data_info = race(race_id);
      let data = race_data_info[0];
      let race_distance = race_data_info[1];

      // 新しいシートを挿入
      spd.insertSheet(r, race_index-1);
      let sheet = spd.getSheets()[race_index-1];

      // シートに過去戦績を書き込み
      sheet.getRange(`A1:S${data.length}`).setValues(data);

      // シートにレース距離を書き込み
      let sheet_info = spd.getSheets()[spd.getSheets().length-1];
      sheet_info.getRange(race_index,1 ,1,2).setValues([[race_index, race_distance]]);

      // 確認用ログ
      Logger.log(`${track_name}${r}R`);
    }
    let current_time = new Date().getTime();
    let timedelta = (current_time -start_time) /(1000 *60);

    // すべての開催場のデータを取得したら終了します。
    if (track == track_name_array.length-1) {return}
    
    // 未完了の開催場があり、かつ、実行時間が4分以上の場合
    // スクリプトを保存して一時中断し、2分後に再開します。
    else if (timedelta >= 4) {
      properties.setProperty('track',track+1);
      ScriptApp.newTrigger('write_to_sheet').timeBased().after(2 * 60 * 1000).create();
      return 
    }
  }
}

function dir_folder(yyyy, mmdd) {
  /* スプレッドシートを保存するフォルダーを作成する関数です。
   * create_folder関数を用いています。
   */
  let root = DriveApp.getRootFolder();
  let folder_root = create_folder(root, 'RACE');
  let folder_year = create_folder(folder_root, yyyy);
  let folder_date = create_folder(folder_year, mmdd);

  return folder_date
}

function create_folder(parent_folder,folder_name) {
  /* 親フォルダの直下に指定した名前のフォルダーを作る関数です。
   * 同名フォルダーがある場合は新しいフォルダーを作りません。
   * 戻り値：指定した名前のフォルダー
   */

  // 指定した名前のフォルダを取得
  let folderIterator = parent_folder.getFoldersByName(folder_name);
  let targetFolder;
  if (folderIterator.hasNext()) {
    // 存在する場合
    targetFolder = folderIterator.next();
  } else {
    // 存在しない場合
    targetFolder = parent_folder.createFolder(folder_name);
  }
  return targetFolder
}

function todays_track_name() {
  /* 今日の日付と開催場所を取得します。
   * 戻り値：race_name_arrayの中身は以下のようになっています。
   * race_name_array[0] = yyyy
   * race_name_array[1] = mmdd
   * race_name_array[i](i>1) = track_name 
   */

  // 今日の日付を取得
  let date = new Date();
  let yyyy = Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy');
  let mmdd = Utilities.formatDate(date, 'Asia/Tokyo', 'MMdd');
  
  // ネット競馬から開催場所を取得
  let source = UrlFetchApp.fetch(`https://nar.netkeiba.com/?kaisai_date=${yyyy}${mmdd}&rf=race_list`).getContentText('euc-jp');
  let track_name_array = Parser.data(source).from('<h3>').to('</h3>').iterate();

  // 取得した開催場所と日付をまとめてリストに保存
  let race_name_array = [yyyy, mmdd];
  for (i=0; i<track_name_array.length; i++){
    let track_name = track_name_array[i];
    if (track_name.includes('R') | track_name.includes('帯広(ば)')) continue;
    race_name_array.push(track_name);
  }
  return race_name_array
}

function track_name_to_number(track_name) {
  /* 開催場所の名前をコード番号に変換します。
   * たぶん、trackに含まれない開催場所は'00'と出力されるはずです。
   * （要検討）中央の開催場所は.includes()を使わないと抽出できないかもしれません。
   */
  let track = {'札幌':'01','函館':'02', '福島':'03', '新潟':'04', '東京':'05', '中山':'06', '中京':'07', '京都':'08', '阪神':'09', '小倉':'10', '門別':'30', '北見':'31', '岩見沢':'32', '帯広':'33', '旭川':'34', '盛岡':'35', '水沢':'36', '上山':'37', '三条':'38', '足利':'39', '宇都宮':'40', '高崎':'41', '浦和':'42', '船橋':'43', '大井':'44', '川崎':'45', '金沢':'46', '笠松':'47', '名古屋':'48', '園田':'50', '姫路':'51', '益田':'52', '福山':'53', '高知':'54', '佐賀':'55', '荒尾':'56', '中津':'57', '札幌（地方競馬）':'58', '函館（地方競馬）':'59', '新潟（地方競馬）':'60', '中京（地方競馬）':'61', '帯広（ば）':'83'};
  let track_id = '00';
  track_id = track[track_name];
  return track_id
}

function race(race_id) {
  /* 出馬表から出走馬のリストを取得します。
   * 出走馬の戦績をすべて取得します。
   * 戻り値：レースに出走するすべての馬のすべての戦績をまとめたデータ（ヘッダー付き）
   */
  let race_url = `https://nar.netkeiba.com/race/shutuba.html?race_id=${race_id}`;
  let source = UrlFetchApp.fetch(race_url).getContentText('euc-jp');

  let race_info = Parser.data(source).from('<div class="RaceData01">').to('</div>').build();
  let race_distance = Parser.data(race_info).from('<span>').to('m</span>').build();

  let race_data_info = [];
  let horse_list_array;
  if (source.includes('span class="HorseName')) {
    horse_list_array = Parser.data(source).from('span class="HorseName"').to('</span>').iterate();
  }
  let data = all_horses_results(horse_list_array);
  let header = sheet_header();
  header.push(...data);
  let sorted_header = sort_header(header);

  race_data_info.push(sorted_header);
  race_data_info.push(race_distance);

  return race_data_info
}

function sort_header(header) {
  /* 戦績から必要な項目を抽出し、見やすい順序に並び替えます。
   * 
   */
  let sorted_header = header.map(elm => [elm['horse_name'],elm['騎手'], elm['開催'], elm['馬場'], elm['日付'], elm['horse_index'], elm['タイム'], elm['距離'], elm['通過'][0], elm['通過'][1], elm['通過'][2], elm['通過'][3], elm['上り'], elm['ペース'], elm['着順'], elm['着差'], elm['馬番'], elm['頭数'], elm['R']]);
  return sorted_header
}

function all_horses_results(horse_list_array) {
  /* レースに出走する馬の名前とidと戦績を取得します。
   * すべての出走馬のデータをまとめます。
   * 戻り値：すべての出走馬の名前とidと戦績がまとめられた二次元配列
   */
  let all_horses_results = [];
  
  for (let i =1; i < horse_list_array.length; i++) {
    let horse_list = horse_list_array[i];
    let horse = Parser.data(horse_list).from('<a>').to('</a>').build();

    let horse_name = Parser.data(horse).from('title=').to('id').build().toString().slice(1,-2);
    let horse_id = Parser.data(horse).from('id=').to('>').build().toString().slice(9,-1);

    let data = horse_result(i, horse_name, horse_id);
    if (typeof data !== 'boolean' && Array.isArray(data)) {
      all_horses_results.push(...data);
    }
  }
  return all_horses_results
}

function horse_result(horse_index, horse_name, horse_id) {
  /* ネット競馬から戦績をスクレイピングします。
   * 戦績の各セルを1行ごとに連想配列にします
   * すべての行をnew_tableに格納します。
   * 戻り値：new_table
   */
  let source = UrlFetchApp.fetch(`https://db.netkeiba.com/horse/${horse_id}`).getContentText('euc-jp');

  // データがなければ戦績を取得しない
  if (source.includes('競走データがありません')) {return false}
 
  let table = Parser.data(source).from('の競走戦績').to('</table>').build();
  let tbody = Parser.data(table).from('<tbody').to('</tbody>').build();
  let trs = Parser.data(tbody).from('<tr').to('</tr>').iterate();

  let new_table = [];

  for (let i_trs=0; i_trs < trs.length; i_trs++) {

    let tr = trs[i_trs];

    let tds = Parser.data(tr).from('<td').to('</td>').iterate();

    let new_row = {'horse_index':horse_index, 'horse_name':horse_name};
    let new_row_key = ['日付', '開催', '天気', 'R', 'レース名', '映像', '頭数', '枠番', '馬番', 'オッズ', '人気', '着順', '騎手', '斤量', '距離', '馬場', '馬場指数', 'タイム', '着差', 'タイム指数', '通過', 'ペース', '上り', '馬体重', '厩舎コメント', '備考', '勝ち馬(2着馬)', '賞金'];

    for (let i=0; i < tds.length; i++) {
      let td = tds[i];
      let data = 0;
      if (i==0 | i==1 | i==12 | i==26) {
        data = Parser.data(td).from('">').to('</').build();
      }
      else if (i==3 | i==6 | i==7 | i==8 | i==9 | i==17 | i==18) {
        data = td.slice(19);
      }
      else if (i==4) {
        data = Parser.data(td.slice(9)).from('">').to('</').build();
      }
      else if (i==5 | i==16 | i==19 | i==24 | i==25) {
        data = '有料';
      }
      else if (i==20) {
        let pass_order = td.slice(td.indexOf('>')+1);
        data = convertStringToArray(pass_order);
      }
      else if (i==10 | i==11 | i==22) {
        data = td.slice(td.indexOf('>')+1);
      }
      else {data = td.slice(1);}
      new_row[new_row_key[i]] = data;
    }
    new_table.push(new_row);
  }
  return new_table
}

function convertStringToArray(text) {
  var parts = text.split("-");
  var result = [
    (parts[0] ? parseInt(parts[0]) : 0),
    (parts[1] ? parseInt(parts[1]) : 0),
    (parts[2] ? parseInt(parts[2]) : 0),
    (parts[3] ? parseInt(parts[3]) : 0)
  ];
  return result;
}

function sheet_header() {
  /* テーブルのヘッダーです
   * 
   */
  let header = [{'horse_index':'horse_index', 'horse_name':'horse_name', '日付':'日付', '開催':'開催', '天気':'天気', 'R':'R', 'レース名':'レース名', '映像':'映像', '頭数':'頭数', '枠番':'枠番', '馬番':'馬番', 'オッズ':'オッズ', '人気':'人気', '着順':'着順', '騎手':'騎手', '斤量':'斤量', '距離':'距離', '馬場':'馬場', '馬場指数':'馬場指数', 'タイム':'タイム', '着差':'着差', 'タイム指数':'タイム指数', '通過':['通過1','通過2','通過3','通過4'], 'ペース':'ペース', '上り':'上り', '馬体重':'馬体重', '厩舎コメント':'厩舎コメント', '備考':'備考', '勝ち馬(2着馬)':'勝ち馬(2着馬)', '賞金':'賞金'}];
  return header
}