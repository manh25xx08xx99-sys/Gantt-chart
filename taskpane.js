// ============================================================
// 工程表ツール（施工計画書 5.工程表）— Excel Add-in taskpane
// ・入力UI：Z1.html の「5.工程表（作業リスト入力）」と同じ形の表をタスクペインに表示
// ・出力：祝日を考慮したガントチャートを「ガントチャート」シートに直接書き込む
// ============================================================

// ================= 純粋ロジック（Excelに依存しない・Nodeでテスト可能） =================

var SCHEDULE_COLORS = ["#0d2f52", "#e53935", "#fb8c00", "#fdd835", "#43a047", "#1e88e5", "#8e24aa"];
var DEFAULT_BAR_COLOR = SCHEDULE_COLORS[0];
var WEEKDAY_JP = ["日", "月", "火", "水", "木", "金", "土"];
var LEGEND_TEXT = "表示部は、基本的には作業はおこないません。　※工事の進捗状況により多少の日程の前後が発生する場合があります。";

function toISODate(d){
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
}
function fromISODate(s){ // "YYYY-MM-DD" → 正午のDate（無効ならnull）
  var m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(String(s || "").trim());
  if(!m) return null;
  var d = noonDate(+m[1], +m[2] - 1, +m[3]);
  return isNaN(d) ? null : d;
}
function parseDateInput(s){ // YYYY-MM-DD または日本式 YYYY/MM/DD
  var m = /^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/.exec(String(s || "").trim());
  if(!m) return null;
  var d = noonDate(+m[1], +m[2] - 1, +m[3]);
  // Date constructor rolls invalid dates into the following month; reject them.
  if(isNaN(d) || d.getFullYear() !== +m[1] || d.getMonth() !== +m[2] - 1 || d.getDate() !== +m[3]) return null;
  return d;
}
function formatJapaneseDate(d){
  return d.getFullYear() + "/" + String(d.getMonth() + 1).padStart(2, "0") + "/" + String(d.getDate()).padStart(2, "0");
}
function formatDateInputValue(iso){
  var d = fromISODate(iso);
  return d ? formatJapaneseDate(d) : "";
}
function noonDate(y, m, day){ // m: 0-11。計算を安定させるため正午に固定する
  return new Date(y, m, day, 12, 0, 0, 0);
}
function nthMondayOfMonth(year, month, nth){ // month: 1-12
  var d = noonDate(year, month - 1, 1);
  var count = 0;
  while(true){
    if(d.getDay() === 1){ count++; if(count === nth) return new Date(d); }
    d.setDate(d.getDate() + 1);
  }
}

// ---- 日本の祝日（Z1.htmlと同じ計算式：固定日＋ハッピーマンデー＋春分/秋分近似＋振替休日＋国民の休日） ----
var jpHolidayCache = {};     // year -> Set("YYYY-MM-DD")
var jpHolidayNameCache = {}; // year -> {"YYYY-MM-DD": name}
function getJapanHolidays(year){
  if(jpHolidayCache[year]) return jpHolidayCache[year];
  var list = [];
  var add = function(m, d, name){ list.push({ date: noonDate(year, m - 1, d), name: name }); };

  add(1, 1, "元日");
  list.push({ date: nthMondayOfMonth(year, 1, 2), name: "成人の日" });
  add(2, 11, "建国記念の日");
  add(2, 23, "天皇誕生日");
  var springDay = Math.floor(20.8431 + 0.242194 * (year - 1980) - Math.floor((year - 1980) / 4));
  var autumnDay = Math.floor(23.2488 + 0.242194 * (year - 1980) - Math.floor((year - 1980) / 4));
  add(3, springDay, "春分の日");
  add(4, 29, "昭和の日");
  add(5, 3, "憲法記念日");
  add(5, 4, "みどりの日");
  add(5, 5, "こどもの日");
  list.push({ date: nthMondayOfMonth(year, 7, 3), name: "海の日" });
  add(8, 11, "山の日");
  list.push({ date: nthMondayOfMonth(year, 9, 3), name: "敬老の日" });
  add(9, autumnDay, "秋分の日");
  list.push({ date: nthMondayOfMonth(year, 10, 2), name: "スポーツの日" });
  add(11, 3, "文化の日");
  add(11, 23, "勤労感謝の日");

  var set = new Set(list.map(function(h){ return toISODate(h.date); }));
  var nameMap = {};
  list.forEach(function(h){ nameMap[toISODate(h.date)] = h.name; });
  var substitutes = [];
  list.forEach(function(h){
    if(h.date.getDay() === 0){
      var sub = new Date(h.date);
      do { sub.setDate(sub.getDate() + 1); } while(set.has(toISODate(sub)));
      substitutes.push(toISODate(sub));
    }
  });
  substitutes.forEach(function(s){ set.add(s); nameMap[s] = "振替休日"; });

  var sortedDates = Array.from(set).map(function(s){ return new Date(s + "T12:00:00"); }).sort(function(a, b){ return a - b; });
  for(var i = 0; i < sortedDates.length - 1; i++){
    var gapDays = Math.round((sortedDates[i + 1] - sortedDates[i]) / 86400000);
    if(gapDays === 2){
      var mid = new Date(sortedDates[i]);
      mid.setDate(mid.getDate() + 1);
      var midKey = toISODate(mid);
      var dow = mid.getDay();
      if(dow !== 0 && dow !== 6 && !set.has(midKey)){
        set.add(midKey);
        nameMap[midKey] = "国民の休日";
      }
    }
  }

  jpHolidayCache[year] = set;
  jpHolidayNameCache[year] = nameMap;
  return set;
}
function isJapanHoliday(d){
  return getJapanHolidays(d.getFullYear()).has(toISODate(d));
}
function isSpecialHoliday(d, specialHolidayDates){
  return Array.isArray(specialHolidayDates) && specialHolidayDates.indexOf(toISODate(d)) !== -1;
}
function isWorkingDay(d, workOnSaturday, specialHolidayDates){
  var dow = d.getDay();
  if(dow === 0) return false; // 日曜は常に休み
  if(dow === 6 && !workOnSaturday) return false; // 土曜は設定がオフなら休み
  return !isJapanHoliday(d) && !isSpecialHoliday(d, specialHolidayDates);
}
function countWorkingDays(start, end, workOnSaturday, specialHolidayDates){
  if(!start || !end || isNaN(start) || isNaN(end) || end < start) return null;
  var count = 0;
  var cur = new Date(start);
  while(cur <= end){
    if(isWorkingDay(cur, workOnSaturday, specialHolidayDates)) count++;
    cur.setDate(cur.getDate() + 1);
  }
  return count;
}
// 開始日から数えて、稼働日でN日目にあたる日付を返す（土日・祝日はスキップ）。Z1.htmlと同じ規則
function addWorkingDays(start, numDays, workOnSaturday, specialHolidayDates){
  if(!start || isNaN(start) || !numDays || numDays < 1) return null;
  var cur = new Date(start);
  var count = 0;
  var safety = 0;
  while(safety++ < 3660){
    if(isWorkingDay(cur, workOnSaturday, specialHolidayDates)){
      count++;
      if(count === numDays) return new Date(cur);
    }
    cur.setDate(cur.getDate() + 1);
  }
  return null;
}

// ---- 入力値の正規化 ----
function normColor(v){
  if(v === null || v === undefined) return DEFAULT_BAR_COLOR;
  var s = String(v).trim();
  if(!/^#?[0-9a-fA-F]{6}$/.test(s)) return DEFAULT_BAR_COLOR;
  return s.charAt(0) === "#" ? s.toLowerCase() : "#" + s.toLowerCase();
}
function contrastTextColor(hex){
  var r = parseInt(hex.substr(1, 2), 16), g = parseInt(hex.substr(3, 2), 16), b = parseInt(hex.substr(5, 2), 16);
  return (0.299 * r + 0.587 * g + 0.114 * b) / 255 >= 0.6 ? "#000000" : "#ffffff";
}

// ---- ガントチャートの描画モデル（HTML版renderGanttChartと同じ考え方） ----
// tasks: [{name, start:Date|null, end:Date|null, note, color}]
function buildGanttModel(tasks, opts){
  opts = opts || {};
  var workOnSaturday = opts.workOnSaturday !== false;
  var specialHolidayDates = Array.isArray(opts.specialHolidayDates) ? opts.specialHolidayDates : [];
  var hasValidDates = function(t){
    return t.start && t.end && !isNaN(t.start) && !isNaN(t.end) && t.end >= t.start;
  };
  var dated = tasks.filter(hasValidDates);
  if(dated.length === 0) return null;

  var minT = Math.min.apply(null, dated.map(function(t){ return t.start.getTime(); }));
  var maxT = Math.max.apply(null, dated.map(function(t){ return t.end.getTime(); }));
  var minD = new Date(minT), maxD = new Date(maxT);
  var periodStart = noonDate(minD.getFullYear(), minD.getMonth(), minD.getDate());
  var periodEndRaw = noonDate(maxD.getFullYear(), maxD.getMonth(), maxD.getDate());
  var periodEnd = noonDate(maxD.getFullYear(), maxD.getMonth() + 1, 0); // 最大月の月末まで表示

  var days = [];
  for(var d = new Date(periodStart); d <= periodEnd; d.setDate(d.getDate() + 1)){
    days.push(new Date(d));
  }

  var years = [];
  days.forEach(function(day, idx){
    var label = day.getFullYear() + "年";
    var last = years[years.length - 1];
    if(last && last.label === label) last.count++;
    else years.push({ label: label, startIndex: idx, count: 1 });
  });

  var months = [];
  days.forEach(function(day, idx){
    var label = (day.getMonth() + 1) + "月";
    var last = months[months.length - 1];
    if(last && last.label === label) last.count++;
    else months.push({ label: label, startIndex: idx, count: 1 });
  });

  var dayInfos = days.map(function(day){
    var dow = day.getDay();
    var special = isSpecialHoliday(day, specialHolidayDates);
    var holiday = isJapanHoliday(day) || special;
    return {
      date: day,
      dowChar: WEEKDAY_JP[dow],
      shaded: dow === 0 || holiday || (dow === 6 && !workOnSaturday),
      // 特別休業日は列の網掛けは日曜・祝日と同じ扱いにするが、文字色は黒のままにする
      textColor: special ? null : dow === 0 ? "#e53935" : dow === 6 ? "#1e6fd9" : holiday ? "#e53935" : null,
    };
  });

  var rows = tasks.map(function(t){
    var color = normColor(t.color);
    var valid = hasValidDates(t);
    var noteIndex = -1;
    var cells = days.map(function(day, idx){
      var work = !!(valid && day >= t.start && day <= t.end && isWorkingDay(day, workOnSaturday, specialHolidayDates));
      if(work && noteIndex === -1) noteIndex = idx; // 備考は最初の稼働日に1回だけ書く
      return work;
    });
    var workingDays = valid ? countWorkingDays(t.start, t.end, workOnSaturday, specialHolidayDates) : null;
    return { name: t.name || "", color: color, note: t.note || "", cells: cells, noteIndex: noteIndex, workingDays: workingDays };
  });

  return {
    periodStart: periodStart,
    periodEnd: periodEndRaw,
    years: years,
    months: months,
    days: days,
    dayInfos: dayInfos,
    rows: rows,
    legendText: LEGEND_TEXT,
    progress: buildProgressCurve(rows, days),
  };
}

// 日ごとの稼働件数から、工事全体の累計進捗率（％）を計算する。
// 全体100％＝全作業の稼働日数の合計（休みは含まない）。
// ある日の増分＝その日に稼働している作業の件数（例：2件が同時に稼働していれば2日分としてカウント）。
// 各日の累計％をその日までの増分の合計として求め、最終日にちょうど100％になる
function buildProgressCurve(rows, days){
  var dailyCount = days.map(function(_, idx){
    var n = 0;
    for(var i = 0; i < rows.length; i++) if(rows[i].cells[idx]) n++;
    return n;
  });
  var total = dailyCount.reduce(function(a, b){ return a + b; }, 0);
  var cum = 0;
  var cumulativePct = dailyCount.map(function(c){
    cum += c;
    return total > 0 ? (cum / total) * 100 : 0;
  });
  return { dailyCount: dailyCount, total: total, cumulativePct: cumulativePct };
}

// 各作業の「人工数入力行」の値（文字列や空欄を含む2次元配列, rows[i][dayIdx]）から、
// 日ごとの労務者数合計を求める。数値でないセルは0として扱う
function sumManpowerByDay(manpowerGrid, dayCount){
  var daily = new Array(dayCount).fill(0);
  manpowerGrid.forEach(function(row){
    for(var idx = 0; idx < dayCount; idx++){
      var v = parseFloat(row[idx]);
      if(isFinite(v) && v > 0) daily[idx] += v;
    }
  });
  return daily;
}

// 労務者数グラフの目盛り間隔（majorUnit）を、その日の最大人数に応じて決める。
// 最大値が10未満なら2刻み、10以上20未満なら4刻み、20以上40未満なら8刻み…と
// 上限が10→20→40→80…と倍になるたびに、刻み幅も2→4→8→16…と倍になる
function computeManpowerMajorUnit(maxValue){
  var threshold = 10, unit = 2;
  while(maxValue >= threshold){
    threshold *= 2;
    unit *= 2;
  }
  return unit;
}

// 日ごとの数値配列（労務者数の合計など）から累計％カーブを作る（buildProgressCurveの汎用版）。
// 全体100％＝dailyValuesの合計。負の値・NaNは0として扱う
function buildCumulativeCurve(dailyValues){
  var clean = dailyValues.map(function(v){ return (typeof v === "number" && isFinite(v) && v > 0) ? v : 0; });
  var total = clean.reduce(function(a, b){ return a + b; }, 0);
  var cum = 0;
  var cumulativePct = clean.map(function(v){
    cum += v;
    return total > 0 ? (cum / total) * 100 : 0;
  });
  return { dailyCount: clean, total: total, cumulativePct: cumulativePct };
}

// シート上のレイアウト（各ブロックの行位置・ピクセル座標）を計算する純粋関数。
// writeGanttSheet（新規作成）と労務者数グラフの更新処理の両方から同じ計算式を使うことで、
// レイアウトのずれによるバグを防ぐ。
// colWidths: {COL_W_NAME, COL_W_DAY} を渡すと既定値（100/18pt）の代わりに使う。
// Excelは指定したポイント数の列幅をそのまま保持するとは限らない（内部の文字幅単位に
// 丸められる）ため、実際にシートへ反映された列幅を読み直して渡すと、図形の座標が
// セルの罫線から少しずつずれていく問題を補正できる
function computeLayout(model, colWidths){
  var days = model.days;
  var rowsOut = model.rows;
  var nCols = 1 + days.length;
  var headRows = 4;                 // 年／月／日／曜日
  var manpowerMode = !!model.manpowerMode;
  var ROWS_PER_TASK = manpowerMode ? 3 : 2; // 備考＋（人工数入力）＋矢印
  var headerTop = 1;                // 行2〜5：年／月／日／曜日
  var taskTop = headerTop + headRows;
  var LEGEND_ROWS = 3;              // 行事・備考等は3行
  var LEGEND_WHITE1 = 1, LEGEND_GRAY = 2, LEGEND_WHITE2 = 1;
  var LEGEND_SAMPLE_COLS = LEGEND_WHITE1 + LEGEND_GRAY + LEGEND_WHITE2;
  var legendTop = taskTop + rowsOut.length * ROWS_PER_TASK;
  var showProgress = !!model.showProgressChart && !!model.progress;
  var PROGRESS_ROWS = 10;           // 進捗率グラフ：10行＝10％刻み
  var PROGRESS_INFO_ROWS = 1;       // 総稼働日数・1日あたり％の説明行
  var progressTop = legendTop + LEGEND_ROWS;
  var progressInfoTop = progressTop + (showProgress ? PROGRESS_ROWS : 0);
  var MANPOWER_CHART_ROWS = 6;      // 労務者数グラフの高さ（行数）
  var manpowerChartTop = progressInfoTop + (showProgress ? PROGRESS_INFO_ROWS : 0);
  var nRows = manpowerChartTop + (manpowerMode ? MANPOWER_CHART_ROWS : 0);

  var COL_W_NAME = (colWidths && colWidths.COL_W_NAME) || 100;
  var COL_W_DAY = (colWidths && colWidths.COL_W_DAY) || 18;
  var ROW_H_TITLE = 30;
  var ROW_H_HEAD = 15;
  var ROW_H_TASK = 18;
  var ROW_H_PROGRESS = 15;
  var colLeft = function(colIndex){ return colIndex === 0 ? 0 : COL_W_NAME + (colIndex - 1) * COL_W_DAY; };
  var taskRowTop = function(i){ return ROW_H_TITLE + headRows * ROW_H_HEAD + i * ROWS_PER_TASK * ROW_H_TASK; };
  var progressBlockTop = taskRowTop(rowsOut.length) + LEGEND_ROWS * ROW_H_HEAD;
  var progressChartHeight = PROGRESS_ROWS * ROW_H_PROGRESS;
  var pctToY = function(pct){
    var clamped = Math.max(0, Math.min(100, pct));
    return progressBlockTop + (1 - clamped / 100) * progressChartHeight;
  };
  var manpowerBlockTop = progressBlockTop + (showProgress ? (progressChartHeight + PROGRESS_INFO_ROWS * ROW_H_HEAD) : 0);
  var manpowerChartHeight = MANPOWER_CHART_ROWS * ROW_H_HEAD;

  return {
    days: days, rowsOut: rowsOut, nCols: nCols, headRows: headRows, manpowerMode: manpowerMode, ROWS_PER_TASK: ROWS_PER_TASK,
    headerTop: headerTop, taskTop: taskTop, LEGEND_ROWS: LEGEND_ROWS,
    LEGEND_WHITE1: LEGEND_WHITE1, LEGEND_GRAY: LEGEND_GRAY, LEGEND_WHITE2: LEGEND_WHITE2, LEGEND_SAMPLE_COLS: LEGEND_SAMPLE_COLS,
    legendTop: legendTop, showProgress: showProgress, PROGRESS_ROWS: PROGRESS_ROWS, PROGRESS_INFO_ROWS: PROGRESS_INFO_ROWS,
    progressTop: progressTop, progressInfoTop: progressInfoTop,
    MANPOWER_CHART_ROWS: MANPOWER_CHART_ROWS, manpowerChartTop: manpowerChartTop, nRows: nRows,
    COL_W_NAME: COL_W_NAME, COL_W_DAY: COL_W_DAY, ROW_H_TITLE: ROW_H_TITLE, ROW_H_HEAD: ROW_H_HEAD,
    ROW_H_TASK: ROW_H_TASK, ROW_H_PROGRESS: ROW_H_PROGRESS,
    colLeft: colLeft, taskRowTop: taskRowTop, progressBlockTop: progressBlockTop, progressChartHeight: progressChartHeight,
    pctToY: pctToY, manpowerBlockTop: manpowerBlockTop, manpowerChartHeight: manpowerChartHeight,
  };
}

// ================= 入力UI（Z1.htmlの入力表と同じ見た目・操作） =================

var SHEET_GANTT = "ガントチャート";
var HEADER_FILL = "#f5f1e9";
var SHADE_FILL = "#d9d9d9";
var GRID_LINE = "#cccccc";

var rows = [];          // [{id, name, start:"YYYY-MM-DD", end:"YYYY-MM-DD", note, color}]
var nextId = 1;
var workOnSaturday = true;
var specialHolidays = []; // 特別休業日 ["YYYY-MM-DD", ...]（日曜・祝日と同じ休み扱い）
var showProgressChart = false; // 進捗率グラフ（出来高累計％の折れ線）も出力するか
var manpowerMode = false; // 労務者数グラフ（各作業に人工数入力行を追加）を使うか

// ---- 入力表の列幅（ドラッグした列だけ幅が変わり、他の列はそのまま。ただし備考(note)列が
//      残り幅を自動で吸収するので、パネル幅を超えて画面からはみ出すことはない） ----
var COL_KEYS = ["name", "start", "end", "days", "note", "color", "del"];
var RESIZABLE_KEYS = ["name", "start", "end", "days", "color"]; // note=残り幅を吸収, del=固定
var COL_MIN_PX = { name: 40, start: 34, end: 34, days: 24, note: 55, color: 45, del: 20 };
var DEFAULT_COL_PCT = { name: 15, start: 13, end: 13, days: 8, note: 26, color: 19, del: 6 };
var colWidthsPx = null; // 未設定ならパネル幅からデフォルト比率で計算する

function setStatus(msg, isError){
  var el = document.getElementById("status");
  if(!el) return;
  el.textContent = msg;
  el.style.color = isError ? "#c0392b" : "#2e7d32";
}

// ---- 保存／復元（localStorageが使えない環境でも続行できるようにtry/catch） ----
function saveState(){
  try{
    localStorage.setItem("m5ganttRows", JSON.stringify({ rows: rows, nextId: nextId, workOnSaturday: workOnSaturday, specialHolidays: specialHolidays, showProgressChart: showProgressChart, manpowerMode: manpowerMode, colWidthsPx: colWidthsPx }));
  }catch(err){}
}
function loadState(){
  try{
    var saved = JSON.parse(localStorage.getItem("m5ganttRows") || "null");
    if(saved && Array.isArray(saved.rows)){
      rows = saved.rows;
      nextId = saved.nextId || (rows.length + 1);
      workOnSaturday = saved.workOnSaturday !== false;
      specialHolidays = Array.isArray(saved.specialHolidays) ? saved.specialHolidays : [];
      showProgressChart = !!saved.showProgressChart;
      manpowerMode = !!saved.manpowerMode;
      if(saved.colWidthsPx && COL_KEYS.every(function(k){ return typeof saved.colWidthsPx[k] === "number"; })){
        colWidthsPx = saved.colWidthsPx;
      }
      return true;
    }
  }catch(err){}
  return false;
}

// ---- 特別休業日（追加した日は日曜・祝日と同じ「休み」扱いになる） ----
function addSpecialHoliday(){
  var input = document.getElementById("specialHolidayDate");
  if(!input) return;
  // <input type="date"> の値は必ず YYYY-MM-DD 形式で返る
  var d = input.value ? fromISODate(input.value) : null;
  if(!d){
    setStatus("特別休業日の日付を指定してください。", true);
    return;
  }
  var iso = toISODate(d);
  if(specialHolidays.indexOf(iso) !== -1){
    setStatus("その日は既に特別休業日として追加されています。", true);
    return;
  }
  specialHolidays.push(iso);
  specialHolidays.sort(); // ISO文字列は辞書順＝日付順
  saveState();
  renderSpecialHolidays();
  renderRows(); // 稼働日の扱いが変わるので所要日数の表示を作り直す
  setStatus(formatJapaneseDate(d) + "（" + WEEKDAY_JP[d.getDay()] + "）を特別休業日にしました。🔄 ガントチャートを生成し直すと反映されます。");
  input.value = "";
}
function removeSpecialHoliday(iso){
  specialHolidays = specialHolidays.filter(function(x){ return x !== iso; });
  saveState();
  renderSpecialHolidays();
  renderRows();
}
function renderSpecialHolidays(){
  var list = document.getElementById("specialHolidayList");
  if(!list) return;
  if(specialHolidays.length === 0){
    list.innerHTML = '<span class="holiday-empty">まだありません。追加すると、その日は日曜・祝日と同じ休み扱いになります。</span>';
    return;
  }
  list.innerHTML = specialHolidays.map(function(iso){
    var d = fromISODate(iso);
    var label = d ? formatJapaneseDate(d) + "（" + WEEKDAY_JP[d.getDay()] + "）" : iso;
    return '<span class="holiday-chip">' + escHtml(label) +
      '<button type="button" class="holiday-del" data-iso="' + iso + '" title="この日を削除">✕</button></span>';
  }).join("");
}

function escHtml(s){
  return String(s == null ? "" : s).replace(/[&<>"']/g, function(c){
    return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
  });
}

function addRow(){
  rows.push({ id: nextId++, name: "", start: "", end: "", note: "", color: DEFAULT_BAR_COLOR });
  saveState();
  renderRows();
  var body = document.getElementById("schedBody");
  if(body){
    var last = body.querySelector("tr:last-child .name-input");
    if(last) last.focus();
  }
}
function removeRow(id){
  rows = rows.filter(function(r){ return r.id !== id; });
  saveState();
  renderRows();
}
function rowById(id){
  for(var i = 0; i < rows.length; i++) if(rows[i].id === id) return rows[i];
  return null;
}

function renderRows(){
  var body = document.getElementById("schedBody");
  if(!body) return;
  if(rows.length === 0){
    body.innerHTML = '<tr><td colspan="7" class="empty-hint">まだ作業がありません。下のボタンから追加してください。</td></tr>';
    return;
  }
  body.innerHTML = rows.map(function(r){
    var dots = SCHEDULE_COLORS.map(function(c){
      return '<span class="color-dot' + (c === r.color ? " active" : "") + '" style="background:' + c + ';" data-color="' + c + '" title="' + c + '"></span>';
    }).join("");
    return '<tr data-id="' + r.id + '">' +
      '<td><div class="name-cell-inner">' +
        '<span class="row-drag-handle" title="ドラッグで並べ替え">⠿</span>' +
        '<input type="text" class="name-input" value="' + escHtml(r.name) + '" placeholder="作業内容">' +
      '</div></td>' +
      '<td class="date-cell"><input type="date" class="start-input" value="' + escHtml(r.start) + '"></td>' +
      '<td class="date-cell"><input type="date" class="end-input" value="' + escHtml(r.end) + '"></td>' +
      '<td><input type="number" min="1" class="days-input" value="' + workingDaysOf(r) + '" placeholder="－" title="日数を入力すると、開始日から自動で終了日を計算します"></td>' +
      '<td><input type="text" class="note-input" value="' + escHtml(r.note) + '" placeholder="備考"></td>' +
      '<td class="color-cell"><div class="color-dots">' + dots + "</div></td>" +
      '<td><button type="button" class="del-btn" title="この行を削除">✕</button></td>' +
      "</tr>";
  }).join("");
}


// 行の稼働日数（表示用）。開始日・終了日が揃っていなければ空
function workingDaysOf(r){
  var s = fromISODate(r.start), e = fromISODate(r.end);
  var n = countWorkingDays(s, e, workOnSaturday, specialHolidays);
  return n === null ? "" : n;
}
function refreshDaysInput(tr){
  var r = rowById(Number(tr.dataset.id));
  if(!r) return;
  var input = tr.querySelector(".days-input");
  if(input) input.value = workingDaysOf(r);
}

function bindTableEvents(){
  var body = document.getElementById("schedBody");
  if(!body) return;

  body.addEventListener("input", function(e){
    var t = e.target;
    var tr = t.closest("tr[data-id]");
    if(!tr) return;
    var r = rowById(Number(tr.dataset.id));
    if(!r) return;
    if(t.classList.contains("name-input")) r.name = t.value;
    else if(t.classList.contains("note-input")) r.note = t.value;
    else return;
    saveState();
  });

  body.addEventListener("change", function(e){
    var t = e.target;
    var tr = t.closest("tr[data-id]");
    if(!tr) return;
    var r = rowById(Number(tr.dataset.id));
    if(!r) return;

    if(t.classList.contains("start-input") || t.classList.contains("end-input")){
      // <input type="date"> の値は必ず YYYY-MM-DD 形式で返る
      var inputDate = t.value ? fromISODate(t.value) : null;
      var dateKey = t.classList.contains("start-input") ? "start" : "end";
      r[dateKey] = inputDate ? toISODate(inputDate) : "";
      refreshDaysInput(tr);
      saveState();
    } else if(t.classList.contains("days-input")){
      // 所要日数 → 開始日から稼働日で数えて終了日を自動計算する（Z1.htmlと同じ規則）
      var n = parseInt(t.value, 10);
      var s = fromISODate(r.start);
      if(s && n > 0){
        var end = addWorkingDays(s, n, workOnSaturday, specialHolidays);
        if(end){
          r.end = toISODate(end);
          var endInput = tr.querySelector(".end-input");
          if(endInput) endInput.value = formatJapaneseDate(end);
        }
      }
      saveState();
    }
  });

  body.addEventListener("click", function(e){
    var dot = e.target.closest(".color-dot");
    var tr = e.target.closest("tr[data-id]");
    if(!tr) return;
    var r = rowById(Number(tr.dataset.id));
    if(!r) return;
    if(dot){
      r.color = dot.dataset.color;
      tr.querySelectorAll(".color-dot").forEach(function(d){
        d.classList.toggle("active", d.dataset.color === r.color);
      });
      saveState();
    } else if(e.target.closest(".del-btn")){
      removeRow(r.id);
    }
  });
}

// ---- 行のドラッグ並べ替え（マウス操作。ハンドル(⠿)からのみ開始する） ----
function moveRow(sourceId, targetId, before){
  var fromIdx = rows.findIndex(function(r){ return r.id === sourceId; });
  if(fromIdx === -1) return;
  var moved = rows.splice(fromIdx, 1)[0];
  var toIdx = rows.findIndex(function(r){ return r.id === targetId; });
  var insertIdx = toIdx === -1 ? rows.length : toIdx + (before ? 0 : 1);
  rows.splice(insertIdx, 0, moved);
  saveState();
  renderRows();
}

function bindRowDrag(){
  var body = document.getElementById("schedBody");
  if(!body) return;
  var dragTr = null, dragId = null;

  function clearDropMarkers(){
    body.querySelectorAll("tr[data-id]").forEach(function(r){
      r.classList.remove("drag-over-top", "drag-over-bottom");
    });
  }
  function rowUnderPoint(x, y){
    var el = document.elementFromPoint(x, y);
    return el ? el.closest("tr[data-id]") : null;
  }
  function onMove(e){
    if(!dragTr) return;
    var overTr = rowUnderPoint(e.clientX, e.clientY);
    clearDropMarkers();
    if(overTr && overTr !== dragTr){
      var rect = overTr.getBoundingClientRect();
      var isTop = (e.clientY - rect.top) < rect.height / 2;
      overTr.classList.add(isTop ? "drag-over-top" : "drag-over-bottom");
    }
  }
  function onUp(e){
    document.removeEventListener("mousemove", onMove);
    document.removeEventListener("mouseup", onUp);
    document.body.classList.remove("row-dragging");
    if(!dragTr){ return; }
    var overTr = rowUnderPoint(e.clientX, e.clientY);
    clearDropMarkers();
    dragTr.classList.remove("dragging");
    if(overTr && overTr !== dragTr){
      var rect = overTr.getBoundingClientRect();
      var isTop = (e.clientY - rect.top) < rect.height / 2;
      moveRow(dragId, Number(overTr.dataset.id), isTop);
    }
    dragTr = null; dragId = null;
  }

  body.addEventListener("mousedown", function(e){
    var handle = e.target.closest(".row-drag-handle");
    if(!handle) return;
    var tr = handle.closest("tr[data-id]");
    if(!tr) return;
    e.preventDefault();
    dragTr = tr;
    dragId = Number(tr.dataset.id);
    tr.classList.add("dragging");
    document.body.classList.add("row-dragging");
    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup", onUp);
  });
}

// ---- 入力表の列幅リサイズ（Excelのセル境界のようにドラッグで変更） ----
function initColumnResize(){
  var table = document.getElementById("schedTable");
  var colgroup = table && table.querySelector("colgroup");
  if(!table || !colgroup) return;
  var cols = {};
  COL_KEYS.forEach(function(k){
    cols[k] = colgroup.querySelector('col[data-col="' + k + '"]');
  });

  if(!colWidthsPx){
    var wrap0 = table.closest(".sched-wrap");
    var initWidth = (wrap0 && wrap0.clientWidth) || 340;
    colWidthsPx = {};
    COL_KEYS.forEach(function(k){
      colWidthsPx[k] = Math.round(initWidth * (DEFAULT_COL_PCT[k] / 100));
    });
  }

  function containerWidth(){
    var wrap = table.closest(".sched-wrap");
    return (wrap && wrap.clientWidth) || 340;
  }

  // 備考(note)以外の合計幅から、備考が吸収すべき残り幅を計算する。
  // 保存されていた幅が今のパネル幅に対して大きすぎる場合は、伸縮可能な列を
  // 比例縮小してでも必ずパネル幅に収める（スクロールバーを出さないため）。
  function recalcNoteWidth(){
    var cw = containerWidth();
    var fixedSum = colWidthsPx.del;
    var sumResizable = 0;
    RESIZABLE_KEYS.forEach(function(k){ sumResizable += colWidthsPx[k]; });
    var available = cw - fixedSum - COL_MIN_PX.note;
    if(available > 0 && sumResizable > available){
      var scale = available / sumResizable;
      RESIZABLE_KEYS.forEach(function(k){
        colWidthsPx[k] = Math.max(COL_MIN_PX[k], Math.round(colWidthsPx[k] * scale));
      });
      sumResizable = 0;
      RESIZABLE_KEYS.forEach(function(k){ sumResizable += colWidthsPx[k]; });
    }
    colWidthsPx.note = Math.max(COL_MIN_PX.note, cw - fixedSum - sumResizable);
  }

  function applyColWidths(){
    recalcNoteWidth();
    var total = 0;
    COL_KEYS.forEach(function(k){
      if(cols[k]) cols[k].style.width = colWidthsPx[k] + "px";
      total += colWidthsPx[k];
    });
    table.style.width = total + "px";
  }
  applyColWidths();

  table.querySelectorAll(".col-resizer").forEach(function(handle){
    var key = handle.dataset.col;
    if(RESIZABLE_KEYS.indexOf(key) === -1) return;

    handle.addEventListener("mousedown", function(e){
      e.preventDefault();
      var startX = e.clientX;
      var startWidth = colWidthsPx[key];
      var minWidth = COL_MIN_PX[key] || 24;
      handle.classList.add("dragging");
      document.body.classList.add("col-resizing");

      function onMove(ev){
        var newWidth = startWidth + (ev.clientX - startX);
        if(newWidth < minWidth) newWidth = minWidth;
        // 備考が最小幅を割り込む手前までしか広げられないようにする（画面からはみ出させない）
        var sumOthers = 0;
        COL_KEYS.forEach(function(k){
          if(k !== key && k !== "note") sumOthers += colWidthsPx[k];
        });
        var maxWidth = containerWidth() - sumOthers - COL_MIN_PX.note;
        if(newWidth > maxWidth) newWidth = Math.max(minWidth, maxWidth);
        colWidthsPx[key] = newWidth;
        applyColWidths();
      }
      function onUp(){
        document.removeEventListener("mousemove", onMove);
        document.removeEventListener("mouseup", onUp);
        handle.classList.remove("dragging");
        document.body.classList.remove("col-resizing");
        saveState();
      }
      document.addEventListener("mousemove", onMove);
      document.addEventListener("mouseup", onUp);
    });
  });
}

// ================= Excel書き出し =================

// 連続する true をランにまとめる（塗り範囲のAPI呼び出し回数を減らす）
function boolToRuns(flags){
  var runs = [];
  var i = 0;
  while(i < flags.length){
    if(flags[i]){
      var j = i;
      while(j < flags.length && flags[j]) j++;
      runs.push({ start: i, count: j - i });
      i = j;
    } else { i++; }
  }
  return runs;
}

function generateGantt(){
  if(typeof Excel === "undefined" || !Excel.run){ setStatus("Excel内で実行してください。", true); return; }
  var tasks = rows.map(function(r){
    return {
      name: r.name,
      start: fromISODate(r.start),
      end: fromISODate(r.end),
      note: r.note,
      color: r.color,
    };
  });
  var model = buildGanttModel(tasks, { workOnSaturday: workOnSaturday, specialHolidayDates: specialHolidays });
  if(!model){
    setStatus("開始日・終了日が両方入った作業がありません。", true);
    return;
  }
  model.showProgressChart = showProgressChart;
  model.manpowerMode = manpowerMode;
  var L = computeLayout(model);
  setStatus("生成中...");

  // 再生成すると旧シートは消えるので、労務者数モードのときは削除する前に
  // 人工数入力行を読み取っておき、新しいシートができたら書き戻す
  // （前回と作業リスト・日数が同じ場合のみ正しく対応する。読み取りに失敗しても
  // 生成自体は続行し、その場合は人工数なしの状態からになる）
  var savedManpowerGrid = null;

  // (1) 旧シートの削除は独立したExcel.runで試す。ブック保護・唯一のシート・
  // セル編集中などの理由で削除できないことがある（GeneralException 0xA7120001）。
  // その場合は削除を諦めて既存シートを再利用するので、ここでは失敗しても致命的ではない。
  Excel.run(function(ctx){
    var oldSheet = ctx.workbook.worksheets.getItemOrNullObject(SHEET_GANTT);
    return ctx.sync().then(function(){
      if(oldSheet.isNullObject){
        return;
      }
      if(!manpowerMode || L.rowsOut.length === 0){
        oldSheet.delete();
        return;
      }
      var blockRange = oldSheet.getRangeByIndexes(L.taskTop, 1, L.rowsOut.length * L.ROWS_PER_TASK, L.days.length);
      blockRange.load("values");
      return ctx.sync().then(function(){
        var grid = [];
        for(var i = 0; i < L.rowsOut.length; i++){
          grid.push(blockRange.values[i * L.ROWS_PER_TASK + 1]);
        }
        savedManpowerGrid = grid;
      }).catch(function(){
        savedManpowerGrid = null; // 前回と構成が違うなど、読み取れなければ諦める
      }).then(function(){
        oldSheet.delete();
      });
    });
  }).catch(function(err){
    if(typeof console !== "undefined" && console.warn) console.warn("旧シートを削除できなかったため再利用します:", err);
  }).then(function(){
    // (2) シートを改めて開き、なければ新規作成、あれば再利用して中身を全消去する
    return Excel.run(function(ctx){
      var existing = ctx.workbook.worksheets.getItemOrNullObject(SHEET_GANTT);
      return ctx.sync().then(function(){
        var sheet;
        var prep = Promise.resolve();
        if(existing.isNullObject){
          sheet = ctx.workbook.worksheets.add(SHEET_GANTT);
        }else{
          sheet = existing;
          prep = clearSheetForReuse(ctx, sheet);
        }
        return prep.then(function(){
          // writeGanttSheetは内部でsyncしながら値→書式の順に適用する。
          // 実際の列幅で補正されたレイアウトを受け取り、以降の描画にも使う
          return writeGanttSheet(ctx, model, sheet).then(function(result){
            L = result.layout;
            sheet.activate();
            return ctx.sync();
          });
        }).then(function(){
          if(!manpowerMode || !savedManpowerGrid) return;
          var dailyManpower = sumManpowerByDay(savedManpowerGrid, L.days.length);
          var curve = buildCumulativeCurve(dailyManpower);
          // 実際に数値が読み取れた場合のみ書き戻す。前回データなし・構成不一致などで
          // 全て空だった場合は何もせず、生成時の初期グラフ（作業日数ベース）を残す
          if(curve.total <= 0) return;
          // 読み取れた人工数を新しいシートに書き戻してから、
          // 進捗率グラフ（人工数基準）と労務者数グラフを描く
          L.rowsOut.forEach(function(r, i){
            var mpRow = L.taskTop + i * L.ROWS_PER_TASK + 1;
            sheet.getRangeByIndexes(mpRow, 1, 1, L.days.length).values = [savedManpowerGrid[i]];
          });
          return (L.showProgress
            ? drawProgressLineChart(ctx, sheet, L, curve.cumulativePct, curve.dailyCount)
            : Promise.resolve()
          ).then(function(){
            return drawManpowerChart(ctx, sheet, L, dailyManpower);
          }).then(function(){
            if(L.showProgress && curve.total > 0){
              var perUnitPct = 100 / curve.total;
              var infoText = "総人工数：" + curve.total + "人工　（1人工 ＝ 100％ ÷ " + curve.total + "人工 ＝ " + perUnitPct.toFixed(2) + "％）";
              sheet.getRangeByIndexes(L.progressInfoTop, 0, 1, 1).values = [[infoText]];
            }
            return ctx.sync();
          });
        });
      });
    });
  }).then(function(){
    setStatus("✅ 「ガントチャート」シートを生成しました（" + model.rows.length + "件）。");
  }).catch(function(err){
    var msg = "エラー：" + ((err && err.message) ? err.message : err);
    // err.code は文字列（例："GeneralException"）。数値のときだけ16進で出す
    if(err && typeof err.code === "number") msg += " ／code: 0x" + (err.code >>> 0).toString(16).toUpperCase();
    if(err && err.debugInfo){
      var dbg = err.debugInfo;
      if(dbg.errorLocation) msg += " ／発生場所: " + dbg.errorLocation;
      if(dbg.errorResponseCode) msg += " ／http: " + dbg.errorResponseCode;
      if(dbg.fullStatements && dbg.fullStatements.length){
        msg += " ／先頭の命令: " + dbg.fullStatements.slice(0, 3).join(" ; ");
      }
    }
    // シート操作系の失敗はセル編集中・ブック保護が多いので、対処のヒントを添える
    var loc = err && err.debugInfo ? String(err.debugInfo.errorLocation || "") : "";
    if(loc.indexOf("delete") !== -1 || loc.indexOf("add") !== -1){
      msg += " —セル編集中、またはブックの構成保護中の可能性があります。EnterやEscで編集を確定、保護を解除してからもう一度お試しください。";
    }
    setStatus(msg, true);
    if(typeof console !== "undefined" && console.error) console.error(err);
  });
}

// 削除できないシートを再利用するため、中身を空にする（結合解除→図形削除→値・書式クリア）
function clearSheetForReuse(ctx, sheet){
  var used = sheet.getUsedRangeOrNullObject(false);
  return ctx.sync().then(function(){
    if(!used.isNullObject) used.unmerge();
    return ctx.sync();
  }).then(function(){
    var shapes = sheet.shapes;
    shapes.load("items");
    return ctx.sync().then(function(){
      shapes.items.slice().forEach(function(s){ s.delete(); });
      return ctx.sync();
    });
  }).then(function(){
    if(!used.isNullObject) used.clear(Excel.ClearApplyTo.all);
    return ctx.sync();
  });
}

// 進捗率グラフ（出来高累計％の折れ線）を描く。既存の同名グラフ・終点マーカーがあれば
// 削除してから描き直すので、新規作成時にも「労務者数グラフを更新」時にも使える。
// pctArr: 日ごとの累計％、dailyCount: 日ごとの増分（最後に増分>0だった日で線を止める）
async function drawProgressLineChart(ctx, sheet, L, pctArr, dailyCount){
  var days = L.days, nCols = L.nCols, colLeft = L.colLeft, COL_W_DAY = L.COL_W_DAY,
      progressTop = L.progressTop, progressBlockTop = L.progressBlockTop,
      progressChartHeight = L.progressChartHeight, pctToY = L.pctToY;

  // 描ける有効なデータがあるかを先に確認する。ない場合は何もせず終わる
  // （＝今あるグラフ（生成時点の作業日数ベースのものなど）をそのまま残す）
  var lastWorkIdx = -1;
  for(var dw = dailyCount.length - 1; dw >= 0; dw--){
    if(dailyCount[dw] > 0){ lastWorkIdx = dw; break; }
  }
  if(lastWorkIdx < 0) return;

  var oldChart = sheet.charts.getItemOrNullObject("progress-chart");
  var oldDot = sheet.shapes.getItemOrNullObject("progress-line-end-dot");
  await ctx.sync();
  if(!oldChart.isNullObject) oldChart.delete();
  if(!oldDot.isNullObject) oldDot.delete();
  await ctx.sync();

  try{
    // 先頭に「0％」の起点を1つ追加する（実際の日数より1点多い）。
    // カテゴリ数を日数+1にして、グラフ全体を半列分左にずらすことで、
    // 起点は1列目の左端（＝マス目の左下角＝0％）に、各日の点はその日の列の
    // 右端（＝翌日との境目の角）に来るようにする。稼働最終日より後ろは空白にして、
    // 「プロットしない」設定でそこで線が自然に止まるようにする
    var helperCol = nCols + 2;
    var helperRows = lastWorkIdx + 1;
    var dataRange = sheet.getRangeByIndexes(progressTop, helperCol, days.length + 1, 1);
    dataRange.clear(Excel.ClearApplyTo.contents);
    var valuesToWrite = [[0]].concat(pctArr.slice(0, helperRows).map(function(v){ return [v]; }));
    sheet.getRangeByIndexes(progressTop, helperCol, valuesToWrite.length, 1).values = valuesToWrite;
    // 一部の行だけでなく列全体を指定して確実に非表示にする
    dataRange.getEntireColumn().columnHidden = true;

    var chartWidth = Math.max(1, (days.length + 1) * COL_W_DAY);
    var chartHeight = progressChartHeight;

    var progChart = sheet.charts.add(Excel.ChartType.line, dataRange, Excel.ChartSeriesBy.columns);
    progChart.name = "progress-chart";
    progChart.plotVisibleOnly = false; // データ列を非表示にしても描画されるようにする
    progChart.displayBlanksAs = Excel.ChartDisplayBlanksAs.notPlotted;
    progChart.left = colLeft(1) - COL_W_DAY / 2;
    progChart.top = progressBlockTop;
    progChart.width = chartWidth;
    progChart.height = chartHeight;
    progChart.title.visible = false;
    progChart.legend.visible = false;
    progChart.format.fill.clear();
    progChart.format.border.clear();

    var catAxis = progChart.axes.categoryAxis;
    catAxis.visible = false;
    catAxis.majorGridlines.visible = false;
    catAxis.minorGridlines.visible = false;

    var valAxis = progChart.axes.valueAxis;
    valAxis.minimum = 0;
    valAxis.maximum = 100;
    valAxis.visible = false;
    valAxis.majorGridlines.visible = false;
    valAxis.minorGridlines.visible = false;

    var progSeries = progChart.series.getItemAt(0);
    progSeries.name = "累計進捗率";
    progSeries.smooth = false;
    progSeries.markerStyle = Excel.ChartMarkerStyle.none;
    progSeries.format.line.color = "#1565c0";
    progSeries.format.line.weight = 2; // weightは整数のみ受け付ける

    // 軸・凡例・タイトルの非表示を確定させてから余白を計算させるため、
    // 一度syncしてからプロットエリアの位置・サイズをぴったり指定する
    await ctx.sync();
    progChart.plotArea.position = "Custom";
    progChart.plotArea.left = 0;
    progChart.plotArea.top = 0;
    progChart.plotArea.width = chartWidth;
    progChart.plotArea.height = chartHeight;
    progChart.plotArea.insideLeft = 0;
    progChart.plotArea.insideTop = 0;
    progChart.plotArea.insideWidth = chartWidth;
    progChart.plotArea.insideHeight = chartHeight;

    // 終点に赤い丸マーカーを置く（グラフの上に重ねる図形）
    var endX = colLeft(1 + lastWorkIdx) + COL_W_DAY; // その日の列の右端（マス目の角）に合わせる
    var endY = pctToY(pctArr[lastWorkIdx]);
    var DOT_R = 4;
    var dot = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.ellipse);
    dot.left = endX - DOT_R;
    dot.top = endY - DOT_R;
    dot.width = DOT_R * 2;
    dot.height = DOT_R * 2;
    dot.fill.setSolidColor("#e53935");
    dot.lineFormat.visible = false;
    dot.name = "progress-line-end-dot";
  }catch(eChart){
    if(typeof console !== "undefined" && console.warn) console.warn("進捗率グラフの作成に失敗しました:", eChart);
  }
}

// 労務者数グラフ（日ごとの人数の棒グラフ）を描く。既存の同名グラフがあれば削除してから描き直す。
// 値軸はExcelの自動スケールに任せる（最大値が入力次第で変わるため）
async function drawManpowerChart(ctx, sheet, L, dailyManpower){
  var days = L.days, nCols = L.nCols, colLeft = L.colLeft, COL_W_DAY = L.COL_W_DAY;

  // A列の結合・外枠罫線は、生成時（writeGanttSheet）だけでなく「更新」ボタンから
  // 呼ばれたときも毎回確実に適用されるよう、ここでも設定する
  var mpLabelCell = sheet.getRangeByIndexes(L.manpowerChartTop, 0, L.MANPOWER_CHART_ROWS, 1);
  mpLabelCell.merge();
  // 結合していても.valuesは結合前の行数分の配列を要求するため、先頭行だけ埋めて残りは空にする
  var mpLabelValues = [];
  for(var mpr = 0; mpr < L.MANPOWER_CHART_ROWS; mpr++) mpLabelValues.push([mpr === 0 ? "労務者数" : ""]);
  mpLabelCell.values = mpLabelValues;
  mpLabelCell.format.font.size = 9;
  mpLabelCell.format.horizontalAlignment = "Left";
  mpLabelCell.format.verticalAlignment = "Center";
  mpLabelCell.format.wrapText = true;
  ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"].forEach(function(k){
    mpLabelCell.format.borders.getItem(k).style = "Continuous";
    mpLabelCell.format.borders.getItem(k).weight = "Thin";
    mpLabelCell.format.borders.getItem(k).color = GRID_LINE;
  });
  var mpBorderRange = sheet.getRangeByIndexes(L.manpowerChartTop, 0, L.MANPOWER_CHART_ROWS, nCols);
  ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"].forEach(function(k){
    mpBorderRange.format.borders.getItem(k).style = "Continuous";
    mpBorderRange.format.borders.getItem(k).weight = "Thin";
    mpBorderRange.format.borders.getItem(k).color = GRID_LINE;
  });

  var total = dailyManpower.reduce(function(a, b){ return a + b; }, 0);
  if(total <= 0) return;

  var oldChart = sheet.charts.getItemOrNullObject("manpower-chart");
  await ctx.sync();
  if(!oldChart.isNullObject) oldChart.delete();
  await ctx.sync();

  try{
    var helperCol = nCols + 4; // 進捗率グラフの隠し列(nCols+2)と重ならないよう間隔をあける
    var dataRange = sheet.getRangeByIndexes(L.manpowerChartTop, helperCol, days.length, 1);
    dataRange.clear(Excel.ClearApplyTo.contents);
    dataRange.values = dailyManpower.map(function(v){ return [v > 0 ? v : ""]; });
    // 一部の行だけでなく列全体を指定して確実に非表示にする
    dataRange.getEntireColumn().columnHidden = true;

    // 値軸は表示しない（各棒の中に数値ラベルを直接表示するので不要）。
    // 軸を消すことで、軸ラベル用の余白がプロットエリアの下・左に残って
    // 棒の底が罫線からずれる問題も避けられる。プロットエリアは
    // チャート全体にぴったり合わせて、colLeft(1)から日数分の幅で配置する
    var chartWidth = Math.max(1, days.length * COL_W_DAY);
    var chartHeight = L.manpowerChartHeight;

    var mpChart = sheet.charts.add(Excel.ChartType.columnClustered, dataRange, Excel.ChartSeriesBy.columns);
    mpChart.name = "manpower-chart";
    mpChart.plotVisibleOnly = false;
    mpChart.left = colLeft(1);
    mpChart.top = L.manpowerBlockTop;
    mpChart.width = chartWidth;
    mpChart.height = chartHeight;
    mpChart.title.visible = false;
    mpChart.legend.visible = false;
    mpChart.format.fill.clear();
    mpChart.format.border.clear();

    var catAxis = mpChart.axes.categoryAxis;
    catAxis.visible = false;
    catAxis.majorGridlines.visible = false;
    catAxis.majorTickMark = Excel.ChartAxisTickMark.none;

    var valAxis = mpChart.axes.valueAxis;
    valAxis.visible = false;
    valAxis.minimum = 0;
    // 最大人数に応じて目盛り間隔を決める（10未満→2刻み、10〜20未満→4刻み…）
    valAxis.majorUnit = computeManpowerMajorUnit(Math.max.apply(null, dailyManpower));
    valAxis.majorGridlines.visible = true;
    valAxis.majorTickMark = Excel.ChartAxisTickMark.none;

    var mpSeries = mpChart.series.getItemAt(0);
    mpSeries.name = "労務者数";
    mpSeries.format.fill.setSolidColor("#43a047");
    mpSeries.hasDataLabels = true;
    var mpLabels = mpSeries.dataLabels;
    mpLabels.position = Excel.ChartDataLabelPosition.outsideEnd;
    mpLabels.format.font.size = 7;

    // 軸の非表示を確定させてから余白を計算させるため、一度syncしてから
    // プロットエリアをチャート全体にぴったり合わせる
    await ctx.sync();
    mpChart.plotArea.position = "Custom";
    mpChart.plotArea.left = 0;
    mpChart.plotArea.top = 0;
    mpChart.plotArea.width = chartWidth;
    mpChart.plotArea.height = chartHeight;
    mpChart.plotArea.insideLeft = 0;
    mpChart.plotArea.insideTop = 0;
    mpChart.plotArea.insideWidth = chartWidth;
    mpChart.plotArea.insideHeight = chartHeight;
  }catch(eChart){
    if(typeof console !== "undefined" && console.warn) console.warn("労務者数グラフの作成に失敗しました:", eChart);
  }
}

async function writeGanttSheet(ctx, model, sheet){
  var L = computeLayout(model);
  var days = L.days, rowsOut = L.rowsOut, nCols = L.nCols, headRows = L.headRows,
      manpowerMode = L.manpowerMode, ROWS_PER_TASK = L.ROWS_PER_TASK,
      headerTop = L.headerTop, taskTop = L.taskTop, LEGEND_ROWS = L.LEGEND_ROWS,
      LEGEND_WHITE1 = L.LEGEND_WHITE1, LEGEND_GRAY = L.LEGEND_GRAY, LEGEND_WHITE2 = L.LEGEND_WHITE2,
      LEGEND_SAMPLE_COLS = L.LEGEND_SAMPLE_COLS, legendTop = L.legendTop,
      showProgress = L.showProgress, PROGRESS_ROWS = L.PROGRESS_ROWS, PROGRESS_INFO_ROWS = L.PROGRESS_INFO_ROWS,
      progressTop = L.progressTop, progressInfoTop = L.progressInfoTop,
      MANPOWER_CHART_ROWS = L.MANPOWER_CHART_ROWS, manpowerChartTop = L.manpowerChartTop, nRows = L.nRows,
      COL_W_NAME = L.COL_W_NAME, COL_W_DAY = L.COL_W_DAY, ROW_H_TITLE = L.ROW_H_TITLE, ROW_H_HEAD = L.ROW_H_HEAD,
      ROW_H_TASK = L.ROW_H_TASK, ROW_H_PROGRESS = L.ROW_H_PROGRESS,
      colLeft = L.colLeft, taskRowTop = L.taskRowTop, progressBlockTop = L.progressBlockTop,
      progressChartHeight = L.progressChartHeight, pctToY = L.pctToY,
      manpowerBlockTop = L.manpowerBlockTop, manpowerChartHeight = L.manpowerChartHeight;

  // ---- 値を一括書き込み ----
  // Excelのrange.valuesは各行が必ずnCols個の要素を持つ矩形配列である必要があるため、
  // 一部のセルしか埋めない行はpadRowでnCols分に揃える（穴や末尾欠けをできる文字列にする）。
  var padRow = function(arr, len){
    var out = new Array(len);
    for(var i = 0; i < len; i++){
      out[i] = (arr[i] === undefined || arr[i] === null) ? "" : arr[i];
    }
    return out;
  };
  var grid = [];
  grid.push(padRow(["工程表", "工事名："], nCols));
  var yearRow = [""];
  model.years.forEach(function(y){ yearRow[1 + y.startIndex] = y.label; });
  grid.push(padRow(yearRow, nCols));
  var monthRow = [""];
  model.months.forEach(function(m){ monthRow[1 + m.startIndex] = m.label; });
  grid.push(padRow(monthRow, nCols));
  grid.push([""].concat(days.map(function(d){ return d.getDate(); })));
  grid.push([""].concat(model.dayInfos.map(function(d){ return d.dowChar; })));
  rowsOut.forEach(function(r){
    var noteLine = [r.name];
    for(var i = 0; i < days.length; i++){
      noteLine.push(r.cells[i] && i === r.noteIndex ? r.note : "");
    }
    grid.push(noteLine);
    if(manpowerMode){
      var mpLine = [""];
      for(var j = 0; j < days.length; j++) mpLine.push("");
      grid.push(mpLine);
    }
    var barLine = [""];
    for(var i2 = 0; i2 < days.length; i2++) barLine.push("");
    grid.push(barLine);
  });
  grid.push(padRow(["行事・備考等"], nCols));
  var legendLine2 = [];
  // 本文は見本セル（白1・灰2・白1）の後ろから書く。列が足りないときは先頭に書く
  if(days.length > LEGEND_SAMPLE_COLS) legendLine2[1 + LEGEND_SAMPLE_COLS] = model.legendText;
  else if(days.length > 0) legendLine2[1] = model.legendText;
  grid.push(padRow(legendLine2, nCols));
  grid.push(padRow([], nCols));
  if(showProgress){
    // 各行のA列に、その行の上端が示す％（100→10）を入れる。0％は帯の下端（ラベルなし）
    for(var pr = 0; pr < PROGRESS_ROWS; pr++){
      grid.push(padRow([(PROGRESS_ROWS - pr) * 10 + "%"], nCols));
    }
    var progTotal = model.progress.total;
    var perDayPct = progTotal > 0 ? 100 / progTotal : 0;
    var progressInfoLine = ["総稼働日数：" + progTotal + "日　（1日 ＝ 100％ ÷ " + progTotal + "日 ＝ " + perDayPct.toFixed(2) + "％）"];
    grid.push(padRow(progressInfoLine, nCols));
  }
  if(manpowerMode){
    // グラフ本体はまだ描かない（人工数がまだ入力されていないため）。
    // 行だけ確保しておき、「労務者数グラフを更新」ボタンで実際のグラフを描く
    for(var mr2 = 0; mr2 < MANPOWER_CHART_ROWS; mr2++) grid.push(padRow([], nCols));
  }

  // Range.values must receive a rectangular array with exactly the same
  // dimensions as its target range. Fail with a clear message if a future
  // layout change breaks that invariant.
  if(grid.length !== nRows || grid.some(function(row){ return row.length !== nCols; })){
    throw new Error("ガントチャートの内部データサイズが不正です。");
  }
  sheet.getRangeByIndexes(0, 0, nRows, nCols).values = grid;

  // ---- 書式 ----
  // 矢印図形・グラフをセルの上に正確に重ねるための行高・列幅・座標計算はcomputeLayoutで
  // 済んでいる（変数はファイル冒頭でLから取り出し済み）
  sheet.showGridlines = false;
  // 印刷レイアウトの既定値：横向き・A3・1ページに収める・左右中央
  sheet.pageLayout.orientation = Excel.PageOrientation.landscape;
  sheet.pageLayout.paperSize = Excel.PaperType.a3;
  sheet.pageLayout.zoom = { horizontalFitToPages: 1, verticalFitToPages: 1 };
  sheet.pageLayout.centerHorizontally = true;
  sheet.getRange("A:A").format.columnWidth = COL_W_NAME;
  sheet.getRangeByIndexes(0, 1, 1, days.length).format.columnWidth = COL_W_DAY;
  sheet.getRangeByIndexes(0, 0, 1, nCols).format.rowHeight = ROW_H_TITLE;
  sheet.getRangeByIndexes(headerTop, 0, headRows, nCols).format.rowHeight = ROW_H_HEAD;
  if(rowsOut.length > 0) sheet.getRangeByIndexes(taskTop, 0, rowsOut.length * ROWS_PER_TASK, nCols).format.rowHeight = ROW_H_TASK;
  sheet.getRangeByIndexes(legendTop, 0, LEGEND_ROWS, nCols).format.rowHeight = ROW_H_HEAD;
  if(showProgress){
    sheet.getRangeByIndexes(progressInfoTop, 0, PROGRESS_INFO_ROWS, nCols).format.rowHeight = ROW_H_HEAD;
    sheet.getRangeByIndexes(progressTop, 0, PROGRESS_ROWS, nCols).format.rowHeight = ROW_H_PROGRESS;
  }
  if(manpowerMode){
    sheet.getRangeByIndexes(manpowerChartTop, 0, MANPOWER_CHART_ROWS, nCols).format.rowHeight = ROW_H_HEAD;
  }

  // 値と行高・列幅を先に確定させてから、結合・書式・図形を適用する。
  // 同一バッチに詰めすぎると操作の適用順が崩れて結合や中央寄せが抜けることがあるため
  await ctx.sync();

  // Excelに実際に反映された列幅を読み直し、矢印などの図形の座標計算をそれに合わせて
  // 補正する（列幅は内部の文字幅単位に丸められるため、指定したポイント数とわずかに
  // ずれることがあり、そのままだと列数が多いほど図形がセルの罫線から離れていく）
  // 1行目（A1）はタイトル行として全列を1つに結合しているため、そのセルから読むと
  // 結合の影響で列幅が正しく取れないおそれがある。getEntireColumn()で列全体を
  // 指定することで、行の結合状態に関係なく列そのものの幅を読み取る
  var colAProbe = sheet.getRangeByIndexes(0, 0, 1, 1).getEntireColumn();
  var colBProbe = nCols > 1 ? sheet.getRangeByIndexes(0, 1, 1, 1).getEntireColumn() : null;
  colAProbe.format.load("columnWidth");
  if(colBProbe) colBProbe.format.load("columnWidth");
  await ctx.sync();
  var actualColWName = colAProbe.format.columnWidth || COL_W_NAME;
  var actualColWDay = colBProbe ? (colBProbe.format.columnWidth || COL_W_DAY) : COL_W_DAY;
  if(Math.abs(actualColWName - COL_W_NAME) > 0.01 || Math.abs(actualColWDay - COL_W_DAY) > 0.01){
    L = computeLayout(model, { COL_W_NAME: actualColWName, COL_W_DAY: actualColWDay });
    COL_W_NAME = L.COL_W_NAME; COL_W_DAY = L.COL_W_DAY;
    colLeft = L.colLeft; taskRowTop = L.taskRowTop;
    progressBlockTop = L.progressBlockTop; progressChartHeight = L.progressChartHeight; pctToY = L.pctToY;
    manpowerBlockTop = L.manpowerBlockTop; manpowerChartHeight = L.manpowerChartHeight;
  }

  // A1 は結合しない「工程表」セル。
  var titleRange = sheet.getRangeByIndexes(0, 0, 1, 1);
  titleRange.format.font.size = 14;
  titleRange.format.font.bold = true;
  titleRange.format.horizontalAlignment = "Left";
  titleRange.format.verticalAlignment = "Center";

  // B1以降を結合した「工事名：」欄。A1の「工程表」と同じ行に置く。
  var projectRange = sheet.getRangeByIndexes(0, 1, 1, days.length);
  if(days.length > 1) projectRange.merge();
  projectRange.format.font.size = 14;
  projectRange.format.font.bold = true;
  projectRange.format.horizontalAlignment = "Left";
  projectRange.format.verticalAlignment = "Center";

  // 左側の4行を1セルに結合し、shapeではなくセルの対角罫線で区切る。
  var headerLeft = sheet.getRangeByIndexes(headerTop, 0, headRows, 1);
  headerLeft.merge();
  var diagonalBorder = headerLeft.format.borders.getItem("DiagonalDown");
  diagonalBorder.style = "Continuous";
  diagonalBorder.weight = "Thin";
  diagonalBorder.color = GRID_LINE;

  // 年見出し：年ごとに結合・中央・太字・網掛け
  model.years.forEach(function(y){
    var yr = sheet.getRangeByIndexes(headerTop, 1 + y.startIndex, 1, y.count);
    yr.merge();
    yr.format.fill.color = HEADER_FILL;
    yr.format.font.bold = true;
    yr.format.font.size = 10;
    yr.format.horizontalAlignment = "Center";
  });

  // 月見出し：月ごとに結合・中央・太字・網掛け
  model.months.forEach(function(m){
    var mr = sheet.getRangeByIndexes(headerTop + 1, 1 + m.startIndex, 1, m.count);
    mr.merge();
    mr.format.fill.color = HEADER_FILL;
    mr.format.font.bold = true;
    mr.format.font.size = 10;
    mr.format.horizontalAlignment = "Center";
  });

  // 日・曜日ヘッダー：中央寄せ・文字色・網掛け（土日祝）
  sheet.getRangeByIndexes(headerTop + 2, 1, 2, days.length).format.font.size = 9;
  sheet.getRangeByIndexes(headerTop + 2, 1, 2, days.length).format.horizontalAlignment = "Center";
  var shadedRuns = boolToRuns(model.dayInfos.map(function(d){ return d.shaded; }));
  // 休日の列は日・曜日ヘッダーだけでなく、作業行の範囲まで縦に網掛けする
  // （実際に作業がある稼働日は、この後に描く矢印図形がセルの上に重なって見える）
  var shadeRowCount = 2 + rowsOut.length * ROWS_PER_TASK;
  shadedRuns.forEach(function(run){
    sheet.getRangeByIndexes(headerTop + 2, 1 + run.start, shadeRowCount, run.count).format.fill.color = SHADE_FILL;
  });
  model.dayInfos.forEach(function(d, idx){
    if(d.textColor){
      sheet.getRangeByIndexes(headerTop + 2, 1 + idx, 2, 1).format.font.color = d.textColor;
    }
  });

  // Header labels are textboxes. The diagonal itself is the cell border above,
  // so it cannot be rerouted or resized by Excel's connector engine.
  var diagTop = ROW_H_TITLE;
  try{
    var nameLabelBox = sheet.shapes.addTextBox("名称等");
    nameLabelBox.placement = "Absolute";
    nameLabelBox.left = 5;
    nameLabelBox.top = diagTop + 30;
    // Keep the label contained in the lower-left triangle. Zero text margins
    // below still leave enough room for all three Japanese characters.
    nameLabelBox.width = 42;
    nameLabelBox.height = 14;
    nameLabelBox.lineFormat.visible = false;
    nameLabelBox.name = "gantt-header-name-label";
    nameLabelBox.textFrame.leftMargin = 0;
    nameLabelBox.textFrame.rightMargin = 0;
    nameLabelBox.textFrame.topMargin = 0;
    nameLabelBox.textFrame.bottomMargin = 0;
    nameLabelBox.textFrame.autoSizeSetting = "AutoSizeNone";
    nameLabelBox.textFrame.textRange.font.size = 10;

    var dateLabelBox = sheet.shapes.addTextBox("年度月日");
    dateLabelBox.placement = "Absolute";
    dateLabelBox.left = COL_W_NAME - 54;
    dateLabelBox.top = diagTop + 7;
    dateLabelBox.width = 52;
    dateLabelBox.height = 14;
    dateLabelBox.lineFormat.visible = false;
    dateLabelBox.name = "gantt-header-date-label";
    dateLabelBox.textFrame.leftMargin = 0;
    dateLabelBox.textFrame.rightMargin = 0;
    dateLabelBox.textFrame.topMargin = 0;
    dateLabelBox.textFrame.bottomMargin = 0;
    dateLabelBox.textFrame.autoSizeSetting = "AutoSizeNone";
    dateLabelBox.textFrame.horizontalAlignment = "Right";
    dateLabelBox.textFrame.textRange.font.size = 10;
  }catch(eBox){}

  // 稼働日の連続区間ごとに矢印図形を描く（セルの塗りつぶしはしない）。
  // 休日を挟む場合は区間ごとに矢印が分かれる。1作業＝2行（備考・矢印）、
  // 労務者数モードでは3行（備考・人工数入力・矢印）。作業名（A列）はその行数分を縦結合する
  var ARROW_H = 8;
  var barOffset = ROWS_PER_TASK - 1; // 矢印行は常に最後の行
  rowsOut.forEach(function(r, i){
    var noteRow = taskTop + i * ROWS_PER_TASK;
    var barRow = noteRow + barOffset;

    var nameCell = sheet.getRangeByIndexes(noteRow, 0, ROWS_PER_TASK, 1);
    nameCell.merge();
    nameCell.format.horizontalAlignment = "Left";
    nameCell.format.verticalAlignment = "Center";

    if(manpowerMode){
      // 人工数入力行：数値を手入力してもらうための空欄
      var mpRow = noteRow + 1;
      var mpRange = sheet.getRangeByIndexes(mpRow, 1, 1, days.length);
      mpRange.format.horizontalAlignment = "Center";
      mpRange.format.verticalAlignment = "Bottom";
      mpRange.format.font.size = 8;
      mpRange.numberFormat = [days.map(function(){ return "0;;;"; })]; // 0は非表示、入力値だけ見せる
    }

    var top = taskRowTop(i) + barOffset * ROW_H_TASK + (ROW_H_TASK - ARROW_H) / 2;
    boolToRuns(r.cells).forEach(function(run){
      var arrow = sheet.shapes.addGeometricShape(Excel.GeometricShapeType.rightArrow);
      arrow.left = colLeft(1 + run.start);
      arrow.top = top;
      arrow.width = run.count * COL_W_DAY;
      arrow.height = ARROW_H;
      arrow.fill.setSolidColor(r.color);
      arrow.lineFormat.visible = false;
      arrow.name = "gantt-arrow-" + i + "-" + run.start;
    });
    if(r.noteIndex >= 0){
      var noteCell = sheet.getRangeByIndexes(noteRow, 1 + r.noteIndex, 1, 1);
      noteCell.format.font.size = 8;
      noteCell.format.font.bold = true;
      noteCell.format.horizontalAlignment = "Left";
      noteCell.format.verticalAlignment = "Bottom";
    }
  });

  // 凡例行（行事・備考等）：3行のボックス。A列は3行まとめて1つのセルに結合する。
  // 上段・下段は日付列全体を結合。中段は先頭の見本セル（白1・灰2・白1）を残して、
  // その後ろを結合して本文欄にする
  var legendLabel = sheet.getRangeByIndexes(legendTop, 0, LEGEND_ROWS, 1);
  legendLabel.merge();
  if(days.length > LEGEND_SAMPLE_COLS){
    sheet.getRangeByIndexes(legendTop, 1, 1, days.length).merge();
    sheet.getRangeByIndexes(legendTop + 2, 1, 1, days.length).merge();
    sheet.getRangeByIndexes(legendTop + 1, 1 + LEGEND_SAMPLE_COLS, 1, days.length - LEGEND_SAMPLE_COLS).merge();
  } else if(days.length > 1){
    sheet.getRangeByIndexes(legendTop + 1, 1, 1, days.length).merge();
  }
  var legendRange = sheet.getRangeByIndexes(legendTop, 0, LEGEND_ROWS, nCols);
  legendRange.format.font.size = 9;
  legendRange.format.horizontalAlignment = "Left";
  legendRange.format.borders.getItem("EdgeTop").style = "Continuous";
  legendRange.format.borders.getItem("EdgeTop").weight = "Medium";
  legendRange.format.borders.getItem("EdgeTop").color = "#999999";
  legendLabel.format.verticalAlignment = "Center";

  // 労務者数グラフの左（A列）を1つのセルに結合する（罫線が見えなくなり見た目もすっきりする）
  if(manpowerMode){
    sheet.getRangeByIndexes(manpowerChartTop, 0, MANPOWER_CHART_ROWS, 1).merge();
  }

  // 全体に薄い罫線（タイトル行の外枠も含める）。労務者数グラフの領域は
  // 透明なチャートの下にマス目の罫線が透けて見えてしまうため対象外にする
  var gridRowCount = manpowerMode ? manpowerChartTop : nRows;
  var gridRange = sheet.getRangeByIndexes(0, 0, gridRowCount, nCols);
  ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight", "InsideHorizontal", "InsideVertical"].forEach(function(k){
    gridRange.format.borders.getItem(k).style = "Continuous";
    gridRange.format.borders.getItem(k).weight = "Thin";
    gridRange.format.borders.getItem(k).color = GRID_LINE;
  });

  // 労務者数グラフの領域は内側の罫線なし・外枠だけ薄い灰色で囲む
  if(manpowerMode){
    var mpBorderRange = sheet.getRangeByIndexes(manpowerChartTop, 0, MANPOWER_CHART_ROWS, nCols);
    ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"].forEach(function(k){
      mpBorderRange.format.borders.getItem(k).style = "Continuous";
      mpBorderRange.format.borders.getItem(k).weight = "Thin";
      mpBorderRange.format.borders.getItem(k).color = GRID_LINE;
    });
  }

  // 全体罫線を適用した後に対角罫線を再指定して、常に1本の線として残す。
  diagonalBorder.style = "Continuous";
  diagonalBorder.weight = "Thin";
  diagonalBorder.color = GRID_LINE;

  // 同じ作業内の行同士の間の横罫線だけ消す（上の罫線設定より後に行う必要がある）
  rowsOut.forEach(function(r, i){
    var noteRow = taskTop + i * ROWS_PER_TASK;
    for(var sub = 0; sub < ROWS_PER_TASK - 1; sub++){
      sheet.getRangeByIndexes(noteRow + sub, 0, 1, nCols).format.borders.getItem("EdgeBottom").style = "None";
      sheet.getRangeByIndexes(noteRow + sub + 1, 0, 1, nCols).format.borders.getItem("EdgeTop").style = "None";
    }
  });

  // 行事・備考等：日付列の縦罫線は消す。中段（本文行）は上下の横罫線も消して、
  // 罫線のない帯にする。上段・下段との間の罫線は両側のセルで消す必要がある
  for(var lr2 = 0; lr2 < LEGEND_ROWS; lr2++){
    sheet.getRangeByIndexes(legendTop + lr2, 1, 1, days.length).format.borders.getItem("InsideVertical").style = "None";
  }
  sheet.getRangeByIndexes(legendTop, 1, 1, days.length).format.borders.getItem("EdgeBottom").style = "None";
  sheet.getRangeByIndexes(legendTop + 1, 1, 1, days.length).format.borders.getItem("EdgeTop").style = "None";
  sheet.getRangeByIndexes(legendTop + 1, 1, 1, days.length).format.borders.getItem("EdgeBottom").style = "None";
  sheet.getRangeByIndexes(legendTop + 2, 1, 1, days.length).format.borders.getItem("EdgeTop").style = "None";
  if(days.length > LEGEND_SAMPLE_COLS){
    // 見本セル（灰色）と本文欄には罫線を付けない。灰色の塗りつぶしだけを静かに見せる
    sheet.getRangeByIndexes(legendTop + 1, 1 + LEGEND_WHITE1, 1, LEGEND_GRAY).format.fill.color = SHADE_FILL;
    var legendTextRange = sheet.getRangeByIndexes(legendTop + 1, 1 + LEGEND_SAMPLE_COLS, 1, days.length - LEGEND_SAMPLE_COLS);
    legendTextRange.format.verticalAlignment = "Center";
  }

  // 進捗率グラフ（出来高累計％）：A列に％軸ラベル（各行の上端の値）、
  // 日付列の上に、日ごとの累計％の点を直線でつないだ折れ線を描く。
  // 総稼働日数・1日あたり％の説明行はグラフ本体の下に1行（全列を結合）
  if(showProgress){
    var progressInfoRange = sheet.getRangeByIndexes(progressInfoTop, 0, PROGRESS_INFO_ROWS, nCols);
    if(nCols > 1) progressInfoRange.merge();
    progressInfoRange.format.font.size = 9;
    progressInfoRange.format.font.bold = true;
    progressInfoRange.format.horizontalAlignment = "Left";
    progressInfoRange.format.verticalAlignment = "Center";

    var progressLabelRange = sheet.getRangeByIndexes(progressTop, 0, PROGRESS_ROWS, 1);
    progressLabelRange.format.font.size = 8;
    progressLabelRange.format.horizontalAlignment = "Left";
    progressLabelRange.format.verticalAlignment = "Top";

    await drawProgressLineChart(ctx, sheet, L, model.progress.cumulativePct, model.progress.dailyCount);
  }
  // 労務者数グラフ本体は、人工数の入力値が読み取れたとき（generateGantt内）に描く

  return { sheet: sheet, layout: L };
}

function fmtISO(d){
  return d.getFullYear() + "/" + (d.getMonth() + 1) + "/" + d.getDate();
}

// ---- 起動 ----
function initUI(info){
  if(info && info.host && info.host !== "Excel" && info.host !== "Workbook"){
    setStatus("このアドインはExcelで開いてください。", true);
  }
  if(!loadState()) rows = []; // 初回は空（下のボタンで追加）
  var sat = document.getElementById("satCheck");
  if(sat){
    sat.checked = workOnSaturday;
    sat.addEventListener("change", function(e){
      workOnSaturday = e.target.checked;
      saveState();
      // 稼働日の扱いが変わると自動計算の終了日も変わるので、日数表示を作り直す
      renderRows();
    });
  }
  var progCheck = document.getElementById("progressChartCheck");
  if(progCheck){
    progCheck.checked = showProgressChart;
    progCheck.addEventListener("change", function(e){
      showProgressChart = e.target.checked;
      saveState();
    });
  }
  var mpCheck = document.getElementById("manpowerModeCheck");
  if(mpCheck){
    mpCheck.checked = manpowerMode;
    mpCheck.addEventListener("change", function(e){
      manpowerMode = e.target.checked;
      saveState();
    });
  }
  var addBtn = document.getElementById("addBtn");
  if(addBtn) addBtn.addEventListener("click", addRow);
  var genBtn = document.getElementById("genBtn");
  if(genBtn) genBtn.addEventListener("click", generateGantt);
  var addHolidayBtn = document.getElementById("addHolidayBtn");
  if(addHolidayBtn) addHolidayBtn.addEventListener("click", addSpecialHoliday);
  var holidayDateInput = document.getElementById("specialHolidayDate");
  if(holidayDateInput){
    // Enterキーでも追加できるようにする（ボタンを押し忘れを防ぐ）
    holidayDateInput.addEventListener("keydown", function(e){
      if(e.key === "Enter"){ e.preventDefault(); addSpecialHoliday(); }
    });
  }
  var holidayList = document.getElementById("specialHolidayList");
  if(holidayList){
    holidayList.addEventListener("click", function(e){
      var btn = e.target.closest(".holiday-del");
      if(!btn) return;
      removeSpecialHoliday(btn.dataset.iso);
    });
  }
  bindTableEvents();
  bindRowDrag();
  renderRows();
  renderSpecialHolidays();
  initColumnResize();
}
if(typeof Office !== "undefined" && Office.onReady){
  Office.onReady(function(info){ initUI(info); });
} else if(typeof document !== "undefined"){
  if(document.readyState === "loading"){
    document.addEventListener("DOMContentLoaded", function(){ initUI(null); });
  } else {
    initUI(null);
  }
}

// ---- Nodeから純粋ロジックをテストできるようにエクスポート ----
if(typeof module !== "undefined" && module.exports){
  module.exports = {
    toISODate: toISODate,
    fromISODate: fromISODate,
    parseDateInput: parseDateInput,
    formatJapaneseDate: formatJapaneseDate,
    formatDateInputValue: formatDateInputValue,
    getJapanHolidays: getJapanHolidays,
    isJapanHoliday: isJapanHoliday,
    isWorkingDay: isWorkingDay,
    countWorkingDays: countWorkingDays,
    addWorkingDays: addWorkingDays,
    normColor: normColor,
    contrastTextColor: contrastTextColor,
    buildGanttModel: buildGanttModel,
    buildProgressCurve: buildProgressCurve,
    buildCumulativeCurve: buildCumulativeCurve,
    sumManpowerByDay: sumManpowerByDay,
    computeManpowerMajorUnit: computeManpowerMajorUnit,
    computeLayout: computeLayout,
    boolToRuns: boolToRuns,
  };
}
