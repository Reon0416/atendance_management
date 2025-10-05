
/**
 * ページを開いた時に最初に呼ばれるルートメソッド
 */
function doGet(e) {

  var selectedEmpId = e.parameter.empId;
  var selectedPage = e.parameter.page;

  if (selectedEmpId == undefined) {
    // empIdがセットされていない場合はホーム画面を表示
    return HtmlService.createTemplateFromFile("view_home")
      .evaluate().setTitle("Home");
  }

  // オーナー管理画面
  if (selectedEmpId == "manager") {
    return HtmlService.createTemplateFromFile("view_manager")
    .evaluate().setTitle("manager");
  }

  // 選択した従業員IDを後続の処理でも利用するためにPropertyに保存
  PropertiesService.getUserProperties().setProperty('selectedEmpId', selectedEmpId.toString());

  // テンプレートオブジェクトを作成
  var template = HtmlService.createTemplateFromFile("view_detail");
  // ここで変数をテンプレートに渡す
  template.selectedEmpId = selectedEmpId;

   // pageパラメータに応じた分岐処理
    if (selectedPage == "health") {
    // healthページ用のテンプレートを作成し、変数を渡す
    template = HtmlService.createTemplateFromFile("view_health");
    template.selectedEmpId = selectedEmpId;
    return template.evaluate().setTitle("Health");
  } else if (selectedPage == "level") {
    // levelページ用のテンプレートを作成し、変数を渡す
    template = HtmlService.createTemplateFromFile("view_level");
    template.selectedEmpId = selectedEmpId;
    return template.evaluate().setTitle("Level");
  } else if (selectedPage == "form") {
    // formページ用のテンプレートを作成し、変数を渡す
    template = HtmlService.createTemplateFromFile("view_form");
    template.selectedEmpId = selectedEmpId;
    return template.evaluate().setTitle("Form");
  }

  // pageパラメータがなければ詳細画面を表示
  return template.evaluate().setTitle("Detail: " + selectedEmpId.toString());
}

// いまのデプロイのWebアプリURLを取得
function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * 交代申請登録
 */
function replacementRequest(data) {
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  const replacementRequestSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('交代申請登録一覧');
  var dateStr  = data.date;
  var startStr = data.start; 
  var endStr   = data.end;

  replacementRequestSheet.appendRow([selectedEmpId, dateStr, startStr, endStr, "未承認"]);

  return "送信しました"
}

/**
 * 交代申請登録一覧から自分以外の従業員の未承認のデータを返す
 */
function unapprovedData(){
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  const replacementRequestSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('交代申請登録一覧');

  const unapprovedRequests = replacementRequestSheet.getDataRange().getValues().slice(1)
    .filter(row => {
      return row[0] != selectedEmpId && row[4] == '未承認';
    })
    .map(row => {
      const shiftDate = new Date(row[1]);
      const startTimeValue = new Date(row[2]);
      const endTimeValue = new Date(row[3]);

      const startDateTime = new Date(shiftDate.getFullYear(), shiftDate.getMonth(), shiftDate.getDate(), startTimeValue.getHours(), startTimeValue.getMinutes());
      const endDateTime = new Date(shiftDate.getFullYear(), shiftDate.getMonth(), shiftDate.getDate(), endTimeValue.getHours(), endTimeValue.getMinutes());

      return {
        employeeId: row[0],
        shiftDate: Utilities.formatDate(shiftDate, 'Asia/Tokyo', 'yyyy-MM-dd'),
        startTime: Utilities.formatDate(startDateTime, 'Asia/Tokyo', 'HH:mm'),
        endTime: Utilities.formatDate(endDateTime, 'Asia/Tokyo', 'HH:mm')
      };
    });

  return {
    data: unapprovedRequests,
    count: unapprovedRequests.length
  }
}

/**
 * 交代申請を承認し、承認済みシートに記録する関数
 */
function approveShiftChange(requestData) {
  const approverId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  const approvedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('交代承認一覧');
  
  approvedSheet.appendRow([
    approverId,
    requestData.employeeId,
    requestData.shiftDate,
    requestData.startTime,
    requestData.endTime,
    new Date(),
    "未承認"
  ]);
  
  return '申請を承認しました。';
}

/**
 * シフトごとに承認者IDをグループ化して集計する関数
 */
function groupApproversByShift() {
  const employeesArray = getEmployees();

  const employeeMap = {};
  employeesArray.forEach(emp => {
    employeeMap[emp.id] = emp.name;
  });

  const approvalSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('交代承認一覧');
  const approvalData = approvalSheet.getDataRange().getValues().slice(1).filter(row => row[6] == '未承認');
  const groupedData = {};

  approvalData.forEach(row => {
    const approverId = row[0];
    const requesterId = row[1];
    const shiftDate = Utilities.formatDate(new Date(row[2]), 'Asia/Tokyo', 'yyyy-MM-dd');
    const startTime = Utilities.formatDate(new Date(row[3]), 'Asia/Tokyo', 'HH:mm');
    const endTime = Utilities.formatDate(new Date(row[4]), 'Asia/Tokyo', 'HH:mm');
    const key = `${requesterId}|${shiftDate}|${startTime}|${endTime}`;

    if (!groupedData[key]) {
      groupedData[key] = {
        requesterId: requesterId,
        requesterName: employeeMap[requesterId] || `ID:${requesterId}`,
        shiftDate: shiftDate,
        startTime: startTime,
        endTime: endTime,
        approvers: []
      };
    }
    
    const approverName = employeeMap[approverId] || `ID:${approverId}`;

    groupedData[key].approvers.push({
      id: approverId,
      name: approverName
    });
  });

  return Object.values(groupedData);
}

/**
 * 指定されたシフトのステータスを「承認」に更新する関数
 */
function completeApproval(shiftData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 交代承認一覧：未承認→承認
  const approvalSheet = ss.getSheetByName('交代承認一覧');
  const approvalData = approvalSheet.getDataRange().getValues();

  for (let i = 1; i < approvalData.length; i++) {
    const row = approvalData[i];
    const rowRequesterId = row[1].toString();
    const rowShiftDate = Utilities.formatDate(new Date(row[2]), 'Asia/Tokyo', 'yyyy-MM-dd');
    const rowStartTime = Utilities.formatDate(new Date(row[3]), 'Asia/Tokyo', 'HH:mm');
    const rowEndTime = Utilities.formatDate(new Date(row[4]), 'Asia/Tokyo', 'HH:mm');
    
    if (
      rowRequesterId === shiftData.requesterId &&
      rowShiftDate === shiftData.shiftDate &&
      rowStartTime === shiftData.startTime &&
      rowEndTime === shiftData.endTime
    ) {
      approvalSheet.getRange(i + 1, 7).setValue('承認');
    }
  }

  // 交代申請登録一覧：未承認→承認
  const requestSheet = ss.getSheetByName('交代申請登録一覧');
  const requestData = requestSheet.getDataRange().getValues();

  for (let i = 1; i < requestData.length; i++) {
    const row = requestData[i];
    
    const rowEmployeeId = row[0].toString();
    const rowShiftDate = Utilities.formatDate(new Date(row[1]), 'Asia/Tokyo', 'yyyy-MM-dd');
    const rowStartTime = Utilities.formatDate(new Date(row[2]), 'Asia/Tokyo', 'HH:mm');
    const rowEndTime = Utilities.formatDate(new Date(row[3]), 'Asia/Tokyo', 'HH:mm');

    if (
      rowEmployeeId === shiftData.requesterId &&
      rowShiftDate === shiftData.shiftDate &&
      rowStartTime === shiftData.startTime &&
      rowEndTime === shiftData.endTime
    ) {
      requestSheet.getRange(i + 1, 5).setValue('承認');
      break;
    }
  }
  
  return '承認を完了しました';
}

/**
 * 従業員一覧
 */
function getEmployees() {
  const empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('従業員名簿');
  var last_row = empSheet.getLastRow()
  var empRange = empSheet.getRange(2, 1, last_row, 2);// シートの中のヘッダーを除く範囲を取得
  var employees = [];
  var i = 1;
  while (true) {
    var empId = empRange.getCell(i, 1).getValue();
    var empName = empRange.getCell(i, 2).getValue();
    if (empId === "") { //　値を取得できなくなったら終了
      break;
    }
    employees.push({
      'id': empId,
      'name': empName
    })
    i++
  }
  return employees
}

/**
 * 従業員情報の取得
 * ※ デバッグするときにはselectedEmpIdを存在するIDで書き換えてください
 */
function getEmployeeName() {
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId') // ※デバッグするにはこの変数を直接書き換える必要があります
  var empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('従業員名簿');
  var last_row = empSheet.getLastRow()
  var empRange = empSheet.getRange(2, 1, last_row, 2);// シートの中のヘッダーを除く範囲を取得
  var i = 1;
  var empName = ""
  while (true) {
    var id = empRange.getCell(i, 1).getValue();
    var name = empRange.getCell(i, 2).getValue();
    if (id === "") {
      break;
    }
    if (id == selectedEmpId) {
      empName = name
    }
    i++
  }

  return empName
}

/**
 * '目標記録'シートから、指定された従業員IDの最新の行データを配列として取得する
 */
function getLatestGoalRecord() {
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('目標記録');
  const data = sheet.getDataRange().getValues();

  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] == selectedEmpId) {
      return data[i];
    }
  }
  
  return null;
}

/**
 * 目標の登録
 */
function saveGoalRecord(form) {
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  var goalTexts = [form.first_text, form.second_text, form.third_text];
  var goalMoneys = [form.first_money, form.second_money, form.third_money];
  const saveGoalRecordSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('目標記録');
  var goalRow = saveGoalRecordSheet.getLastRow() + 1
  saveGoalRecordSheet.getRange(goalRow, 1).setValue(selectedEmpId);
  saveGoalRecordSheet.getRange(goalRow, 2).setValue(goalTexts[0]);
  saveGoalRecordSheet.getRange(goalRow, 3).setValue(goalMoneys[0]);
  saveGoalRecordSheet.getRange(goalRow, 4).setValue(goalTexts[1]);
  saveGoalRecordSheet.getRange(goalRow, 5).setValue(goalMoneys[1]);
  saveGoalRecordSheet.getRange(goalRow, 6).setValue(goalTexts[2]);
  saveGoalRecordSheet.getRange(goalRow, 7).setValue(goalMoneys[2]);
  return '登録しました'
}

/**
 * 勤怠情報の取得
 * 今月における今日までの勤怠情報が取得される
 */
function getTimeClocks() {
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId') // ※デバッグするにはこの変数を直接書き換える必要があります
  var timeClocksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('打刻履歴');
  var last_row = timeClocksSheet.getLastRow()
  var timeClocksRange = timeClocksSheet.getRange(2, 1, last_row, 3);// シートの中のヘッダーを除く範囲を取得
  var empTimeClocks = [];
  var i = 1;
  while (true) {
    var empId = timeClocksRange.getCell(i, 1).getValue();
    var type = timeClocksRange.getCell(i, 2).getValue();
    var datetime = timeClocksRange.getCell(i, 3).getValue();
    if (empId === "") {
      break;
    }
    if (empId == selectedEmpId) {
      empTimeClocks.push({
        'date': Utilities.formatDate(datetime, "Asia/Tokyo", "yyyy-MM-dd HH:mm"),
        'type': type
      })
    }
    i++
  }
  return empTimeClocks
}

/**
 * 勤怠情報登録
 */
function saveWorkRecord(type) {
  const spreadsheet = SpreadsheetApp.openById("1_uVRyBIjiFbI0BT1nf0ReLfd2BWZ1R9w2fI5lfpR3mE");
  const sheet = spreadsheet.getSheetByName('打刻履歴');
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId');

  const now = new Date();
  const options = { timeZone: 'Asia/Tokyo', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' };
  const formattedDate = Utilities.formatDate(now, options.timeZone, 'yyyy-MM-dd HH:mm:ss');
  
  sheet.appendRow([selectedEmpId, type, formattedDate]); // 一番下にデータを挿入する
  
  let statusMessage;
  let status;
  switch(type) {
    case '出勤':
      statusMessage = '出勤しました';
      status = '出勤中';
      break;
    case '退勤':
      statusMessage = '退勤しました';
      status = '未出勤';
      break;
    case '休憩開始':
      statusMessage = '休憩を開始しました';
      status = '休憩中';
      break;
    case '休憩終了':
      statusMessage = '休憩を終了しました';
      status = '出勤中';
      break;
    default:
      statusMessage = '記録しました';
      status = '不明';
  }

  const message = { statusMessage: statusMessage, status: status };

  return message;
}

/**
 * スプレッドシートから最新の打刻履歴を取得して状態を判断します。
 */
function getInitialStatus() {
  const spreadsheet = SpreadsheetApp.openById("1_uVRyBIjiFbI0BT1nf0ReLfd2BWZ1R9w2fI5lfpR3mE");
  const sheet = spreadsheet.getSheetByName('打刻履歴');
  const lastRow = sheet.getLastRow();
  var selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId');

  if (lastRow <= 1) {
    return '未出勤';
  }

  const dataRange = sheet.getRange(2, 1, lastRow - 1, 3);
  const data = dataRange.getValues();
  
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] == selectedEmpId) {
      const lastActionType = data[i][1];
      switch(lastActionType) {
        case '出勤':
        case '休憩終了':
          return '出勤中';
        case '休憩開始':
          return '休憩中';
        case '退勤':
          return '未出勤';
        default:
          return '未出勤';
      }
    }
  }

  return '未出勤'; 
}

/**
 * spreadSheetに保存されている指定のemployee_idの行番号を返す
 */
function getTargetEmpRowNumber(empId) {
  // 開いているシートを取得
  var sheet = SpreadsheetApp.getActiveSheet()
  // 最終行取得
  var last_row = sheet.getLastRow()
  // 2行目から最終行までの1列目(emp_id)の範囲を取得
  var data_range = sheet.getRange(1, 1, last_row, 1);
  // 該当範囲のデータを取得
  var sheetRows = data_range.getValues();
  // ループ内で検索
  for (var i = 0; i <= sheetRows.length - 1; i++) {
    var row = sheetRows[i]
    if (row[0] == empId) {
      // spread sheetの行番号は1から始まるが配列のindexは0から始まるため + 1して行番号を返す
      return i + 1;
    }
  }
  // 見つからない場合にはnullを返す
  return null
}


// オーナーのメールアドレス
const OWNER_EMAIL = 'reo040116@gmail.com';
// 警告を出すポイントの基準値
const BORDER_VALUE = 20;

/**
 * 累計ポイント超過の警告メールを送信し、記録する関数
 */
function sendAlertEmail(name, empId, points) {
  const subject = `【要注意・累計ポイント超過】${name}さんの体調について`;
  const body = `
    ${name}さんの健康チェックアンケートの【累計ポイント】が【${points}点】となり、
    設定された基準値（${BORDER_VALUE}点）を超えました。

    継続的に体調が優れない可能性があります。
    面談や個別の声かけなどの対応をご検討ください。

    詳しくはスプレッドシートの健康チェックシートを見てください。

    ※この通知をもって、${name}さんの累計ポイントは0にリセットされました。
  `;
  GmailApp.sendEmail(OWNER_EMAIL, subject, body);

  const recordSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('警告メール送信履歴');

  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  recordSheet.appendRow([empId, timestamp, points]);
}

/**
 * 健康チェックアンケートのデータを記録し、ポイントを管理する
 */
function saveHealthCheck(formObject) {
  const sheetId = '1_uVRyBIjiFbI0BT1nf0ReLfd2BWZ1R9w2fI5lfpR3mE';
  const sheetName = '健康チェックアンケート';
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName(sheetName);

  const q1Mood = formObject.q1_mood;
  const q2Severity = Number(formObject.q2_severity);
  const q3Severity = Number(formObject.q3_severity);

  const userProperties = PropertiesService.getUserProperties();
  const selectedEmpId = userProperties.getProperty('selectedEmpId') || '';
  const timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');

  // シートへ: [従業員ID, Q1(ラジオ), Q2(数値), タイムスタンプ]
  sheet.appendRow([selectedEmpId, q1Mood, q2Severity, q3Severity, timestamp]);

  // 累計ポイントは「質問2（数値）」のみで加算
  const pointStorageKey = 'points_' + selectedEmpId;
  const savedPoints = Number(userProperties.getProperty(pointStorageKey)) || 0;
  const newTotalPoints = savedPoints + q2Severity;

  var responseObject = {
    message: 'アンケートを送信しました！',
    isAlert: false,
    alertMessage: ''
  };

  if (newTotalPoints >= BORDER_VALUE) {
    const employeeName = (typeof getEmployeeName === 'function')
      ? getEmployeeName(selectedEmpId)
      : selectedEmpId;

    sendAlertEmail(employeeName, selectedEmpId, newTotalPoints);
    userProperties.setProperty(pointStorageKey, '0');

    responseObject.isAlert = true;
    responseObject.alertMessage = '【警告】体調が連続して優れていないようです。次の出勤日が近い場合は交代の申請をおすすめします。';
    responseObject.message = 'アンケートは送信されました。';
  } else {
    userProperties.setProperty(pointStorageKey, String(newTotalPoints));
  }

  return responseObject;
}

/**
 * 現在の月の給与目安を計算して返す
 */
function calculateCurrentSalary() {
  const selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  const timeClocksSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('打刻履歴');

  const regularWage = 1200;
  const lateNightWage = 1500;

  const now = new Date();
  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const allTimeClocks = timeClocksSheet.getDataRange().getValues();
  const monthlyClocks = [];
  for (let i = 1; i < allTimeClocks.length; i++) {
    if (allTimeClocks[i][0] == selectedEmpId) {
      const recordDate = new Date(allTimeClocks[i][2]);
      if (recordDate >= startOfMonth) {
        monthlyClocks.push({
          type: allTimeClocks[i][1],
          time: recordDate
        });
      }
    }
  }
  monthlyClocks.sort((a, b) => a.time - b.time);

  let totalRegularMinutes = 0;
  let totalLateNightMinutes = 0;
  let clockInTime = null;
  let breakStartTime = null;

  for (const clock of monthlyClocks) {
    if (clock.type === '出勤' && !clockInTime) {
      clockInTime = clock.time;
    } else if (clock.type === '休憩開始' && clockInTime && !breakStartTime) {
      breakStartTime = clock.time;
    } else if (clock.type === '休憩終了' && breakStartTime) {
      const breakMillis = clock.time.getTime() - breakStartTime.getTime();
      clockInTime.setTime(clockInTime.getTime() + breakMillis);
      breakStartTime = null;
    } else if (clock.type === '退勤' && clockInTime) {
      const workStartTime = clockInTime;
      const workEndTime = clock.time;

      const midnight = new Date(workStartTime);
      midnight.setDate(midnight.getDate() + 1);
      midnight.setHours(0, 0, 0, 0);

      if (workEndTime <= midnight) {
        totalRegularMinutes += (workEndTime - workStartTime) / (1000 * 60);
      } else {
        totalRegularMinutes += (midnight - workStartTime) / (1000 * 60);
        totalLateNightMinutes += (workEndTime - midnight) / (1000 * 60);
      }

      clockInTime = null;
      breakStartTime = null;
    }
  }

  const regularSalary = (totalRegularMinutes / 60) * regularWage;
  const lateNightSalary = (totalLateNightMinutes / 60) * lateNightWage;
  const totalSalary = Math.floor(regularSalary + lateNightSalary);

  return totalSalary;
}

/**
 * 目標と現在の給与にもとづく達成率データを取得する
 */
function getGoalAndSalaryData() {
  const goalRecord = getLatestGoalRecord();
  const currentSalary = calculateCurrentSalary();

  if (!goalRecord || typeof currentSalary !== 'number') {
    return null;
  }

  const goals = [
    { level: 1, text: goalRecord[1], amount: Number(goalRecord[2] || 0) },
    { level: 2, text: goalRecord[3], amount: Number(goalRecord[4] || 0) },
    { level: 3, text: goalRecord[5], amount: Number(goalRecord[6] || 0) }
  ];

  let cumulativeGoal = 0;
  const processedGoals = goals.map(goal => {
    let progressForThisLevel = 0;

    if (currentSalary > cumulativeGoal) {
      progressForThisLevel = currentSalary - cumulativeGoal;
    }
    
    let percentage = 0;
    if (goal.amount > 0) {
      percentage = Math.min(100, (progressForThisLevel / goal.amount) * 100);
    }

    cumulativeGoal += goal.amount;

    return {
      level: goal.level,
      percentage: Math.round(percentage)
    };
  });

  return processedGoals;
}

/**
 * 健康チェックアンケートの今月のデータを取得する
 */
function getHealthCheckData() {
  const selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  const sheet = SpreadsheetApp.openById('1_uVRyBIjiFbI0BT1nf0ReLfd2BWZ1R9w2fI5lfpR3mE').getSheetByName('健康チェックアンケート');
  
  const now = new Date();
  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);

  return sheet.getDataRange().getValues().slice(1)
    .filter(row => {
      return row[0] == selectedEmpId && new Date(row[4]) >= startOfMonth;
    })
    .map(row => ({
      q1Mood: row[1],
      q2Severity: row[2],
      q3Severity: row[3],
      timestamp: Utilities.formatDate(new Date(row[4]), 'Asia/Tokyo', 'yyyy/MM/dd')
    }));
}

/**
 * 警告メールの送信履歴（今月分）と件数を取得する
 */
function getAlertHistory() {
  const selectedEmpId = PropertiesService.getUserProperties().getProperty('selectedEmpId');
  const sheet = SpreadsheetApp.openById('1_uVRyBIjiFbI0BT1nf0ReLfd2BWZ1R9w2fI5lfpR3mE').getSheetByName('警告メール送信履歴');
  
  const now = new Date();
  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);

  const historyData = sheet.getDataRange().getValues().slice(1)
    .filter(row => {
      return row[0] == selectedEmpId && new Date(row[1]) >= startOfMonth;
    })
    .map(row => ({
      timestamp: Utilities.formatDate(new Date(row[1]), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
      points: row[2]
    }));

  return {
    data: historyData.reverse(),
    count: historyData.length
  };
}