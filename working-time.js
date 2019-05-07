function obtainCalendarData() {    
  // IDを指定してカレンダーを取得
  const firstMemberCalendar = CalendarApp.getCalendarById('first@gmail.com');
  const secondMemberCalendar = CalendarApp.getCalendarById('second@gmail.com');
  const members = {
    'ando': {
      'calender': firstMemberCalendar,
      'line': 2,
      'column': 'H'
    },
    'takatori': {
      'calender': secondMemberCalendar,
      'line': 3,
      'column': 'I' 
    }
  } 
  // 開始日と終了日を決定
  const now = Moment.moment();
  const startDate = findStartDate(now);
  const endDate = findEndDate(now);

  //書き出したいシートを取得
  var sheet = SpreadsheetApp.getActive().getSheetByName('workTime');
  
  for (member in members) {
    writeSpreadSheetByMember(startDate, endDate, sheet, members[member]['calender'], members[member]['line'], members[member]['column'], member)
  }
  sheet.getRange('A8').setValue('最終更新日時: '+ Moment.moment().format('YYYY/MM/DD HH:mm') +'(更新期間: '+ startDate.format('YYYY/MM/DD ddd') + ' - ' + endDate.add(-1, 'day').format('YYYY/MM/DD ddd') + ')')
}

/*
 * 各人のカレンダーを取得し、スプレットシートのそれぞれの行に書き出す
 */
function writeSpreadSheetByMember(startDate, endDate, sheet, calendar, line, column, member) {
  // 日毎の勤務時間
  const workingTime = 6.5;
  // 日付を範囲指定して予定を取得
  var events = calendar.getEvents(makeDate(startDate), makeDate(endDate));

  // 取得した予定
  var eventList = [];
  //繰り返す回数は予定の個数分(列を指定するためにforEachは使わない)
  for (var i=0; i < events.length; i++) {
    var title = events[i].getTitle();
    var startTime = events[i].getStartTime()     
    var endTime = events[i].getEndTime()
    var dayOfWeek = startTime.getDay();
    var duration = calculateDuration(startTime, endTime, title);
    var satus = events[i].getMyStatus()

    // 予定を連想配列の型にして配列に保存
    var event = {
      "title": title,
      "dayOfWeek": dayOfWeek,
      "startTime": startTime,
      "endTime": endTime,
      "duration": duration,
      "status": satus
    }
    eventList.push(event)
  }
  eventList = filterOnlyActiveEvent(eventList)
  // 計算されているイベントリストを人毎に1列ごとに表示 デバック用
  /*
   var k = 0;
   eventList.forEach(function(e) {
    sheet.getRange(column+1).setValue(member);
     sheet.getRange(column+(k+2)).setValue(e);
     k++;
   })
   */
  
  var eventsList = makeEventsList(eventList)  
  var workingTimeByDay = calculateWorkingTimeByDay(eventsList)
  
    
  // 祝日のリストを取得
  const japaneseCalendar = CalendarApp.getCalendarById('en.japanese#holiday@group.v.calendar.google.com');
  const japaneseHolidays = japaneseCalendar.getEvents(makeDate(startDate), makeDate(endDate));
  
  if(japaneseHolidays.length < 1) {
    for(wtbd in workingTimeByDay) {
      var timeIncludeNinus = workingTime - workingTimeByDay[wtbd]
      // マイナスになったら強制的に0にする。
      var time = timeIncludeNinus < 0 ? 0 : timeIncludeNinus
      if(wtbd == 'Tue') { 
        writeSpreadSheet(sheet, 'B', line, time)
      } else if(wtbd == 'Wed') { 
        writeSpreadSheet(sheet, 'C', line, time)
      } else if(wtbd == 'Thu') { 
        writeSpreadSheet(sheet, 'D', line, time)
      } else if(wtbd == 'Fri') {
        writeSpreadSheet(sheet, 'E', line, time)
      }
    }
  } else {
    for(wtbd in workingTimeByDay) {
      var d = 0;
      for (holiday in japaneseHolidays) {
        var holidayStart = japaneseHolidays[d].getStartTime()
        var timeIncludeNinus = workingTime - workingTimeByDay[wtbd]
        // マイナスになったら強制的に0にする。
        var time = timeIncludeNinus < 0 ? 0 : timeIncludeNinus
        if(wtbd == 'Tue' && wtbd != holidayStart.toString().slice(0,2)) { 
          writeSpreadSheet(sheet, 'B', line, time)
        } else if(wtbd == 'Wed' && wtbd != holidayStart.toString().slice(0,2)) { 
          writeSpreadSheet(sheet, 'C', line, time)
        } else if(wtbd == 'Thu' && wtbd != holidayStart.toString().slice(0,2)) { 
          writeSpreadSheet(sheet, 'D', line, time)
        } else if(wtbd == 'Fri' && wtbd != holidayStart.toString().slice(0,2)) {
          writeSpreadSheet(sheet, 'E', line, time)
        } 
      }
    }
  }
}

/*
 * momentの時間をスプレットシートの処理で使えるよにDate型に変換
 */
function makeDate(momentDate) {
  var date = new Date(momentDate.format('YYYY/MM/DD')) 
  return date
}


/*
 * start: 火曜
 * end: 金曜
 */
function findStartDate(today) {
  if(today.day() == 1) { // 月曜日
    return Moment.moment().add(1, 'day')
  } else if(today.day() == 2) { // 火曜日
    return Moment.moment()
  } else if(today.day() == 3) { // 水曜日
    return Moment.moment().add(-1, 'day')
  } else if(today.day() == 4) { // 木曜日
    return Moment.moment().add(-2, 'day')
  } else if(today.day() == 5) { // 金曜日
    return Moment.moment().add(-3, 'day')
  } else if(today.day() == 6) { // 土曜日
    return Moment.moment().add(-4, 'day')
  } else { // 日曜日
    return Moment.moment().add(-5, 'day')
  }
}

function findEndDate(today) {
  if(today.day() == 1) { // 月曜日
    return today.add(7, 'day')
  } else if(today.day() == 2) { // 火曜日
    return today.add(6, 'day')
  } else if(today.day() == 3) { // 水曜日
    return today.add(5, 'day')
  } else if(today.day() == 4) { // 木曜日
    return today.add(4, 'day')
  } else if(today.day() == 5) { // 金曜日
    return today.add(3, 'day')
  } else if(today.day() == 6) { // 土曜日
    return today.add(2, 'day')
  } else { // 日曜日
    return today.add(1, 'day')
  }
}

/*
 * 勤務時間(10:00-12:00 / 13:00-19:00)以外の予定を排除
 */
function filterOnlyActiveEvent(el) {
  var onlyActiveEventList = el.filter( function(e) {
    return (
      (e['status'] == 'YES' || e['status'] == 'OWNER' || e['status'] == null) && ( // 出席 or オーナー
        Utilities.formatDate( e['startTime'], 'Asia/Tokyo', 'HH') == 0 ||
        Utilities.formatDate( e['startTime'], 'Asia/Tokyo', 'HH') == 9 ||
        Utilities.formatDate( e['startTime'], 'Asia/Tokyo', 'HH') == 10 ||
        Utilities.formatDate( e['startTime'], 'Asia/Tokyo', 'HH') == 11 ||
        Utilities.formatDate( e['startTime'], 'Asia/Tokyo', 'HH') == 13 ||
        Utilities.formatDate( e['startTime'], 'Asia/Tokyo', 'HH') == 14 ||
        Utilities.formatDate( e['startTime'], 'Asia/Tokyo', 'HH') == 15 ||
        Utilities.formatDate( e['startTime'], 'Asia/Tokyo', 'HH') == 16 ||
        Utilities.formatDate( e['startTime'], 'Asia/Tokyo', 'HH') == 17 ||
        Utilities.formatDate( e['startTime'], 'Asia/Tokyo', 'HH') == 18
      )
    )
  })
  return onlyActiveEventList;
}

/*
 * 予定の所要時間取得
 */
function calculateDuration(start, end, t) {
  if (t.indexOf('除外したい予定') > -1){
    return 0;
  } else {
    return (end - start) / (1000 * 60 * 60);
  }
}

/*
 * 曜日毎の連想配列にする
 *   - ex.) {'Mon': [{'予定'},{'予定'}]}
 */
function makeEventsList(el) {
  var dayOfWeek;
  var eventsByDate = {'Tue': [], 'Wed': [], 'Thu': [], 'Fri': []}; // 同じ日のeventList
  for (e in el) {
    dayOfWeek = el[e]['dayOfWeek']
    if(dayOfWeek == 2) { // 火曜日の予定
      eventsByDate['Tue'].push(el[e]);
    } else if(dayOfWeek == 3) { // 水曜日の予定
      eventsByDate['Wed'].push(el[e]);
    } else if(dayOfWeek == 4) { // 木曜日の予定
      eventsByDate['Thu'].push(el[e]);
    } else if(dayOfWeek == 5) { // 金曜日の予定
      eventsByDate['Fri'].push(el[e]);
    }
  }
  return eventsByDate;
}

/*
 * 曜日毎に、タスク消費に使える時間を算出し、連想配列で返す
 */
function calculateWorkingTimeByDay(esl) {
  var result = {'Tue': 0.0, 'Wed': 0.0, 'Thu': 0.0, 'Fri': 0.0}; 
  for(key in result){    
    for(r in esl[key]) {
      result[key] += esl[key][r]['duration']
    }
  }
  return result;
}

function writeSpreadSheet(sheet, column, line, time) {
  sheet.getRange(column + line).setValue(time);
}
