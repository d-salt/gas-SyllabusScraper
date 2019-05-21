function onOpen () {
  var spreadsheet = SpreadsheetApp.getActive();
  var subMenus = [
    { name: 'Get Classes', functionName: 'getClasses' },
    { name: 'Update TimeTable', functionName: 'updateTimeTable' }
  ];
  
  spreadsheet.addMenu('シラバス', subMenus)
}

function keyvaluesToObject (values) {
  var obj = {}
  values.forEach(function (row) {
    obj[row[0]] = row[1]
  });
  Logger.log(obj);
  return obj
}

function valuesToObject (values) {
  var array = []
  var header = values[0];
  values = values.slice(1);
  values.forEach(function (row) {
    var obj = {};
    row.forEach(function (value, i) {
      if (header[i] == '') return;
      obj[header[i]] = value;
    });
    array.push(obj)
  });
  return array
}

function getClasses () {
  var spreadsheet = SpreadsheetApp.getActive();
  var settingsSheet = spreadsheet.getSheetByName('Settings');
  var rawSettings = settingsSheet.getRange(1, 1, settingsSheet.getLastRow(), 2).getValues();
  var settings = keyvaluesToObject(rawSettings);
  
  var classesSheet = spreadsheet.getSheetByName('Classes');
  classesSheet.clear();
  classesSheet.getRange(1, 1).setFormula('=IMPORTXML("' + settings.url + '", "//*[@id=\'CPH1_gvw_kensaku\']/thead/tr")');
  var lastRow = 1;
  
  for (var i = 1; i <= settings.pageAmount; ++i) {
    classesSheet.getRange(lastRow + 1, 1).setFormula('=IMPORTXML("' + settings.url + '&PG=' + i + '", "//*[@id=\'CPH1_gvw_kensaku\']/tbody/tr")');
    lastRow = classesSheet.getLastRow();
  }

}

function getClassList () {
  var spreadsheet = SpreadsheetApp.getActive();
  var classesSheet = spreadsheet.getSheetByName('Classes');
  var rawClasses = classesSheet.getRange(1, 1, classesSheet.getLastRow(), classesSheet.getLastColumn()).getValues();
  var classes = valuesToObject(rawClasses);
  
  var classList = getClassListFrame();
  classes.forEach(function (class) {
    var matches = /\[.*\](..)\（(...*)\）/.exec(class['時限']);
    var semester = matches[2];
    var slot = matches[1];
    var dayOfWeek = slot.slice(0,1);
    var unit = slot.slice(1);
    
    if (classList[semester] === undefined) classList[semester] = {};

    if (/[月火水木金土]/.exec(dayOfWeek) && semester.length == 2) {
      classList[semester][dayOfWeek][unit].push(class['科目']);
    } else {
      if (classList[semester][slot] === undefined) classList[semester][slot] = [];
      classList[semester][slot].push(class['科目']);
    }
  });
  return classList;
}

function addTAList (classList) {
  var spreadsheet = SpreadsheetApp.getActive();
  var classesSheet = spreadsheet.getSheetByName('TA');
  var rawClasses = classesSheet.getRange(1, 1, classesSheet.getLastRow(), classesSheet.getLastColumn()).getValues();
  var classes = valuesToObject(rawClasses);
  
  classes.forEach(function (class) {
    var semester = class['semester'];
    var dayOfWeek = class['dayOfWeek'];
    var unit = class['unit'];
    var className = '[' + class['tag'] + ']' + class['class'];
    
    if (classList[semester] === undefined) classList[semester] = {};

    if (/[月火水木金土]/.exec(dayOfWeek) && semester.length == 2) {
      classList[semester][dayOfWeek][unit].push(className);
    } else {
      if (classList[semester][dayOfWeek] === undefined) classList[semester][dayOfWeek] = [];
      classList[semester][dayOfWeek].push(className);
    }
  });
  return classList;
}

function getClassListFrame () {
  var semesterList = ['前期', '後期']
  var dayOfWeekList = ['月', '火', '水', '木', '金', '土'];
  var unitList = ['１', '２', '３', '４', '５', '６', '７'];
  
  var classList = {};
  semesterList.forEach(function (semester) {
    classList[semester] = {};
    dayOfWeekList.forEach(function (dayOfWeek) {
      classList[semester][dayOfWeek] = {};
      unitList.forEach(function (unit) {
        classList[semester][dayOfWeek][unit] = [];
      });
    });
  });
  return classList;
}

function setClassList (semester, range, classList) {
  var dayOfWeek = ['月', '火', '水', '木', '金', '土'];
  var unit = ['１', '２', '３', '４', '５', '６', '７'];

  var rules = range.getDataValidations();
  var bgs = [];
  rules.forEach(function (row, i) {
    var bgRow = [];
    row.forEach(function (col, j) {
      var rule = classList[semester][dayOfWeek[j]][unit[i]];
      rules[i][j] = SpreadsheetApp.newDataValidation().requireValueInList(rule, true).build();
      if (rule.length === 0) {
        bgRow.push('#dddddd');
      } else {
        bgRow.push('#ffffff');
      }
    });
    bgs.push(bgRow);
  });
  range.setBackgrounds(bgs);
  range.setDataValidations(rules);
}

function updateTimeTable () {
  var spreadsheet = SpreadsheetApp.getActive();
  var timeTableSheet = spreadsheet.getSheetByName('TimeTable');
  var firstSemester = timeTableSheet.getRange(2, 2, 7, 6);
  var secondSemester = timeTableSheet.getRange(11, 2, 7, 6);
  
  timeTableSheet.getRange(1,1,timeTableSheet.getMaxRows(),timeTableSheet.getMaxColumns()).clearDataValidations();
  
  var classList = getClassList();
  classList = addTAList(classList);
  setClassList('前期', firstSemester, classList);
  setClassList('後期', secondSemester, classList);
}