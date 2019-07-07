interface Setup { task: string; date: string; period: string }
var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets()
let folder: GoogleAppsScript.Drive.Folder, setup: Setup
function getSettings(settings = sheets[1].getSheetValues(1, 2, 3, 1)): Setup {
  return { task: settings[0][0], date: settings[2][0], period: settings[1][0] }
}
function runPayroll() {
  setup = getSettings()
  if (sheets[0].getName() !== setup.period) copyInput()
  folder = getFolder(setup.period + ' PAYROLL')
  if (setup.task === "RESET") return folder.setTrashed(true)
  let actives = sheets[1].getSheetValues(2, 3, -1, 4).filter(s => s[3] === setup.task).map(s => s[0])
  saveStatementsToProps(runners[setup.task](actives))
}
let runners = {
  RUN: function runStatements(actives, period: string) {
    let template = DriveApp.getFileById("1Ultvt-EETMHGrJ9ttrCSlFyOLAccP7U0_J5aGspabu8")
    let charges = sheets[0].getSheetValues(2, 1, -1, -1)
      .reduce(function toItems(a: {}, i: any[]) {
        a[i[12]] = (a[i[12]] || []).concat([toRow(i)])
        return a
      }, {})
    let subjectData = getSubjectData(actives)
    return actives.map(function RUN(a) { return create(a) })
    function toRow(i: any[]) {
      return [i[1], i[3]].concat(i.slice(5, 9), i.slice(12, 15))
    }
    function insertData(sheet, subject) {
      var ranges = sheet.getNamedRanges()
      for (var i = 0; i < ranges.length; i++) {
        const name = ranges[i].getName()
        const val = setup[name] || subject[name]
        if (val) ranges[i].getRange().setValue(val)
      }
    }
    function insertItems(s, items) {
      s.insertRows(16, items.length)
      s.getRange(16, 1, items.length, items[0].length).setValues(items).setFontSize(10).setWrap(true)
      s.getRange(16, items[0].length - 1, items.length, 2).setNumberFormat('$0.00')
      SpreadsheetApp.flush()
    }
    function create(active) {
      const subject = subjectData[active]
      const statement = template.makeCopy(subject.id, folder)
      const sheet = SpreadsheetApp.open(statement).getSheets()[0]
      insertData(sheet, subject)
      insertItems(sheet, charges[subject.id])
      return statement
    }
  },
  PRINT: function printStatements(actives) {
    let savedProps = PropertiesService.getDocumentProperties().getProperties()
    function getStatement(subject) { return DriveApp.getFileById(JSON.parse(savedProps[subject]).id) }
    function print(subject) {
      let draft = getStatement(subject)
      draftsFolder.addFile(draft)
      draft.getParents().next().removeFile(draft)
      return savePdfCopy(draft.getName(), draft)
    }
    let draftsFolder = ((drafts = folder.getFoldersByName("DRAFTS")) => {
      return drafts.hasNext() ? drafts.next() : folder.createFolder("DRAFTS")
    })()
    function savePdfCopy(name, draft) {
      let url = draft.getUrl().replace('edit?usp=drivesdk', '')
      let headers = { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }
      let ext = 'export?exportFormat=pdf&format=pdf&size=letter&portrait=false'
        + '&fitw=true&gridlines=false&gid=0'
      let blob = UrlFetchApp.fetch(url + ext, { headers }).getBlob().setName(name)
      return folder.createFile(blob)
    }
    return actives.map(function PRINT(a) { return print(a) })
  }
}
function saveStatementsToProps(statements) {
  PropertiesService.getDocumentProperties().setProperties(statements.reduce(
    function getProps(a, s) {
      a[s.getName()] = JSON.stringify({
        id: s.getId(), state: setup.task, period: setup.period, date: setup.date
      })
      return a
    }, {})
  )
}
function getFolder(name) {
  let directory = DriveApp.getFolderById("1-MkOt90C9CKciGUK2zVLSCSxzuKxSVFh")
  let find = directory.getFoldersByName(name)
  return find.hasNext() ? find.next() : directory.createFolder(name)
}
function copyInput() {
  let masterId = "1pgM6_t_v5LsfLyFOLZTK6klW766m27-jbp1lhgbkfgk"
  let source = SpreadsheetApp.openById(masterId).getSheetByName(setup.period).getDataRange()
  sheets[0].clear().getRange(source.getA1Notation()).setValues(source.getValues())
  SpreadsheetApp.flush()
}
function getSubjectData(actives) {
  return sheets[2].getDataRange().getValues().reduce(function getDataObject(obj, row, _index, data) {
    if (actives.indexOf(row[0] > -1)) obj[row[0]] = data[0].reduce(
      function getPropsObject(obj, prop, index, row) {
        obj[prop] = row[index]
        return obj
      }, {})
    return obj
  }, {})
}