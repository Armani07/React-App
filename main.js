const electron = require('electron');
const app = electron.app;
const BrowserWindow = electron.BrowserWindow;
const dialog = electron.dialog;
var { ipcMain } = electron;
var Excel = require('exceljs');
var fs = require('fs');
var war = [];
var setting;
var default_setting = { screenSizeSetting: 'Full Screen' };
var mainWindow;
var size;
var trelloWindow;
var reactEvent;
var number;
var worksheet;
var tableName1;
var PotentialCII1;
var ActualCII1;
var ColumnsProcessed1;
var percentPotentialCII1;
var percentActualCII1;
var percentCombinedCII1;
var holdPercentActual = 0;
var holdPercentPotential = 0;
var holdPercentCombined = 0;
var PotentialCIIMath = 0;
var ActualCIIMath = 0;
var ColumnsProcessedMath = 0;
var percentPotentialCIIMath = 0;
var percentActualCIIMath = 0;
var percentCombinedCIIMath = 0;
var metric;
const logger = require('./logger.js');
const winston = require('winston');
const { clipboard } = require('electron');
var moment = require('moment');
require('electron-reload')(__dirname);
app.on('ready', () => {
  read = fs.readFile(app.getPath('userData') + "/userSetting.json", 'utf-8', (err, data) => {

    if (err) {

      mainWindow = new BrowserWindow({
        height: 800,
        width: 1400,
        x: 0,
        y: 0,
      })

      mainWindow.maximize()

      let url = require('url').format({
        protocol: 'file',
        slashes: true,
        pathname: require('path').join(__dirname, 'index.html')
      })
      // openExcelDialog();
      mainWindow.loadURL(url)
    }

    else {
      setting = JSON.parse(data);
      size = JSON.parse(data).screenSizeSetting;
      mainWindow = new BrowserWindow({
        height: (JSON.parse(data).screenSizeSetting === 'hsv') ? 730 : (JSON.parse(data).screenSizeSetting === 'hsh') ? 384 : 500,
        width: (JSON.parse(data).screenSizeSetting === 'hsv') ? 684 : (JSON.parse(data).screenSizeSetting === 'hsh') ? 1366 : 500,
        x: 0,
        y: 0,
      })
      if (JSON.parse(data).screenSizeSetting === 'Full Screen') { mainWindow.maximize() }

      let url = require('url').format({
        protocol: 'file',
        slashes: true,
        pathname: require('path').join(__dirname, 'index.html')
      })
      // openExcelDialog();
      mainWindow.loadURL(url)
    }
  })
  // mainWindow.webContents.openDevTools()
})
app.on('window-all-closed', function () {
  if (process.platform != 'darwin')
    app.quit();
});
ipcMain.on('initialSettings', function (event, arg) {
  reactEvent = event;
  if (setting) {
    for (var key in arg) {
      if (!setting.hasOwnProperty(key)) {
        setting[key] = arg[key]
      }
    }
  }
  else {
    setting = arg
  }
  fs.writeFile(app.getPath('userData') + "/userSetting.json", JSON.stringify(setting), function (err) {
    if (err) {

      event.sender.send('Error', "Error writing UserData file");
      alert("An error ocurred updating the file" + err.message);
      return;
    }
  })
  event.sender.send('writingUserSettings', setting);
})
ipcMain.on('gotToken', function (event, arg) {
  reactEvent.sender.send('trelloSend', arg);
  trelloWindow.close();
});
ipcMain.on('exportPhrases', (event, saveDir, arg1) => {
  exportPhrasesSave(function (callback) { }, event, saveDir, arg1)
})
ipcMain.on('phrasesMerged', (event, openDir, arg1, arg2) => {
  phrasesMergedChooseFile(function (callback) { }, event, arg1, arg2, openDir)
});
ipcMain.on('copyToClipboard', (event, script, snackBarMessage) => {
  clipboard.writeText(script);
  reactEvent.sender.send('copyMessage', snackBarMessage);
});
ipcMain.on('saveSettings', function (event, arg) {
  if (arg.screenSizeSetting === 'hsh') { mainWindow.setSize(1366, 435), mainWindow.setPosition(0, 0) }
  else if (arg.screenSizeSetting === 'hsv') { mainWindow.setSize(684, 730), mainWindow.setPosition(0, 0) }
  else { mainWindow.maximize() }

  fs.writeFile(app.getPath('userData') + "/userSetting.json", JSON.stringify(arg), function (err) {
    if (err) {
      alert("An error ocurred updating the file" + err.message);
      event.sender.send('Error', " Error writing userSetting");
      return;
    }
  })
})
// Listen for sync message from renderer process
ipcMain.on('sync', (event, arg, arg2, saveDir, save, overwrite) => {
  saveExcelDialog(function (callback) { }, event, arg, arg2, saveDir, save, overwrite);
});
ipcMain.on('saveOracle', (event, arg, schemaName, saveDir, save, overwrite) => {
  saveOracleDialog(function (callback) { }, event, arg, schemaName, saveDir, save, overwrite);
});
ipcMain.on('saveValidation', (event, arg, workingSet, saveDir, save, acronym, overwrite) => {
  saveValidationDialog(function (callback) { }, event, arg, workingSet, saveDir, save, acronym, overwrite);
});
//merge save
ipcMain.on('saveMetrics', (event, arg, saveDir) => {
  merge(function (callback) {
  }, event, arg, saveDir);
});
ipcMain.on('setSave', (event) => {
  setSave(function (callback) { }, event);
});
ipcMain.on('fetchFile', (event, file) => {
  openExcelDialog(function (callback) { }, event, file);
});
ipcMain.on('oracleOpen', (event, file) => {
  openOracleDialog(function (callback) { }, event, file);
});
ipcMain.on('qaOpen', (event, file) => {
  openQADialog(function (callback) { }, event, file);
});
ipcMain.on('trelloLoginToken', (event, file) => {
});
ipcMain.on('metricsFiles', (event, arg) => {
  mergeDialog(function (callback) { }, event, arg);
});
ipcMain.on('openSetting', (event, arg) => {
  openSetting(function (callback) { }, event);
});
ipcMain.on('fileUpdate', (event, arg1) => {
  fs.writeFile(app.getPath('userData') + "/userSetting.json", JSON.stringify(arg1), function (err) {
    if (err) {
      event.sender.send('Error', "fileUpdate Error");
      alert("An error ocurred updating the file" + err.message);
      return;
    }
  })
});

ipcMain.on('trelloLogin', (event) => {
  open(function (callback) { }, event);
});

function open() {
  trelloWindow = new BrowserWindow({
    height: 1000,
    width: 1400,
    x: 0,
    y: 0,
    webPreferences: { nodeIntegration: false, preload: __dirname + '/js/preload.js' }
  })

  let url = require('url').format({
    protocol: 'file',
    slashes: true,
    pathname: require('path').join(__dirname, 'trello.html')
  })
  trelloWindow.loadURL(url)
}

function mergeDialog(callback, event, file) {
  dialog.showOpenDialog({
    properties: ['openFile', 'multiSelections'],
    defaultPath: (('' === file) || (null === file)) ? app.getPath('documents') : file.toString(),
    filters: [
      { name: 'Excel Files', extensions: ['xlsx', 'xls'] }
    ]
  },

    function (fileNames) {
      if (fileNames === undefined) return event.sender.send('saved');
      mergeFilesOpen(event, fileNames)
    });
}

function merge(callback, event, arg, saveDir) {
  dialog.showSaveDialog({
    defaultPath: (('' === saveDir) || (null === saveDir)) ? app.getPath('documents') : saveDir.toString(),
    filters: [
      { name: 'Excel 2007 XLSX', extensions: ['xlsx'] }
    ]
  },

    function (fileName) {

      if (fileName === undefined) return event.sender.send('saved');
      var bucketSize = 5;
      var indexPos = 0;
      var bucketArray = [[]];
      arg.forEach(function (listItem, index) {
        if (bucketArray[indexPos].length === bucketSize) {
          bucketArray.push([])
          indexPos++;
        }
        bucketArray[indexPos].push(listItem);
      })
      indexPos = 0;
      bucketData = [];
      saveMerge(bucketArray, indexPos, bucketData).then(function (data) {
        saveMergeToFile(event, data, fileName)
      });
    });
}

function phrasesMergedChooseFile(callback, event, commonPhrasesTheArray, commonPhrases, openDir) {
  dialog.showOpenDialog({
    defaultPath: (('' === openDir) || (null === openDir)) ? app.getPath('documents') : openDir.toString(),
    filters: [
      { name: 'Excel 2007 XLSX', extensions: ['xlsx'] }
    ]
  }, function (fileName) {
    if (fileName === undefined) return event.sender.send('saved');
    phrasesImport(fileName, event, commonPhrasesTheArray, commonPhrases)
  }
  )
}

function exportPhrasesSave(callback, event, saveDir, commonPhrasesTheArray) {
  dialog.showSaveDialog({
    defaultPath: (('' === saveDir) || (null === saveDir)) ? app.getPath('documents') : saveDir.toString(),
    filters: [
      { name: 'Excel 2007 XLSX', extensions: ['xlsx'] }
    ]
  },

    function (fileName) {
      if (fileName === undefined) return event.sender.send('saved');
      exportPhrases(fileName, event, commonPhrasesTheArray);
    })
}

function exportPhrasesSave(callback, event, saveDir, commonPhrasesTheArray) {
  dialog.showSaveDialog({
    defaultPath: (('' === saveDir) || (null === saveDir)) ? app.getPath('documents') : saveDir.toString(),
    filters: [
      { name: 'Excel 2007 XLSX', extensions: ['xlsx'] }
    ]
  },

    function (fileName) {
      if (fileName === undefined) return event.sender.send('saved');
      exportPhrases(fileName, event, commonPhrasesTheArray);
    })
}

function saveMerge(arg, indexPos, data) {
  return new Promise(function (resolve, reject) {
    var arrayofMetrics = [];
    var object;
    arg[indexPos].forEach(function (listItem, index) {
      var workbook = new Excel.Workbook();
      data.push(workbook.xlsx.readFile(listItem.metricName).then(function () {
        worksheet = workbook.getWorksheet('Metrics');
        return {
          tableName: worksheet.getCell('A2').value,
          PotentialCII1: worksheet.getCell('B2').value.result,
          ActualCII1: worksheet.getCell('C2').value.result,
          ColumnsProcessed1: worksheet.getCell('D2').value.result,
          percentPotentialCII1: worksheet.getCell('E2').value.result,
          percentActualCII1: worksheet.getCell('F2').value.result,
          percentCombinedCII1: worksheet.getCell('G2').value.result
        }
      }));
    });
    Promise.all(data).then(array => {
      indexPos++;
      if (indexPos < arg.length) {
        saveMerge(arg, indexPos, array).then(data => {
          resolve(data)
        })
      }
      else {
        resolve(array)
      }
    });
  })
}

function saveMergeToFile(event, arrayofMetrics, fileName) {
  var workbook = new Excel.Workbook();//__dirname\\template.xlsx
  workbook.xlsx.readFile(__dirname + '\\MetricTemplate.xlsx')

    .then(function () {
      var metricSheet = workbook.getWorksheet('Metrics')
      var style1 = worksheet.getCell('A2').style;
      var style2 = worksheet.getCell('B2').style;
      var style3 = worksheet.getCell('B4').style;
      var style4 = worksheet.getCell('B5').style;
      for (var i = 2; (arrayofMetrics.length + 2) > i; i++) {
        metricSheet.getCell('A' + i.toString()).value = arrayofMetrics[i - 2].tableName;
        metricSheet.getCell('A' + i.toString()).style = style1;
        metricSheet.getCell('B' + i.toString()).value = arrayofMetrics[i - 2].PotentialCII1;
        metricSheet.getCell('B' + i.toString()).style = style2;
        metricSheet.getCell('C' + i.toString()).value = arrayofMetrics[i - 2].ActualCII1;
        metricSheet.getCell('C' + i.toString()).style = style2;
        metricSheet.getCell('D' + i.toString()).value = arrayofMetrics[i - 2].ColumnsProcessed1;
        metricSheet.getCell('D' + i.toString()).style = style2;
        metricSheet.getCell('E' + i.toString()).value = Math.round(arrayofMetrics[i - 2].percentPotentialCII1);
        metricSheet.getCell('E' + i.toString()).style = style2;
        metricSheet.getCell('F' + i.toString()).value = Math.round(arrayofMetrics[i - 2].percentActualCII1);
        metricSheet.getCell('F' + i.toString()).style = style2;
        metricSheet.getCell('G' + i.toString()).value = Math.round(arrayofMetrics[i - 2].percentCombinedCII1);
        metricSheet.getCell('G' + i.toString()).style = style2;
        metricSheet.getCell('H' + i.toString()).value = " ";
        metricSheet.getCell('H' + i.toString()).style = style2;

        if (i === (arrayofMetrics.length + 1)) {
          metricSheet.getCell('B' + (i + 1).toString()).value = "Total Potential"
          metricSheet.getCell('B' + (i + 1).toString()).style = style3;
          metricSheet.getCell('C' + (i + 1).toString()).value = "Total Actual"
          metricSheet.getCell('C' + (i + 1).toString()).style = style3;
          metricSheet.getCell('D' + (i + 1).toString()).value = "Total Columns Processed"
          metricSheet.getCell('D' + (i + 1).toString()).style = style3;
          metricSheet.getCell('E' + (i + 1).toString()).value = "Average Potential %"
          metricSheet.getCell('E' + (i + 1).toString()).style = style3;
          metricSheet.getCell('F' + (i + 1).toString()).value = "Average Actual %"
          metricSheet.getCell('F' + (i + 1).toString()).style = style3;
          metricSheet.getCell('G' + (i + 1).toString()).value = "Average Combined %"
          metricSheet.getCell('G' + (i + 1).toString()).style = style3;
          metricSheet.getCell('B' + (i + 2).toString()).value = { formula: "SUM(B2:B" + (1 + arrayofMetrics.length) + ")" }
          metricSheet.getCell('B' + (i + 2).toString()).style = style4;
          metricSheet.getCell('C' + (i + 2).toString()).value = { formula: "SUM(C2:C" + (1 + arrayofMetrics.length) + ")" }
          metricSheet.getCell('C' + (i + 2).toString()).style = style4;
          metricSheet.getCell('D' + (i + 2).toString()).value = { formula: "SUM(D2:D" + (1 + arrayofMetrics.length) + ")" }
          metricSheet.getCell('D' + (i + 2).toString()).style = style4;
          metricSheet.getCell('E' + (i + 2).toString()).value = { formula: "AVERAGE(E2:E" + (1 + arrayofMetrics.length) + ")" }
          metricSheet.getCell('E' + (i + 2).toString()).style = style4;
          metricSheet.getCell('F' + (i + 2).toString()).value = { formula: "AVERAGE(F2:F" + (1 + arrayofMetrics.length) + ")" }
          metricSheet.getCell('F' + (i + 2).toString()).style = style4;
          metricSheet.getCell('G' + (i + 2).toString()).value = { formula: "AVERAGE(G2:G" + (1 + arrayofMetrics.length) + ")" }
          metricSheet.getCell('G' + (i + 2).toString()).style = style4;
          workbook.xlsx.writeFile(fileName);
        }
      }
    })
  metric = [];
  holdPercentActual = 0
  holdPercentPotential = 0
  holdPercentCombined = 0
  PotentialCIIMath = 0
  ActualCIIMath = 0
  ColumnsProcessedMath = 0
  percentPotentialCIIMath = 0
  percentActualCIIMath = 0
  percentCombinedCIIMath = 0
  tableName1 = null
  PotentialCII1 = null
  ActualCII1 = null
  ColumnsProcessed1 = null
  percentPotentialCII1 = null
  percentActualCII1 = null
  percentCombinedCII1 = null
  event.sender.send('saved');
}
function openExcelDialog(callback, event, file) {

  dialog.showOpenDialog({
    properties: ['openFile'],
    defaultPath: (('' === file) || (null === file)) ? app.getPath('documents') : file.toString(),
    filters: [
      { name: 'Excel Files', extensions: ['xlsx', 'xls'] }
    ]
  }, function (fileNames) {
    if (fileNames === undefined) return event.sender.send('saved');
    var fileName = fileNames[0];
    makeSupraColumnList(fileName, event);
  });
}

function openOracleDialog(callback, event, file) {
  dialog.showOpenDialog({
    properties: ['openFile'],
    defaultPath: (('' === file) || (null === file)) ? app.getPath('documents') : file.toString(),
    filters: [
      { name: 'Excel Files', extensions: ['xlsx', 'xls'] }
    ]
  }, function (fileNames) {
    if (fileNames === undefined) return event.sender.send('saved');
    var fileName = fileNames[0];
    makeOracleList(fileName, event);
  });
}
function openQADialog(callback, event, file) {
  dialog.showOpenDialog({
    properties: ['openFile'],
    defaultPath: (('' === file) || (null === file)) ? app.getPath('documents') : file.toString(),
    filters: [
      { name: 'Excel Files', extensions: ['xlsx', 'xls'] }
    ]
  }, function (fileNames) {
    if (fileNames === undefined) return event.sender.send('saved');
    var fileName = fileNames[0];
    makeQAList(fileName, event);
  });
}
function openSetting(callback, event) {
  dialog.showOpenDialog({
    properties: ['openDirectory'],
    filters: [
      { name: 'Excel Files', extensions: ['xlsx', 'xls'] }
    ]
  }, function (fileName) {
    if (fileName === undefined) return
    event.sender.send('settingOpenReturn', fileName);
  });
}
function mergeFilesOpen(event, files) {
  var newArray = []
  for (i = 0; files.length > i; i++) {
    newArray.push({ metricName: files[i] })
  }
  event.sender.send('mergeOpenFilesReturn', newArray)
  event.sender.send('saved');
}

function saveExcelDialog(callback, event, stored, ok, saveDir, save, overwrite) {

  if (!save) {
    dialog.showSaveDialog({
      defaultPath: (('' === saveDir) || (null === saveDir)) ? app.getPath('documents') : saveDir.toString(),
      filters: [
        { name: 'Excel 2007 XLSX', extensions: ['xlsx'] }
      ]
    }, function (fileName) {
      if (fileName === undefined) {
        // I need a return message that says I didn't save due to the filename being empty or just throw back an error
        return event.sender.send('saved');
      }
      // Check if file exists, if not save. If it does throw back a message asking then writeExcel
      writeExcel(fileName, stored, ok, event);
    });
  } else if (overwrite) {
    writeExcel((('' === saveDir) || (null === saveDir)) ? app.getPath('documents') + '\\' + ok + '.xlsx' : saveDir.toString() + '\\' + ok + '.xlsx', stored, ok, event);
  }
  else {
    fs.readFile(app.getPath('userData') + (('' === saveDir) || (null === saveDir)) ? app.getPath('documents') + '\\' + ok + '.xlsx' : saveDir.toString() + '\\' + ok + '.xlsx', 'utf-8', (err, data) => {
      if (err) {
        writeExcel((('' === saveDir) || (null === saveDir)) ? app.getPath('documents') + '\\' + ok + '.xlsx' : saveDir.toString() + '\\' + ok + '.xlsx', stored, ok, event);
      }
      else {
        event.sender.send('existing file')
      }
    })
  }
}

function saveOracleDialog(callback, event, stored, schemaName, saveDir, save, overwrite) {
  if (!save) {
    dialog.showSaveDialog({
      defaultPath: (('' === saveDir) || (null === saveDir)) ? app.getPath('documents') : saveDir.toString(),
      filters: [
        { name: 'Excel 2007 XLSX', extensions: ['xlsx'] }
      ]
    }, function (fileName) {
      if (fileName === undefined) {
        // I need a return message that says I didn't save due to the filename being empty or just throw back an error
        return event.sender.send('saved');
      }
      // Check if file exists, if not save. If it does throw back a message asking then writeExcelOracle
      writeExcelOracle(fileName, stored, schemaName, event);
    });
  } else if (overwrite) {
    writeExcelOracle((('' === saveDir) || (null === saveDir)) ? app.getPath('documents') + '\\' + schemaName + '.xlsx' : saveDir.toString() + '\\' + schemaName + '.xlsx', stored, schemaName, event);
  }
  else {
    fs.readFile(app.getPath('userData') + (('' === saveDir) || (null === saveDir)) ? app.getPath('documents') + '\\' + schemaName + '.xlsx' : saveDir.toString() + '\\' + schemaName + '.xlsx', 'utf-8', (err, data) => {
      if (err) {
        writeExcelOracle((('' === saveDir) || (null === saveDir)) ? app.getPath('documents') + '\\' + schemaName + '.xlsx' : saveDir.toString() + '\\' + schemaName + '.xlsx', stored, schemaName, event);
      }
      else {
        event.sender.send('existing file')
      }
    })
  }
}

function saveValidationDialog(callback, event, stored, workingSet, saveDir, save, acronym, overwrite) {
  if (!save) {
    dialog.showSaveDialog({
      defaultPath: (('' === saveDir) || (null === saveDir)) ? app.getPath('documents') : saveDir.toString(),
      filters: [
        { name: 'Excel 2007 XLSX', extensions: ['xlsx'] }
      ]
    }, function (fileName) {
      if (fileName === undefined) {
        // I need a return message that says I didn't save due to the filename being empty or just throw back an error
        return event.sender.send('saved');
      }
      // Check if file exists, if not save. If it does throw back a message asking then writeExcelValidation
      writeExcelValidation(fileName, stored, workingSet, acronym, event);
    });
  } else if (overwrite) {
    writeExcelValidation((('' === saveDir) || (null === saveDir)) ? app.getPath('documents') + '\\' + workingSet + '.xlsx' : saveDir.toString() + '\\' + workingSet + '.xlsx', stored, workingSet, acronym, event);
  }
  else {
    fs.readFile(app.getPath('userData') + (('' === saveDir) || (null === saveDir)) ? app.getPath('documents') + '\\' + workingSet + '.xlsx' : saveDir.toString() + '\\' + workingSet + '.xlsx', 'utf-8', (err, data) => {
      if (err) {
        writeExcelValidation((('' === saveDir) || (null === saveDir)) ? app.getPath('documents') + '\\' + workingSet + '.xlsx' : saveDir.toString() + '\\' + workingSet + '.xlsx', stored, workingSet, acronym, event);
      }
      else {
        event.sender.send('existing file')
      }
    })
  }
}

function setSave(callback, moo) {
  dialog.showOpenDialog({
    properties: ['openDirectory']
  }, function (fileName) {
    if (fileName === undefined) return;
    moo.sender.send('ok', fileName);
  });
}

function readTemplate(fileName) {
  if (typeof require !== 'undefined') XLSX = require('xlsx');
  var workbook = XLSX.readFile('C:\\Users\\AB87148\\Documents\\test1.xlsx');
  XLSX.writeFile(workbook, 'C:\\Users\\AB87148\\Documents\\test5.xlsx', { bookType: 'xlsx', bookSST: true, type: 'binary' });
}

function readExcel(fileName) {
  if (typeof require !== 'undefined') XLSX = require('xlsx');
  var workbook = XLSX.readFile(fileName);
  /* DO SOMETHING WITH workbook HERE */
  var first_sheet_name = workbook.SheetNames[0];
  var address_of_cell = 'A1';
  /* Get worksheet */
  var worksheet = workbook.Sheets[first_sheet_name];
  /* Find desired cell */
  var desired_cell = worksheet[address_of_cell];
  /* Get the value */
  var desired_value = (desired_cell ? desired_cell.v : undefined);
  return desired_value;
}

function exportPhrases(fileName, event, commonPhrasesTheArray) {
  var workbook = new Excel.Workbook();
  workbook.xlsx.readFile(__dirname + '\\phrasesTemplate.xlsx')
    .then(function () {
      var worksheet = workbook.getWorksheet('Phrases');
      for (var i = 0; commonPhrasesTheArray.length > i; i++) {
        worksheet.getCell('A' + (i + 2).toString()).value = commonPhrasesTheArray[i];
      }
      workbook.xlsx.writeFile(fileName).then(function () { event.sender.send('saved', "Table Saved", ); }, function () { event.sender.send('Error', "Export Phrases Error", ); })
    })
}

function makeSupraColumnList(file, event) {
  var workbook = new Excel.Workbook();
  var superDuperArray = [];

  workbook.xlsx.readFile(file)
    .then(function () {
      var worksheet = workbook.getWorksheet(1);
      var i = 0;
      column = {}
      n = false;
      idValue = 0;
      while (n === false) {
        if (worksheet.getCell('A' + (i + 6)).value === null) {
          n = true;
          break
        }
        idValue = i
        var id = i;
        var questions = worksheet.getCell('G' + (i + 6)).value.richText[2].text;
        var comments = worksheet.getCell('G' + (i + 6)).value.richText[0].text;
        var dataSizeType = worksheet.getCell('B' + (i + 6)).value;
        var dataSize = (dataSizeType) ? ((worksheet.getCell('B' + (i + 6)).value[worksheet.getCell('B' + (i + 6)).value.length - 1] === ')') ? worksheet.getCell('B' + (i + 6)).value.split('(')[1].split(')')[0] : "") : '';
        var dataType = (dataSizeType) ? (worksheet.getCell('B' + (i + 6)).value.split('(')[0]) : '';
        var columnName = worksheet.getCell('A' + (i + 6)).value;
        var nullable = (worksheet.getCell('C' + (i + 6)).value == 'Yes') ? true : false;
        var definition = (worksheet.getCell('D' + (i + 6)).value === '?') ? null : worksheet.getCell('D' + (i + 6)).value;
        var metric = (worksheet.getCell('F' + (i + 6)).value && worksheet.getCell('E' + (i + 6)).value === 'X') ? (worksheet.getCell('F' + (i + 6)).value.toUpperCase()) : null;
        var cii = (worksheet.getCell('E' + (i + 6)).value === 'X') ? true : false;
        column = { id, questions, comments, dataSize, dataType, columnName, nullable, definition, metric, cii, completed: false }
        superDuperArray[i] = column
        i++
      }

      event.sender.send('synchronous-reply', superDuperArray, worksheet.getCell('G1').value, idValue);
    }, function () { event.sender.send('Error', "Error loading in Excel Sheet", ); })
}

function makeOracleList(file, event) {
  var workbook = new Excel.Workbook();
  var oracleArray = [];

  workbook.xlsx.readFile(file)
    .then(function () {
      var worksheet = workbook.getWorksheet(1);
      var i = 0;

      column = {}
      n = false;
      idValue = 0;
      while (n === false) {

        if (worksheet.getCell('A' + (i + 4)).value === null) {
          n = true;
          break
        }
        idValue = i
        var id = i
        var tableName = worksheet.getCell('A' + (i + 4)).value;
        var columnName = worksheet.getCell('B' + (i + 4)).value;
        var dataType = worksheet.getCell('C' + (i + 4)).value;
        var objectType = worksheet.getCell('D' + (i + 4)).value;
        var status = worksheet.getCell('E' + (i + 4)).value;
        var compliant = worksheet.getCell('F' + (i + 4)).value;
        var loadDate = worksheet.getCell('G' + (i + 4)).value;
        if (loadDate) {
          loadDate = moment(loadDate).add(1, 'days').format('DD-MMM-YY')
        }
        column = { id, tableName, columnName, dataType, objectType, status, compliant, loadDate }
        oracleArray[i] = column
        i++
      }

      event.sender.send('OracleOpenReply', oracleArray, worksheet.getCell('B1').value, idValue);
    }, function () { event.sender.send('Error', "Error loading in Excel Sheet", ); })
}
function makeQAList(file, event) {
  var workbook = new Excel.Workbook();
  var qaArray = [];

  workbook.xlsx.readFile(file)
    .then(function () {
      var worksheet = workbook.getWorksheet(1);
      var i = 0;

      column = {}
      n = false;
      idValue = 0;
      while (n === false) {

        if (worksheet.getCell('A' + (i + 2)).value === null) {
          n = true;
          break
        }
        idValue = i
        var id = i
        var appName = worksheet.getCell('A' + (i + 2)).value;
        var acronym = worksheet.getCell('B' + (i + 2)).value;
        var q1 = (worksheet.getCell('C' + (i + 2)).value == 'YES') ? true : false;
        var q2Answer = worksheet.getCell('D' + (i + 2)).value 
        var q3 = (worksheet.getCell('E' + (i + 2)).value == 'YES') ? true : false;
        var q4 = (worksheet.getCell('F' + (i + 2)).value == 'YES') ? true : false;
        var q5 = (worksheet.getCell('G' + (i + 2)).value == 'YES') ? true : false;
        var q6 = (worksheet.getCell('H' + (i + 2)).value == 'YES') ? true : false;
        var q7Answer = worksheet.getCell('I' + (i + 2)).value 
        var q8 = (worksheet.getCell('J' + (i + 2)).value == 'YES') ? true : false;
        // var question = worksheet.getCell('K' + (i + 2)).value;
        // var comments = worksheet.getCell('L' + (i + 2)).value;
        // var loadDate = worksheet.getCell('M' + (i + 2)).value;
        column = { id, appName,acronym, q1, q2Answer, q3, q4, q5, q6, q7Answer, q8, completed: false }
        qaArray[i] = column
        i++

      }

      event.sender.send('QAOpenReply', qaArray, worksheet.getCell('K2').value, idValue);
    }, function () { event.sender.send('Error', "Error loading in Excel Sheet", ); })
}
function phrasesImport(file, event, commonPhrasesTheArray, commonPhrases, ) {
  logger.log(commonPhrasesTheArray + JSON.stringify(commonPhrases))
  var workbook = new Excel.Workbook();

  workbook.xlsx.readFile(file[0])
    .then(function () {
      var importPhraseArray = commonPhrasesTheArray;
      var importPhrase = commonPhrases
      var worksheet = workbook.getWorksheet('Phrases');
      var i = 2;
      column = {}
      n = false;
      idValue = 0;
      while (n === false) {
        if (worksheet.getCell('A' + i).value === null) {
          n = true;
          break
        }
        importPhrase.push({ commonPhrase: worksheet.getCell('A' + i).value })
        importPhraseArray.push(worksheet.getCell('A' + i).value)
        i++
      }
      logger.log(JSON.stringify(importPhrase) + importPhraseArray)
      event.sender.send('phrasesImportReply', importPhraseArray, importPhrase);
    })
}
function writeExcel(filePath, columns, tabNamePlusUltraLive, event) {
  columns = columns.sort((a, b) => {
    var aSort = a.columnName.toLowerCase(), bSort = b.columnName.toLowerCase();
    if (aSort < bSort) {
      return -1;
    } else if (aSort > bSort) {
      return 1;
    }
    return 0;
  })
  var workbook = new Excel.Workbook();//__dirname\\template.xlsx
  workbook.xlsx.readFile(__dirname + '\\Template.xlsx')
    .then(function () {
      try {
        var worksheet = workbook.getWorksheet('Sheet1');
        var metricSheet = workbook.getWorksheet('Metrics')
        worksheet.name = tabNamePlusUltraLive
        metricSheet.getCell('A2').value = tabNamePlusUltraLive
        for (i = 6; i <= 25; i++) {
          metricSheet.getCell('A' + i).fill = { type: 'pattern', pattern: 'solid', fgColor: { theme: 1, tint: 0.499984740745262 }, bgColor: { indexed: 64 } }
        }
        var stylelo = worksheet.getCell('A4').style;
        stylelo.border.right.style = 'thick'
        worksheet.getCell('A4').style = stylelo;
        worksheet.getCell('H4').style = stylelo;
        var style1 = worksheet.getCell('B6').style;
        style1.font = { bold: false, size: 8, name: 'Arial', family: 2 };
        style1.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
        var style2 = worksheet.getCell('A6').style;
        style2.border.left = { style: 'thick' };
        style2.font = { bold: false, size: 8, name: 'Arial', family: 2 };
        style2.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
        var style3 = worksheet.getCell('F6').style;
        style3.font = { bold: false, size: 8, name: 'Arial', family: 2 };
        style3.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
        var style4 = worksheet.getCell('G6').style;
        style4.font = { bold: false, size: 8, name: 'Arial', family: 2 };
        style4.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
        var style5 = worksheet.getCell('M6').style;
        style5.font = { bold: false, size: 8, name: 'Arial', family: 2 };
        style5.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
        var bot = worksheet.getCell('B7').style;
        bot.font = { bold: false, size: 8, name: 'Arial', family: 2 };
        bot.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
        var bot2 = worksheet.getCell('A7').style;
        bot2.border.left = { style: 'thick' };
        bot2.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
        bot2.font = { bold: false, size: 8, name: 'Arial', family: 2 };
        var bot3 = worksheet.getCell('F7').style;
        bot3.font = { bold: false, size: 8, name: 'Arial', family: 2 };
        bot3.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
        var bot4 = worksheet.getCell('G7').style;
        bot4.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
        bot4.font = { bold: false, size: 8, name: 'Arial', family: 2 };
        var bot5 = worksheet.getCell('M7').style;
        bot5.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
        bot5.font = { bold: false, size: 8, name: 'Arial', family: 2 };
        bot5.border.top = { style: 'thin' };
        worksheet.properties.tabColor = { argb: '00FF00' }

        for (num = 6; num - 6 !== columns.length; num++) {
          var isNull = 'No';
          var isCust = '';
          var isUndefined = '?'

          if (columns[num - 6].dataSize === null || !columns[num - 6].dataSize) {
            var newDataType = columns[num - 6].dataType
          } else {
            var newDataType = columns[num - 6].dataType + '(' + columns[num - 6].dataSize + ')'
          }
          if (columns[num - 6].cii) {
            isCust = 'X'
          }
          if (columns[num - 6].nullable) {
            isNull = 'Yes';
          }
          if (columns[num - 6].definition) {
            isUndefined = columns[num - 6].definition;
          }
          if (columns[num - 6].questions) {
            worksheet.properties.tabColor = { argb: 'FFFF0000' };
          }
          if (num - 5 === columns.length) {
            worksheet.getCell('A' + num.toString()).value = columns[num - 6].columnName;
            worksheet.getCell('A' + num.toString()).style = bot2;
            worksheet.getCell('B' + num.toString()).value = newDataType;
            worksheet.getCell('B' + num.toString()).style = bot;
            worksheet.getCell('C' + num.toString()).value = isNull;
            worksheet.getCell('C' + num.toString()).style = bot;
            worksheet.getCell('D' + num.toString()).value = isUndefined;
            worksheet.getCell('D' + num.toString()).style = bot;
            worksheet.getCell('E' + num.toString()).value = isCust;
            worksheet.getCell('E' + num.toString()).style = bot;
            worksheet.getCell('F' + num.toString()).value = columns[num - 6].metric;
            worksheet.getCell('F' + num.toString()).style = bot3;
            worksheet.getCell('G' + num.toString()).value = {
              'richText': [
                { 'font': { 'color': { 'argb': '00000000' } }, 'text': (columns[num - 6].comments !== null) ? columns[num - 6].comments : " " },
                { 'font': { 'size': 12, 'color': { 'theme': 1 }, 'name': 'Calibri', 'family': 2, 'scheme': 'minor' }, 'text': '\n ' },
                { 'font': { 'color': { 'argb': 'FFFF0000' }, }, 'text': (columns[num - 6].questions !== null) ? columns[num - 6].questions : " " }
              ]
            };
            worksheet.getCell('G' + num.toString()).style = bot4;
            worksheet.getCell('H' + num.toString()).style = bot2;
            worksheet.getCell('I' + num.toString()).style = bot;
            worksheet.getCell('J' + num.toString()).style = bot;
            worksheet.getCell('K' + num.toString()).style = bot;
            worksheet.getCell('L' + num.toString()).style = bot;
            worksheet.getCell('M' + num.toString()).style = bot5;
            break;
          }
          else {
            worksheet.getCell('A' + num.toString()).value = columns[num - 6].columnName;
            worksheet.getCell('A' + num.toString()).style = style2;
            worksheet.getCell('B' + num.toString()).value = newDataType;
            worksheet.getCell('B' + num.toString()).style = style1;
            worksheet.getCell('C' + num.toString()).value = isNull;
            worksheet.getCell('C' + num.toString()).style = style1;
            worksheet.getCell('D' + num.toString()).value = isUndefined;
            worksheet.getCell('D' + num.toString()).style = style1;
            worksheet.getCell('E' + num.toString()).value = isCust;
            worksheet.getCell('E' + num.toString()).style = style1;
            worksheet.getCell('F' + num.toString()).value = columns[num - 6].metric;
            worksheet.getCell('F' + num.toString()).style = style3;
            worksheet.getCell('G' + num.toString()).value = {
              'richText': [
                { 'font': { 'color': { 'argb': '00000000' }, }, 'text': (columns[num - 6].comments !== null) ? columns[num - 6].comments : " " },
                { 'font': { 'size': 12, 'color': { 'theme': 1 }, 'name': 'Calibri', 'family': 2, 'scheme': 'minor' }, 'text': '\n ' },
                { 'font': { 'color': { 'argb': 'FFFF0000' }, }, 'text': (columns[num - 6].questions !== null) ? columns[num - 6].questions : " " }
              ]
            };
            worksheet.getCell('G' + num.toString()).style = style4;
            worksheet.getCell('H' + num.toString()).style = style2;
            worksheet.getCell('I' + num.toString()).style = style1;
            worksheet.getCell('J' + num.toString()).style = style1;
            worksheet.getCell('K' + num.toString()).style = style1;
            worksheet.getCell('L' + num.toString()).style = style1;
            worksheet.getCell('M' + num.toString()).style = style5;
          }
        }
        worksheet.getCell('G1').value = tabNamePlusUltraLive;
        metricSheet.getCell('B2').value = { formula: metricSheet.getCell('B2').formula }
        metricSheet.getCell('C2').value = { formula: metricSheet.getCell('C2').formula }
        metricSheet.getCell('D2').value = { formula: metricSheet.getCell('D2').formula }
        metricSheet.getCell('E2').value = { formula: metricSheet.getCell('E2').formula }
        metricSheet.getCell('F2').value = { formula: metricSheet.getCell('F2').formula }
        metricSheet.getCell('G2').value = { formula: metricSheet.getCell('G2').formula }
        metricSheet.getCell('B5').value = { formula: metricSheet.getCell('B5').formula }
        metricSheet.getCell('C5').value = { formula: metricSheet.getCell('C5').formula }
        metricSheet.getCell('D5').value = { formula: metricSheet.getCell('D5').formula }
        metricSheet.getCell('E5').value = { formula: metricSheet.getCell('E5').formula }
        metricSheet.getCell('F5').value = { formula: metricSheet.getCell('F5').formula }
        metricSheet.getCell('G5').value = { formula: metricSheet.getCell('G5').formula }
        workbook.xlsx.writeFile(filePath).then(
          function () {
            event.sender.send('saved', "Table Saved");
          },
          function () {
            event.sender.send('Error', "Saving Error", "Error saving Visual Oracle Data Scrubber report")
          }
        );
      } catch (err) {
        event.sender.send('Error', "Saving Error");
      }
    });
}
function writeExcelOracle(filePath, columns, schemaName, event) {
  columns = columns.sort((a, b) => {
    var aSort = a.columnName.toLowerCase(), bSort = b.columnName.toLowerCase();
    if (aSort < bSort) {
      return -1;
    } else if (aSort > bSort) {
      return 1;
    }
    return 0;
  })
  var workbook = new Excel.Workbook();//__dirname\\oracletemplate.xlsx
  workbook.xlsx.readFile(__dirname + '\\OracleTemplate.xlsx')
    .then(function () {
      try {
        var worksheet = workbook.getWorksheet(1);
        worksheet.name = schemaName
        var bot = worksheet.getCell('A2').style;
        worksheet.getCell('B1').value = schemaName;
        bot.font = { bold: false, size: 8, name: 'Arial', family: 2 };
        bot.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true }
        worksheet.getCell('B1').style = bot;
        for (num = 4; num - 4 !== columns.length; num++) {
          worksheet.getCell('A' + num.toString()).value = columns[num - 4].tableName;
          worksheet.getCell('A' + num.toString()).style = bot;
          worksheet.getCell('B' + num.toString()).value = columns[num - 4].columnName;
          worksheet.getCell('B' + num.toString()).style = bot;
          worksheet.getCell('C' + num.toString()).value = columns[num - 4].dataType;
          worksheet.getCell('C' + num.toString()).style = bot;
          worksheet.getCell('D' + num.toString()).value = columns[num - 4].objectType;
          worksheet.getCell('D' + num.toString()).style = bot;
          worksheet.getCell('E' + num.toString()).value = columns[num - 4].status;
          worksheet.getCell('E' + num.toString()).style = bot;
          worksheet.getCell('F' + num.toString()).value = columns[num - 4].compliant;
          worksheet.getCell('F' + num.toString()).style = bot;
          worksheet.getCell('G' + num.toString()).value = columns[num - 4].loadDate;
          worksheet.getCell('G' + num.toString()).style = bot;
        }
        workbook.xlsx.writeFile(filePath).then(function () { event.sender.send('saved', "Table Saved", ); }, function () {
          event.sender.send('Error', "Saving Error", "Error saving Oracle Automation report");
        });
      } catch (err) {
        event.sender.send('Error', "Saving Error");
      }
    });
}


function writeExcelValidation(filePath, columns, workingSet, acronym, event) {
  // columns = columns.sort((a, b) => {
  //   var aSort = a.columnName.toLowerCase(), bSort = b.columnName.toLowerCase();
  //   if (aSort < bSort) {
  //     return -1;
  //   } else if (aSort > bSort) {
  //     return 1;
  //   }
  //   return 0;
  // })
  var workbook = new Excel.Workbook();//__dirname\\QATemplate.xlsx
  workbook.xlsx.readFile(__dirname + '\\AppValidationTemplate.xlsx')

    .then(function () {
      try {
        var worksheet = workbook.getWorksheet(1);
        worksheet.name = 'Application Validation'
        var bot = worksheet.getCell('A2').style;
        //worksheet.getCell('B1').value = workingSet;
        bot.font = { bold: false, size: 11, name: 'Calibri', family: 2 };
        bot.alignment = { horizontal: 'left', vertical: 'middle', wrapText: true }
        worksheet.getCell('A2').style = bot;
        for (num = 2; num - 2 !== columns.length; num++) {
          worksheet.getCell('A' + num.toString()).value = columns[num - 2].appName;
          worksheet.getCell('A' + num.toString()).style = bot;
          worksheet.getCell('B' + num.toString()).value = columns[num - 2].acronym;
          worksheet.getCell('B' + num.toString()).style = bot;
          worksheet.getCell('C' + num.toString()).value = columns[num - 2].q1 ? 'YES' : 'NO'
          worksheet.getCell('C' + num.toString()).style = bot;
          worksheet.getCell('D' + num.toString()).value = columns[num - 2].q2Answer
          worksheet.getCell('D' + num.toString()).style = bot;
          worksheet.getCell('E' + num.toString()).value = columns[num - 2].q3 ? 'YES' : 'NO'
          worksheet.getCell('E' + num.toString()).style = bot;
          worksheet.getCell('F' + num.toString()).value = columns[num - 2].q4 ? 'YES' : 'NO'
          worksheet.getCell('F' + num.toString()).style = bot;
          worksheet.getCell('G' + num.toString()).value = columns[num - 2].q5 ? 'YES' : 'NO'
          worksheet.getCell('G' + num.toString()).style = bot;
          worksheet.getCell('H' + num.toString()).value = columns[num - 2].q6 ? 'YES' : 'NO'
          worksheet.getCell('H' + num.toString()).style = bot;
          worksheet.getCell('I' + num.toString()).value = columns[num - 2].q7Answer
          worksheet.getCell('I' + num.toString()).style = bot;
          worksheet.getCell('J' + num.toString()).value = columns[num - 2].q8 ? 'YES' : 'NO'
          worksheet.getCell('J' + num.toString()).style = bot;
          worksheet.getCell('K' + num.toString()).value = workingSet //This is the user name
          worksheet.getCell('K' + num.toString()).style = bot;
          // worksheet.getCell('L' + num.toString()).value = columns[num - 2].question;
          // worksheet.getCell('L' + num.toString()).style = bot;
          // worksheet.getCell('M' + num.toString()).value = columns[num - 2].comments;
          // worksheet.getCell('M' + num.toString()).style = bot;
        }
        workbook.xlsx.writeFile(filePath).then(
          function () {
            event.sender.send('saved', "Table Saved", );
          },
          function () {
            event.sender.send('Error', "Saving Error", "Error saving Validation report");
          });
      } catch (err) {
        event.sender.send('Error', "Saving Error");
      }
    });
}
function writeExcelBlah(outputFile, columns) {
  var workbook = writeSome();
  workbook.xlsx.writeFile(outputFile)
    .then(function () {
    });
}