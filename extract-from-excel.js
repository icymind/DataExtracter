const XLSX = require("xlsx")
const glob = require("glob")
const assert = require("assert")
const os = require("os")
const path = require("path")
const papa = require("papaparse")
const iconv = require("iconv-lite")

const structure = {
  "Audit Date-label": "B1",
  "Audit Date-value": "D1",
  "Auditor-label": "B2",
  "Auditor-value": "D2",
  "Audit Type-label": "B3",
  "Audit Type-value": "D3",
  "Production Status-label": "B4",
  "Production Status-value": "D4",
  "PO Number-label": "B5",
  "PO Number-value": "D5",
  "Department-label": "F1",
  "Department-value": "G1",
  "Product Name-label": "F2",
  "Product Name-value": "G2",
  "Vendor-label": "F3",
  "Vendor-value": "G3",
  "Factory-label": "F4",
  "Factory-value": "G4",
  "REI Style Number-label": "J1",
  "REI Style Number-value": "L1",
  "Season-label": "J2",
  "Season-value": "L2",
  "GA Product Number-label": "J3",
  "GA Product Number-value": "L3",
  "Product Spec-label": "J4",
  "Product Spec-value": "L4",
  "Product Lifecycle-label": "J5",
  "Product Lifecycle-value": "L5",
  "Audit Level-label": "M1",
  "Audit Level-value": "O1",
  "Audit Quality Level-label": "M2",
  "Audit Quality Level-value": "O2",
  "Audit Lot Size-label": "M3",
  "Audit Lot Size-value": "O3",
  "Audit Sample Quantity-label": "M4",
  "Audit Sample Quantity-value": "O4",
  "Audit Accept Amount-label": "M5",
  "Audit Accept Amount-value": "O5",
  "Non-conformity Type-label": "B8",
  "Total Minor Qty-label": "D8",
  "Total Major Qty-label": "F8",
  "Total Critical Qty-label": "H8",
  "Critical-label": "B9",
  "Critical Minor-value": "D9",
  "Critical Major-value": "F9",
  "Critical Critical-value": "H9",
  "Fabric-label": "B10",
  "Fabric Minor-value": "D10",
  "Fabric Major-value": "F10",
  "Fabric Critical-value": "H10",
  "Hangtag-label": "B11",
  "Hangtag Minor-value": "D11",
  "Hangtag Major-value": "F11",
  "Hangtag Critical-value": "H11",
  "Label-label": "B12",
  "Label Minor-value": "D12",
  "Label Major-value": "F12",
  "Label Critical-value": "H12",
  "Measurement-label": "B13",
  "Measurement Minor-value": "D13",
  "Measurement Major-value": "F13",
  "Measurement Critical-value": "H13",
  "Trim/Findings-label": "B14",
  "Trim/Findings Minor-value": "D14",
  "Trim/Findings Major-value": "F14",
  "Trim/Findings Critical-value": "H14",
  "Workmanship-label": "B15",
  "Workmanship Minor-value": "D15",
  "Workmanship Major-value": "F15",
  "Workmanship Critical-value": "H15",
  "Product Disposition Details-label": "B19",
  "Disposition Type-label": "B24",
}
const CSVHeader = "File Path,Parse Error,Audit ID,Audit Date,Department,REI Style Number,Audit Level,Auditor,Season,Product Name,Audit Quality Level,Audit Type,Vendor,GA Product Number,Audit Lot Size,Production Status,Factory,Product Spec,Audit Sample Quantity,PO Number,Product Lifecycle,Audit Reject Quantity,Audit Accept Amount,Product Disposition Details,Critical Minor,Critical Major,Critical Critical,Critical RSI,Fabric Minor,Fabric Major,Fabric Critical,Fabric RSI,Hangtag Minor,Hangtag Major,Hangtag Critical,Hangtag RSI,Label Minor,Label Major,Label Critical,Label RSI,Measurement Minor,Measurement Major,Measurement Critical,Measurement RSI,Packaging Minor,Packaging Major,Packaging Critical,Packaging RSI,Packing Minor,Packing Major,Packing Critical,Packing RSI,Trim/Findings Minor,Trim/Findings Major,Trim/Findings Critical,Trim/Findings RSI,Vendor Spec Deviation Minor,Vendor Spec Deviation Major,Vendor Spec Deviation Critical,Vendor Spec Deviation RSI,Workmanship Minor,Workmanship Major,Workmanship Critical,Workmanship RSI,REI Spec Inaccuracy Minor,REI Spec Inaccuracy Major,REI Spec Inaccuracy Critical,REI Spec Inaccuracy RSI"
const sheetName = "REI CFA Audit Result Form"

// const folder = "/Users/simon/Downloads/CFA Data files_20170526175935"
const option = {
  cellFormula: false,
  cellHTML: false,
  cellText: true,
}

function getValue(sheet, data) {
  if (/^[^ ]+ (Minor|Major|Critical)$/.test(data)) {
    assert.equal(data.split(" ")[0], sheet[structure[`${data.split(" ")[0]}-label`]].v.trim())
  } else if (data === "Audit Date") {
    assert.equal(sheet[structure[`${data}-label`]].v, "Audit Date (MM/DD/YYYY)")
  } else if (data === "Product Name") {
    const label = sheet[structure[`${data}-label`]].v
    const value = sheet[structure[`${data}-value`]].v
    assert.equal(label === data || label === value, true)
  } else {
    assert.equal(data, sheet[structure[`${data}-label`]].v)
  }
  const cell = sheet[structure[`${data}-value`]]
  if (!cell) {
    return ""
  }
  if (data === "Audit Date" && cell.t === "n" && cell.w) {
    return cell.w
  }
  return cell.v
}

function getProductDD(sheet) {
  const maxRow = /(\d+)/.exec(sheet["!ref"].split(":")[1])[1]
  const labelRow = /(\d+)/.exec(structure["Disposition Type-label"])[1]
  const labelColumn = /([A-Z]+)/i.exec(structure["Disposition Type-label"])[1]
  const result = []
  for (let i = parseInt(labelRow, 10) + 1; i <= maxRow; i += 1) {
    if (sheet[`${labelColumn}${i}`] && sheet[`${labelColumn}${i}`].v) {
      if (!sheet[`${String.fromCharCode(labelColumn.charCodeAt() + 1)}${i}`]
        && !sheet[`${String.fromCharCode(labelColumn.charCodeAt() + 2)}${i}`]) {
          result.push(sheet[`${labelColumn}${i}`].v)
      }
    }
  }
  return result.join(";")
}

function sheetToCSVArray(file, sheet) {
  if (!sheet) {
    throw new Error("sheet must not be empty")
  }
  const result = [file, ""]
  const errors = []
  const fields = CSVHeader.split(",")
  for (let i = 0; i < fields.length; i += 1) {
    try {
      switch (true) {
          /* eslint-disable indent */
        case fields[i] === "File Path":
        case fields[i] === "Parse Error":
          break
        case fields[i] === "Product Disposition Details":
          result.push(getProductDD(sheet))
          break
        case fields[i] === "Audit ID":
        case fields[i] === "Audit Reject Quantity":
        case /^[^ ]+ RSI$/.test(fields[i]):
        case /^(Packaging|Packing|Vendor Spec Deviation|REI Spec Inaccuracy) (Minor|Major|Critical|RSI)$/.test(fields[i]):
          result.push("")
          break
        case fields[i] === "PO Number":
          result.push(`PO Number: ${getValue(sheet, fields[i])}`)
          break
        default:
          result.push(getValue(sheet, fields[i]))
          /* eslint-enable indent */
      }
    } catch (err) {
      errors.push(fields[i])
      result.push("")
    }
  }
  if (errors.length > 0) {
    result[1] = `Can not extract info: [ ${errors.join(",")}]`
  }
  return [result]
}

function extract(file) {
  const obj = { error: null, array: null }
  const extname = path.extname(file)
  if (/^\.(xls|xlsx)$/i.test(extname)) {
    try {
      const sheet = XLSX.readFile(file, option).Sheets[sheetName]
      if (!sheet) {
        obj.error = `can not find worksheet: ${sheetName}`
      } else {
        obj.array = sheetToCSVArray(sheet)
      }
      return obj
    } catch (e) {
      obj.error = e.message
      return obj
    }
  } else {
    obj.error = "unsupport file format."
    return obj
  }
}

function getEncoding() {
  let encoding = "utf8"
  if (os.platform() === "win32" && /^6\.1\.760[01]$/.test(os.release())) {
    encoding = "GB2312"
  }
  return encoding
}

function writeCSVHeader(ws, encoding = "utf8") {
  const coding = encoding.toLowerCase()
  const out = `${CSVHeader}\r\n`
  if (coding === "utf8") {
    ws.write(new Buffer("\xEF\xBB\xBF", "binary"))
    ws.write(out)
  } else {
    ws.write(iconv.encode(out, coding))
  }
}

function writeToCSV(ws, array, encoding = "utf8") {
  const coding = encoding.toLowerCase()
  const data = `${papa.unparse(array, { header: false })}\r\n`
  if (coding === "utf8") {
    ws.write(data)
  } else {
    ws.write(iconv.encode(data, coding))
  }
}

function globFolder(folder) {
  return new Promise((resolve) => {
    glob.glob(`${folder}/**/*.*`, { nocase: true }, (err, files) => {
      resolve(files)
    })
  })
}

module.exports = {
  extract,
  writeCSVHeader,
  writeToCSV,
  globFolder,
  getEncoding,
}
