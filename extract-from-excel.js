const XLSX = require("xlsx")
const assert = require("assert")
const path = require("path")

const { CSVHeader, sheetStructure, sheetName } = require("./constants.js")

// const folder = "/Users/simon/Downloads/CFA Data files_20170526175935"
const option = {
  cellFormula: false,
  cellHTML: false,
  cellText: true,
}

function getValue(sheet, data) {
  if (/^[^ ]+ (Minor|Major|Critical)$/.test(data)) {
    assert.equal(data.split(" ")[0], sheet[sheetStructure[`${data.split(" ")[0]}-label`]].v.trim())
  } else if (data === "Audit Date") {
    assert.equal(sheet[sheetStructure[`${data}-label`]].v, "Audit Date (MM/DD/YYYY)")
  } else if (data === "Product Name") {
    const label = sheet[sheetStructure[`${data}-label`]].v
    const value = sheet[sheetStructure[`${data}-value`]].v
    assert.equal(label === data || label === value, true)
  } else {
    assert.equal(data, sheet[sheetStructure[`${data}-label`]].v)
  }
  const cell = sheet[sheetStructure[`${data}-value`]]
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
  const labelRow = /(\d+)/.exec(sheetStructure["Disposition Type-label"])[1]
  const labelColumn = /([A-Z]+)/i.exec(sheetStructure["Disposition Type-label"])[1]
  const result = []
  for (let i = parseInt(labelRow, 10) + 1; i <= maxRow; i += 1) {
    if (sheet[`${labelColumn}${i}`]
      && sheet[`${labelColumn}${i}`].v
      && sheet[`${String.fromCharCode(labelColumn.charCodeAt() + 3)}${i}`]
      && sheet[`${String.fromCharCode(labelColumn.charCodeAt() + 3)}${i}`].v) {
        result.push(sheet[`${labelColumn}${i}`].v)
    }
  }
  return result.join(";")
}

function sheetToCSVArray(file, sheet) {
  if (!sheet) {
    console.log(sheet)
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
  assert.equal(/^\.(xls|xlsx)$/i.test(extname), true)
  try {
    const sheet = XLSX.readFile(file, option).Sheets[sheetName]
    if (!sheet) {
      obj.error = `can not find worksheet: ${sheetName}`
    } else {
      obj.array = sheetToCSVArray(file, sheet)
    }
  } catch (e) {
    obj.error = e.message
  }
  return obj
}

module.exports = {
  extract,
}
