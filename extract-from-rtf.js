/* eslint no-useless-escape: 0 */
const fs = require("fs")
const path = require("path")
const assert = require("assert")
const { CSVHeader, nonconformityType } = require("./constants.js")

const controlCommand = String.raw`(?:\{|\}|\r|\n|\t|(?:\\\n)|(?:\\[a-z]+\d*[ ]?)|(?:\{\\\*\\[a-z]+[ ].+?\}))`
const plainText = String.raw`[^\{\}\n\r\\]+`

function classify(nonconformityTitle) {
  const classifyError = new Error(`unkown nonconformity type. nonconformityTitle: ${nonconformityTitle}`)
  classifyError.name = "classifyError"

  let type = nonconformityTitle.split("-")[1]
  type = type.toLowerCase()
  for (let i = 0, len = nonconformityType.length; i < len; i += 1) {
    if (nonconformityType[i].toLowerCase() === type) {
      return nonconformityType[i]
    }
  }
  throw classifyError
}
function extractSpansBetween(str, pre, post) {
  let pattern = `${pre}([\\s\\S]*?)${post}`
  let reg = new RegExp(pattern, "m")
  const matchParagraph = reg.exec(str)
  assert.equal(matchParagraph !== null, true)
  const paragraph = matchParagraph[1]

  let spans = []
  const indexes = []
  pattern = `${controlCommand}+`
  reg = new RegExp(pattern, "gmi")
  let span = reg.exec(paragraph)
  assert.equal(span !== null, true)

  while (span) {
    // console.log(span)
    indexes.push({ begin: span.index, end: reg.lastIndex })
    span = reg.exec(paragraph)
  }
  // console.log(indexes)
  assert.equal(indexes.length !== 0, true)
  for (let i = 1, len = indexes.length; i < len; i += 1) {
    spans.push(paragraph.substring(indexes[i - 1].end, indexes[i].begin))
  }

  if (spans[0].trim() === ":") {
    spans = spans.slice(1)
  }
  if (/^ *: .+$/.test(spans[0])) {
    spans[0] = /^ *: (.+)$/.exec(spans[0])[1]
  }
  return spans
}
function getAndValidateSpan(str, fieldName, postFieldName, pattern) {
  const multiLine = ["Audit Quality Level", "Audit Lot Size", "Audit Sample Quantity"]
  const spans = extractSpansBetween(str, fieldName, postFieldName)
  if ((spans.length === 1 || multiLine.includes(fieldName)) && pattern.test(spans[0])) {
    return spans[0]
  }
  throw new Error(`Can not extract ${fieldName}`)
}
function getNonconformityDetails(str) {
  const pattern = String.raw`Nonconformity Details[\s\S]+?_{8,}([\s\S]+?)Nonconformity Summary`
  let reg = new RegExp(pattern, "mi")
  const matchParagraph = reg.exec(str)
  assert.equal(matchParagraph !== null, true)
  const paragraph = matchParagraph[1]

  const spansPattern = `(${controlCommand}+)( *${plainText} *)(${controlCommand}+\\s(\\d+)${controlCommand}+\\s(\\d+)${controlCommand}+\\s(\\d+)${controlCommand}+\\s(\\d+)${controlCommand}+\\s(?:\\d+)${controlCommand}+\\s(?:\\d+)${controlCommand}+\\s(?:\\d+))?`

  reg = new RegExp(spansPattern, "gmi")

  const detailsArray = []
  let matchSpans = reg.exec(paragraph)
  while (matchSpans) {
    const controlCommandGroup = matchSpans[1]
    const firstPlainText = matchSpans[2]
    const qtyGroup = matchSpans[3]
    if (qtyGroup) {
      const [nonconformity, minorQTY, majorQTY, CriticalQTY, RSIQTY] = [
        matchSpans[2],
        matchSpans[4],
        matchSpans[5],
        matchSpans[6],
        matchSpans[7],
      ]
      const type = classify(nonconformity)
      const detail = {
        Nonconformity: nonconformity,
        NonconformityType: type,
        QTY: {
          Minor: minorQTY,
          Major: majorQTY,
          Critical: CriticalQTY,
          RSI: RSIQTY,
        },
      }
      detailsArray.push(detail)
    } else if (firstPlainText !== " " && detailsArray.length !== 0) {
      if (controlCommandGroup.includes("tx120")) {
        detailsArray[detailsArray.length - 1].Nonconformity += firstPlainText
      }
    }
    matchSpans = reg.exec(paragraph)
  }
  return detailsArray
}
function getProductDispositionDetails(str) {
  const paragraphPattern = String.raw`Product Disposition Details[\s\S]+?Quantity[\s\S]+?Comments[\s\S]+?(?:_{8,})?([\s\S]+?)(?:(?:Audit Done)|(?:\\\*\\themedata))`
  let reg = new RegExp(paragraphPattern, "mi")
  const matchParagraphs = reg.exec(str)

  assert.equal(matchParagraphs !== null, true)

  const paragraph = matchParagraphs[1]

  const linePattern = String.raw`\b([\s\S]*?)(?:(?:\\par\b)|(?:\\pard\b))`
  reg = new RegExp(linePattern, "gmi")
  let matchLines = reg.exec(paragraph)

  const dispositionArray = []
  let lineTabPosition = []
  while (matchLines) {
    const spanPattern = `(${controlCommand}+?)\\b(${plainText})`
    const spanReg = new RegExp(spanPattern, "gmi")

    const line = matchLines[1]
    let matchSpans = spanReg.exec(line)

    const fieldArray = []
    while (matchSpans) {
      const positions = matchSpans[1].match(/tx\d+/gmi)
      if (positions) {
        lineTabPosition = positions
      }
      if (matchSpans[2] !== " ") {
        fieldArray.push(matchSpans[2])
      }
      matchSpans = spanReg.exec(line)
    }
    const fieldArrayLength = fieldArray.length
    if (fieldArrayLength === 3) {
      assert.equal(parseInt(fieldArray[1], 10), fieldArray[1], "Quantity must be number")
      dispositionArray.push({
        Disposition: fieldArray[0],
        Quantity: fieldArray[1],
      })
    } else if (fieldArrayLength === 2) {
      assert.equal(lineTabPosition.includes("tx90"), true, "unkown file layout")
      if (lineTabPosition.includes("tx3960")) {
        dispositionArray.push({
          Disposition: fieldArray[0],
          Quantity: fieldArray[1],
        })
      } else {
        dispositionArray[dispositionArray.length - 1].Disposition += fieldArray[0]
      }
    } else if (fieldArrayLength === 1 && dispositionArray.length !== 0 && !lineTabPosition.includes("tx4200")) {
      dispositionArray[dispositionArray.length - 1].Disposition += fieldArray[0]
    }
    matchLines = reg.exec(paragraph)
  }
  return dispositionArray
}
function getAuditID(str) {
  const spans = extractSpansBetween(str, "AuditID\\/Date", "Department")
  const idPattern = /^\d+$/
  if (spans.length !== 0 && idPattern.test(spans[0])) {
    return spans[0]
  }
  throw new Error("Can not extract Audit ID")
}
function getAuditDate(str) {
  const spans = extractSpansBetween(str, "AuditID\\/Date", "Department")
  const datePattern = /^((\d{1,2}\/\d{1,2}\/\d{1,4})|(\d{1,4}\/\d{1,2}\/\d{1,2}))$/
  // beautyLog("dates", spans)
  if (spans.length === 1 && datePattern.test(spans[0])) {
    return spans[0]
  } else if (spans.length >= 2) {
    if (datePattern.test(spans.slice(1).join(""))) { return spans.slice(1).join("") }
    if (datePattern.test(spans.join(""))) { return spans.join("") }
  } else {
    throw new Error("Can not extract Audit Date")
  }
  return null
}
function getAuditType(str) {
  return getAndValidateSpan(str, "Audit Type", "Vendor", /^\w+[\w ]*$/)
}
function getProductSpec(str) {
  return getAndValidateSpan(str, "Product Spec", "Audit Sample Quantity", /^v\d+$/i)
}
function getProductLifecycle(str) {
  return getAndValidateSpan(str, "Product Lifecycle", "Audit Reject Quantity", /^[\w \-]+/)
}
function getREIStyleNumber(str) {
  return getAndValidateSpan(str, "REI Style Number", "Audit Level", /^\d+$/)
}
function getAuditLevel(str) {
  return getAndValidateSpan(str, "Audit Level", "Auditor", /^\w+$/)
}
function getAuditor(str) {
  return getAndValidateSpan(str, "Auditor", "Season", /^.+$/)
}
function getSeason(str) {
  return getAndValidateSpan(str, "Season", "Product Name", /^\w+$/)
}
function getDepartment(str) {
  return getAndValidateSpan(str, "Department", "REI Style Number", /^[\S ]+$/)
}
function getAuditQualityLevel(str) {
  return getAndValidateSpan(str, "Audit Quality Level", "Audit Type", /^(\d+|(\d+\.\d+))$/)
}
function getGAProductNumber(str) {
  return getAndValidateSpan(str, "GA Product Number", "Audit Lot Size", /^[\w\-]+$/)
}
function getAuditLotSize(str) {
  return getAndValidateSpan(str, "Audit Lot Size", "Production Status", /^(\d+|(\d+\.\d+))$/)
}
function getProductionStatus(str) {
  return getAndValidateSpan(str, "Production Status", "(?:(?:Factory)|(?:Non-GA Vendor Name))", /^[\w \-]+$/)
}
function getAuditSampleQuantity(str) {
  return getAndValidateSpan(str, "Audit Sample Quantity", "PO Number", /^\d+$/)
}
function getPONumber(str) {
  const poNumberSpans = extractSpansBetween(str, "PO Number", "(?:(?:Product Lifecycle)|(?:Non-GA Vendor Number))")
  // beautyLog("poNumberSpans", poNumberSpans)
  const poNumer = poNumberSpans.join("")
  return poNumer.replace(/\\\\/, "\\")
}
function getAuditRejectQuantity(str) {
  return getAndValidateSpan(str, "Audit Reject Quantity", "Nonconformity Details", /^\d+$/)
}
function getProductName(str) {
  let productName = getAndValidateSpan(str, "Product Name", "Audit Quality Level", /^.*$/)
  const productNamePart = extractSpansBetween(str, "Audit Quality Level", "Audit Type")
  if (productNamePart.length !== 0) {
    productName += productNamePart.slice(1).join("")
  }
  return productName
}
function getVendor(str) {
  let vendor = getAndValidateSpan(str, "Vendor", "GA Product Number", /^\w.*$/)
  const vendorPart = extractSpansBetween(str, "Audit Lot Size", "Production Status")
  if (vendorPart.length !== 0) {
    vendor += vendorPart.slice(1).join("")
  }
  return vendor
}
function getFactory(str) {
  let factory = getAndValidateSpan(str, "Factory", "Product Spec", /^\w.*$/)
  const factoryPart = extractSpansBetween(str, "Audit Sample Quantity", "PO Number")
  if (factoryPart.length !== 0) {
    factory += factoryPart.slice(1).join("")
  }
  return factory
}

function strToObj(str) {
  /* eslint no-eval: 0 */
  // todo:一旦有某个开始不匹配, 则修改 error 属性
  const rtf = { "Parse Error": "" }
  const errorMsg = []
  CSVHeader.forEach((field) => {
    // for (let field in fields) {
    if (!["Parse Error", "File Path"].includes(field)) {
      try {
        rtf[field] = eval(`get${field.split(" ").join("")}`)(str)
      } catch (err) {
        rtf[field] = ""
        errorMsg.push(field)
      }
    }
  })
  if (errorMsg.length !== 0) {
    rtf["Parse Error"] = `Can not extract info: [ ${errorMsg.join("; ")} ]`
  }
  return rtf
}

function combinePDD(str) {
  const ddArray = getProductDispositionDetails(str)
  const temp = []
  ddArray.forEach(detail => temp.push(detail.Disposition))
  return temp.join("; ")
}

function combineND(str) {
  const result = new Map()
  const ND = getNonconformityDetails(str)
  const subCata = ["Minor", "Major", "Critical", "RSI"]
  ND.forEach((detail) => {
    subCata.forEach((cata) => {
      const key = `${detail.NonconformityType} ${cata}`
      if (!result.has(key)) {
        result.set(key, detail.QTY[cata])
      } else {
        const pre = result.get(key)
        result.set(key, detail.QTY[cata] + pre)
      }
    })
  })
  return result
}

function stringToCSVArray(file, str) {
  if (!str) {
    throw new Error("str must not be empty")
  }
  const result = [file, ""]
  const errors = []
  const fields = CSVHeader.split(",")
  let nonconformityDetails = new Set()
  try {
    nonconformityDetails = combineND(str)
  } catch (error) {
    errors.push("nonconformityDetails")
  }
  for (let i = 0; i < fields.length; i += 1) {
    try {
      switch (true) {
          /* eslint-disable indent */
        case fields[i] === "File Path":
        case fields[i] === "Parse Error":
          break
        case fields[i] === "Product Disposition Details":
          result.push(combinePDD(str))
          break
        case fields[i] === "Audit Accept Amount":
          result.push("")
          break
        case fields[i] === "PO Number":
          result.push(`PO Number: ${getPONumber(str)}`)
          break
        case /^.* (Minor|Major|Critical|RSI)$/.test(fields[i]):
          // if contains push(value)
          if (nonconformityDetails.has(fields[i])) {
            result.push(nonconformityDetails.get(fields[i]))
          } else {
            result.push("")
          }
          break
        default:
          result.push(eval(`get${fields[i].split(" ").join("")}`)(str))
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
  const extname = path.extname(file)
  assert.equal(/^\.rtf$/i.test(extname), true)
  const obj = { error: null, array: null }
  let str
  try {
    str = fs.readFileSync(file)
  } catch (e) {
    obj.error = e.message
    return obj
  }
  obj.array = stringToCSVArray(file, str)
  return obj
}

module.exports = {
  extract,
}
