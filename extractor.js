const glob = require("glob")
const os = require("os")
const papa = require("papaparse")
const iconv = require("iconv-lite")

const extractFromExcel = require("./extract-from-excel.js").extract
const extractFromRtf = require("./extract-from-rtf.js").extract
const { CSVHeader } = require("./constants.js")

function globFolder(folder) {
  return new Promise((resolve) => {
    glob.glob("/**/*.*", { nocase: true, root: folder }, (err, files) => {
      resolve(files)
    })
  })
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

module.exports = {
  globFolder,
  getEncoding,
  writeCSVHeader,
  writeToCSV,
  extractFromExcel,
  extractFromRtf,
}
