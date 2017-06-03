const os = require("os")
const path = require("path")
const fs = require("fs-extra")
const globFolder = require("./extractor.js").globFolder

const xlsxFolder = "protect-xlsx"
// const saveAsFolder = "saveas"
const xlsm = path.join(__dirname, "helper.xlsm")
const template = `
Option Explicit

On Error Resume Next

Helper

Sub Helper()

  Dim xlApp
  Dim xlBook
  Dim objShell

  Set xlApp = CreateObject("Excel.Application")
  Set xlBook = xlApp.Workbooks.Open("${path.join(os.tmpdir(), "helper.xlsm")}", 0, True)
  xlApp.Run "Main"
  xlApp.Quit

  Set xlBook = Nothing
  Set xlApp = Nothing

End Sub
`
function generate() {
  const p = path.join(os.tmpdir(), "helper.vbs")
  return new Promise((resolve, reject) => {
    fs.writeFile(p, template, (err) => {
      if (err) {
        reject(err)
      } else {
        resolve()
      }
    })
  })
}

async function copyProtectedFiles(files) {
  for (let i = 0, len = files.length; i < len; i += 1) {
    await fs.copy(files[i], path.join(os.tmpdir(), xlsxFolder, path.basename(files[i])))
  }
  // await fs.mkdirs(path.join(os.tmpdir(), xlsxFolder, saveAsFolder))
}

function copyXLSM() {
  return fs.copy(xlsm, path.join(os.tmpdir(), path.basename(xlsm)))
}

async function protectedFilesSaveAs(f) {
  let files = f
  if (typeof f === "string") {
    files = await globFolder(f)
  }
  await copyProtectedFiles(files)
  await generate()
  await copyXLSM()
}
console.log(os.tmpdir())
protectedFilesSaveAs("Z:\\protect")

module.exports = {
  protectedFilesSaveAs,
}
