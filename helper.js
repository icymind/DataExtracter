const os = require("os")
const path = require("path")
const fs = require("fs-extra")
const globFolder = require("./extractor.js").globFolder
const { exec } = require("child_process")

const xlsxFolderName = "protect-xlsx"
const xlsmFname = "helper.xlsm"
const vbsFname = "helper.vbs"
const saveAsFolderName = "saveas"
const template = `
Option Explicit

On Error Resume Next

Helper

Sub Helper()

  Dim xlApp
  Dim xlBook
  Dim objShell

  Set xlApp = CreateObject("Excel.Application")
  Set xlBook = xlApp.Workbooks.Open("${path.join(os.tmpdir(), xlsmFname)}", 0, True)
  xlApp.Run "Main"
  xlApp.Quit

  Set xlBook = Nothing
  Set xlApp = Nothing

End Sub
`

function getSavetoPath() {
  return path.join(os.tmpdir(), xlsxFolderName, saveAsFolderName)
}

function generate(filePath) {
  return new Promise((resolve, reject) => {
    fs.writeFile(filePath, template, (err) => {
      if (err) {
        reject(err)
      } else {
        resolve()
      }
    })
  })
}

async function copyProtectedFiles(files, destFolder) {
  await fs.remove(destFolder)
  const newOldMap = new Map()
  for (let i = 0, len = files.length; i < len; i += 1) {
    const newPath = path.join(destFolder, path.basename(files[i]))
    newOldMap.set(path.join(destFolder, saveAsFolderName, path.basename(files[i])), files[i])
    await fs.copy(files[i], newPath)
  }
  return newOldMap
}

function copyXLSM(filePath, destFolder) {
  return fs.copy(filePath, path.join(destFolder, path.basename(filePath)))
}

async function protectedFilesSaveAs(f) {
  let files = f
  if (typeof f === "string") {
    files = await globFolder(f)
  }
  const newOldMap = await copyProtectedFiles(files, path.join(os.tmpdir(), xlsxFolderName))
  await generate(path.join(os.tmpdir(), vbsFname))
  await copyXLSM(path.join(__dirname, xlsmFname), os.tmpdir())
  return newOldMap
}

async function execVBS() {
  return new Promise((resolve, reject) => {
    exec(path.join(os.tmpdir(), vbsFname), (err, stdout, stderr) => {
      if (err) {
        reject(err)
      } else {
        const files = globFolder(path.join(os.tmpdir(), xlsxFolderName, saveAsFolderName))
        resolve(files)
      }
    })
  })
}

async function cleanUp() {
  await fs.remove(path.join(os.tmpdir(), vbsFname))
  await fs.remove(path.join(os.tmpdir(), xlsmFname))
  await fs.remove(path.join(os.tmpdir(), xlsxFolderName))
}
// console.log(os.tmpdir())
// protectedFilesSaveAs("Z:\\protect")

module.exports = {
  protectedFilesSaveAs,
  execVBS,
  cleanUp,
  getSavetoPath,
}
