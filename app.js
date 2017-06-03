/* global $ */
/* eslint no-param-reassign: 0 */
const { app, dialog } = require("electron").remote
const fs = require("fs")
const path = require("path")
const homedir = require("os").homedir()
const extractor = require("./extractor.js")
const { protectedFilesSaveAs, execVBS } = require("./helper.js")

let abortExtracting = false

function processFile(file, ws, encoding, protectedFiles) {
  return new Promise((resolve) => {
    setTimeout(() => {
      const extname = path.extname(file)
      let obj = {}
      if (/^\.rtf$/i.test(extname)) {
        obj = extractor.extractFromRtf(file)
      } else if (/^\.(xls|xlsx)$/i.test(extname)) {
        obj = extractor.extractFromExcel(file)
      } else {
        obj.error = "unsupported format"
      }
      if (!obj.error) {
        extractor.writeToCSV(ws, obj.array, encoding)
      } else if (obj.error === "File is password-protected") {
        protectedFiles.push(file)
      } else {
        extractor.writeToCSV(ws, [[file, obj.error]], encoding)
      }
      resolve(obj.error)
    }, 0)
  })
}

function toggleSteps(action) {
  if (action === "show") {
    document.getElementById("done-steps").disabled = true
    document.getElementById("abort-steps").disabled = false
    document.getElementById("summary-step").style.display = "none"
    document.querySelector("#extracting-step i").classList.add("loading")
    document.querySelector("#extracting-step").classList.remove("completed")
  }
  $(".extracting.modal").modal({
    dimmerSettings: {
      opacity: 0.2,
    },
    closable: false,
  }).modal(action)
}

document.addEventListener("DOMContentLoaded", () => {
  $(".ui.dropdown").dropdown()

  document.getElementById("version").innerHTML = app.getVersion()

  document.getElementById("src-folder-btn").previousElementSibling.value = "/Users/simon/Downloads/CFA Data files_20170526175935"
  document.getElementById("src-folder-btn").addEventListener("click", (event) => {
    const folder = dialog.showOpenDialog({ properties: ["openDirectory"] })
    if (!folder) return
    event.target.previousElementSibling.value = folder
  })

  document.getElementById("save-folder-btn").previousElementSibling.value = path.join(homedir, "Desktop")
  document.getElementById("save-folder-btn").addEventListener("click", (event) => {
    const folder = dialog.showOpenDialog({ properties: ["openDirectory"] })
    if (!folder) return
    event.target.previousElementSibling.value = folder
  })

  document.getElementById("done-steps").addEventListener("click", () => {
    toggleSteps("hide")
  })
  document.getElementById("process-btn").addEventListener("click", async () => {
    const srcFolder = document.getElementById("src-folder-btn").previousElementSibling.value
    const saveFolder = document.getElementById("save-folder-btn").previousElementSibling.value
    const files = await extractor.globFolder(srcFolder)
    if (!files) return
    toggleSteps("show")
    const ws = fs.createWriteStream(path.join(saveFolder, "data-extractor-out.csv"))
    const encoding = extractor.getEncoding()
    extractor.writeCSVHeader(ws, encoding)
    const protectedFiles = []
    let processedCounter = 0
    let span = document.getElementById("processed-indicate")
    let len = files.length
    const startTime = Date.now()
    for (let i = 0; i < len; i += 1) {
      if (abortExtracting) {
        break
      }
      const error = await processFile(files[i], ws, encoding, protectedFiles)
      processedCounter += 1
      span.innerHTML = `${processedCounter}/${len}`
    }

    const ppfLen = protectedFiles.length
    await protectedFilesSaveAs(protectedFiles)
    const noppf = await execVBS()
    const remainppf = []
    processedCounter = 0
    span = document.getElementById("processed-ppf-indicate")
    len = noppf.length
    for (let i = 0; i < noppf.length; i += 1) {
      if (abortExtracting) {
        break
      }
      const error = await processFile(noppf[i], ws, encoding, remainppf)
      processedCounter += 1
      span.innerHTML = `${processedCounter}/${len}`
    }

    document.querySelector("#extracting-step i").classList.remove("loading")
    document.querySelector("#extracting-step").classList.add("completed")
    document.querySelector("#summary-step .description").innerHTML = `${(Date.now() - startTime) / 1000} Seconds`
    document.getElementById("summary-step").style.display = null
    document.getElementById("done-steps").disabled = false
    document.getElementById("abort-steps").disabled = true
    abortExtracting = false
    ws.end()
  })

  document.getElementById("about").addEventListener("click", () => {
    $(".about.modal")
      .modal({
        dimmerSettings: {
          opacity: 0.2,
        },
      })
      .modal("show")
  })

  document.getElementById("abort-steps").addEventListener("click", () => {
    abortExtracting = true
  })
})
