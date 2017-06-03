/* global $ */
const { app, dialog } = require("electron").remote
const fs = require("fs")
const path = require("path")
const homedir = require("os").homedir()
const extracter = require("./extract-from-excel.js")

function processFile(file, ws, encoding, protectedFiles) {
  return new Promise((resolve) => {
    setTimeout(() => {
      const { error, array } = extracter.extract(file)
      if (!error) {
        extracter.writeToCSV(ws, array, encoding)
      } else if (error === "File is password-protected") {
        protectedFiles.push(file)
      } else {
        extracter.writeToCSV(ws, [[file, error]], encoding)
      }
      resolve(error)
    }, 0)
  })
}

function toggleSteps(action) {
  if (action === "show") {
    document.getElementById("done-steps").disabled = true
    document.getElementById("summary-step").style.display = "none"
    document.querySelector("#extracting-step i").classList.add("loading")
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
    const files = await extracter.globFolder(srcFolder)
    if (!files) return
    toggleSteps("show")
    const ws = fs.createWriteStream(path.join(saveFolder, "data-extracter-out.csv"))
    const encoding = extracter.getEncoding()
    extracter.writeCSVHeader(ws, encoding)
    const protectedFiles = []
    let processedCounter = 0
    const span = document.getElementById("processed-indicate")
    const len = files.length
    const startTime = Date.now()
    for (let i = 0; i < len; i += 1) {
      const error = await processFile(files[i], ws, encoding, protectedFiles)
      processedCounter += 1
      span.innerHTML = `${processedCounter}/${len}`
    }
    document.querySelector("#extracting-step i").classList.remove("loading")
    document.querySelector("#summary-step .description").innerHTML = `${(Date.now() - startTime) / 1000} Seconds`
    document.getElementById("summary-step").style.display = null
    document.getElementById("done-steps").disabled = false
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
})
