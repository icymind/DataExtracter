/* global $ */
/* eslint no-param-reassign: 0 */
const { app, dialog } = require("electron").remote
const fs = require("fs")
const path = require("path")
const homedir = require("os").homedir()
const extractor = require("./extractor.js")
const { protectedFilesSaveAs, execVBS, cleanUp } = require("./helper.js")

let abortExtracting = false

function processFile(file, ws, encoding, protectedFiles, type = "normal") {
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
      if (type === "ppf") {
        const dir = path.dirname(path.dirname(file))
        obj.array[0] = path.join(dir, path.basename(file))
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

function resetSteps() {
  document.getElementById("done-steps").disabled = false
  document.getElementById("abort-steps").disabled = true
  abortExtracting = false
  Array.from(document.querySelectorAll(".ui.steps .step")).forEach((step) => {
    step.classList.add("disabled")
    step.classList.remove("completed")
    step.previousElementSibling.classList.remove("loading")
    step.querySelector(".description").innerHTML = ""
  })
}

function activeStep(id) {
  const div = document.getElementById(id)
  div.classList.remove("disabled")
  if (id.indexOf("summary") < 0) {
    div.previousElementSibling.classList.add("loading")
  }
}

function completedStep(id) {
  const div = document.getElementById(id)
  div.classList.add("completed")
  div.previousElementSibling.classList.remove("loading")
}

function toggleStepsModal(action) {
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
    toggleStepsModal("hide")
  })

  document.getElementById("process-btn").addEventListener("click", async () => {
    const srcFolder = document.getElementById("src-folder-btn").previousElementSibling.value
    const saveFolder = document.getElementById("save-folder-btn").previousElementSibling.value
    const files = await extractor.globFolder(srcFolder)
    if (!files) return

    resetSteps()
    toggleStepsModal("show")

    const ws = fs.createWriteStream(path.join(saveFolder, "data-extractor-out.csv"))
    const encoding = extractor.getEncoding()
    extractor.writeCSVHeader(ws, encoding)

    const span = document.getElementById("processed-indicate")
    const ppfspan = document.getElementById("resaved-ppf-indicate")
    const len = files.length
    let ppfLen = 0
    const protectedFiles = []
    let processedCounter = 0

    activeStep("extracting-step")
    const startTime = Date.now()
    for (let i = 0; i < len; i += 1) {
      if (abortExtracting) {
        return
      }
      const error = await processFile(files[i], ws, encoding, protectedFiles)
      if (/file is password-protected/i.test(error)) {
        ppfLen += 1
        ppfspan.innerHTML = `0/${ppfLen}`
      }
      processedCounter += 1
      span.innerHTML = `${processedCounter}/${len}`
    }
    completedStep("extracting-step")

    activeStep("resaving-protected-files-step")
    await protectedFilesSaveAs(protectedFiles)
    const noppf = await execVBS()
    const remainppf = []
    processedCounter = 0
    for (let i = 0; i < noppf.length; i += 1) {
      if (abortExtracting) {
        return
      }
      const error = await processFile(noppf[i], ws, encoding, remainppf, "ppf")
      if (error) {
        console.log(error)
      }
      processedCounter += 1
      ppfspan.innerHTML = `${processedCounter}/${ppfLen}`
    }
    completedStep("resaving-protected-files-step")

    activeStep("summary-step")
    document.querySelector("#summary-step .description").innerHTML = `${(Date.now() - startTime) / 1000} Seconds`
    ws.end()
    await cleanUp()
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
console.log("te")
