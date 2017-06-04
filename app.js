/* global $ */
/* eslint no-param-reassign: 0 */
const { app, dialog } = require("electron").remote
const fs = require("fs")
const path = require("path")
const homedir = require("os").homedir()
const extractor = require("./extractor.js")
const { protectedFilesSaveAs, execVBS, cleanUp, getSavetoPath } = require("./helper.js")

let abortExtracting = false

function intervalCheckSavedProcess(folderPath, max) {
  const div = document.getElementById("resaved-ppf-indicate")
  const id = setInterval(async () => {
    if (div.innerHTML[div.innerHTML.length - 1] === " ") {
      return
    }
    const currentFiles = await extractor.globFolder(folderPath, "*.@(xls|xlsx)")
    div.innerHTML = `${currentFiles.length}`
    if (currentFiles.length === max) {
      div.innerHTML = `${div.innerHTML} `
      clearInterval(id)
    }
  }, 1800)
}

function processFile(file, ws, encoding, protectedFiles, newOldMap) {
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
      if (newOldMap) {
        obj.array[0][0] = newOldMap.get(file)
      }
      if (!obj.error) {
        extractor.writeToCSV(ws, obj.array, encoding)
      } else if (obj.error === "File is password-protected") {
        protectedFiles.push(file)
      } else {
        // obj.error = `can not find worksheet: ${sheetName}`
        extractor.writeToCSV(ws, [[file, obj.error]], encoding)
      }
      resolve(obj.error)
    }, 0)
  })
}

function resetSteps() {
  document.querySelector(".ui.message").style.visibility = "hidden"
  document.getElementById("done-steps").disabled = true
  document.getElementById("abort-steps").disabled = false
  abortExtracting = false
  Array.from(document.querySelectorAll(".ui.steps .step")).forEach((step) => {
    // step.style.display = "none"
    step.classList.add("disabled")
    step.classList.remove("completed")
    step.querySelector("i").classList.remove("loading")
    const description = step.querySelector(".description")
    const children = description.querySelectorAll("span")
    if (children.length > 0) {
      children.forEach(child => child.innerHTML = "0")
    } else {
      description.innerHTML = ""
    }
  })
}

function activeStep(id) {
  const div = document.getElementById(id)
  // div.style.display = null
  div.classList.remove("disabled")
  div.querySelector("i").classList.add("loading")
}

function completedStep(id) {
  const div = document.getElementById(id)
  div.classList.add("completed")
  div.querySelector("i").classList.remove("loading")
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

  document.getElementById("src-folder-btn").previousElementSibling.value = "\\\\Mac\\Downloads\\protect"
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
    const outputName = "data-extractor-out.csv"
    if (!files) return

    resetSteps()
    toggleStepsModal("show")

    activeStep("read-folder")
    const len = files.length
    document.getElementById("total-files").innerHTML = len
    completedStep("read-folder")

    const ws = fs.createWriteStream(path.join(saveFolder, outputName))
    const encoding = extractor.getEncoding()
    extractor.writeCSVHeader(ws, encoding)

    const span = document.getElementById("processed-indicate")
    const ppfspan = document.getElementById("resaved-ppf-indicate")
    const ppfAmount = document.getElementById("ppf-amount")
    const eppfspan = document.getElementById("processed-ppf-indicate")
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
        ppfAmount.innerHTML = `${ppfLen}`
      } else {
        processedCounter += 1
        span.innerHTML = processedCounter
      }
    }
    completedStep("extracting-step")

    activeStep("resaving-ppf-step")
    document.getElementById("abort-steps").disabled = true
    const newOldMap = await protectedFilesSaveAs(protectedFiles)
    intervalCheckSavedProcess(getSavetoPath(), protectedFiles.length)
    const noppf = await execVBS()

    // activeStep("extracting-ppf-step")
    const remainppf = []
    processedCounter = 0
    for (let i = 0; i < noppf.length; i += 1) {
      if (abortExtracting) {
        return
      }
      const error = await processFile(noppf[i], ws, encoding, remainppf, newOldMap)
      if (error) {
        console.log(error)
      }
      processedCounter += 1
      eppfspan.innerHTML = `${processedCounter}`
    }
    // completedStep("extracting-ppf-step")
    completedStep("resaving-ppf-step")

    document.querySelector(".ui.message").style.visibility = "visible"
    const timeTaken = (Date.now() - startTime) / 1000
    document.getElementById("time-taken").innerHTML = `Time Taken: ${timeTaken} Seconds`
    document.getElementById("summary-info").innerHTML = `All infomation (include failure extractions) has been saved to ${path.join(saveFolder, outputName)}`
    ws.end()
    document.getElementById("done-steps").disabled = false
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
// console.log(require("os").tmpdir())
