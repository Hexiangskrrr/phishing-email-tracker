const express = require("express")
const app = express()
const XLSX = require("xlsx")
const { v4: uuidv4 } = require("uuid")
const port = process.env.PORT || 3000

function loadWorkbook() {
  const workbook = XLSX.readFile("list.xlsx")
  const sheetName = workbook.SheetNames[0]
  const worksheet = workbook.Sheets[sheetName]
  const list = XLSX.utils.sheet_to_json(worksheet)


  list.forEach((row) => {
    if (!row.uuid) {
      row.uuid = uuidv4()
    }
  })

  const updatedWorksheet = XLSX.utils.json_to_sheet(list)
  workbook.Sheets[sheetName] = updatedWorksheet
  XLSX.writeFile(workbook, "list.xlsx")

  return { workbook, worksheet, list, sheetName }
}

function saveListToExcel(list, workbook, sheetName) {
  const updatedWorksheet = XLSX.utils.json_to_sheet(list)
  workbook.Sheets[sheetName] = updatedWorksheet
  XLSX.writeFile(workbook, "list.xlsx")
}

let { workbook, worksheet, list, sheetName } = loadWorkbook()

app.get("/:uuid", (req, res) => {
    const uuid = req.params.uuid
    const person = list.find((row) => row.uuid === uuid)
  
    if (!person) {
      return res.status(404).send("Invalid link.")
    }
  
    const timestamp = new Date().toLocaleString("en-SG", {
      weekday: undefined,
      year: "numeric",
      month: "long",
      day: "numeric",
      hour: "numeric",
      minute: "2-digit",
      second: "2-digit",
      hour12: true,
      timeZone: "Asia/Singapore",
    })
  
    person.clickedAt = timestamp
  
    saveListToExcel(list, workbook, sheetName)
  
    res.send(`${person.name}`)
  })
  

app.get("/", (req, res) => {
  res.send("server is working")
})

app.listen(port, () => {
  console.log(`Server listening on port ${port}`)
  console.log("Here are the tracking links:")
  list.forEach((row) => {
    console.log(`${row.name}: http://localhost:${port}/${row.uuid}`)
  })
})
