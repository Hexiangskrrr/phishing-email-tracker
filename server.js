const express = require("express")
const multer = require("multer")
const XLSX = require("xlsx")
const { v4: uuidv4 } = require("uuid")
const path = require("path")

const app = express()
const PORT = process.env.PORT || 3000

let list = []
let workbook, worksheet, sheetName, filename

const upload = multer({ dest: "uploads/" })

function loadWorkbook(filePath) {
  workbook = XLSX.readFile(filePath)
  sheetName = workbook.SheetNames[0]
  worksheet = workbook.Sheets[sheetName]
  list = XLSX.utils.sheet_to_json(worksheet)

  list.forEach((row) => {
    if (!row.uuid) {
      row.uuid = uuidv4()
    }
  })

  const updatedWorksheet = XLSX.utils.json_to_sheet(list)
  workbook.Sheets[sheetName] = updatedWorksheet
  XLSX.writeFile(workbook, filePath)
}

function saveListToExcel(filePath) {
  const updatedWorksheet = XLSX.utils.json_to_sheet(list)
  workbook.Sheets[sheetName] = updatedWorksheet
  XLSX.writeFile(workbook, filePath)
}

app.post("/upload", upload.single("file"), (req, res) => {
  const filePath = req.file.path
  filename = req.file.originalname.replace(/\.xlsx$/, "") + "-result.xlsx"

  loadWorkbook(filePath)
  res.send(`Upload successful. Download result to get link`)
  console.log("Upload received")
})

app.get("/download", (req, res) => {
  if (!workbook) return res.status(400).send("No file uploaded yet")

  const filePath = path.join(__dirname, filename)
  saveListToExcel(filePath)

  res.download(filePath, (err) => {
    if (err) console.error(err)
  })
})

app.get("/:uuid", (req, res) => {
  if (!list.length) return res.status(400).send("No data")

  const uuid = req.params.uuid
  const person = list.find((row) => row.uuid === uuid)

  if (!person) return res.status(404).send("Invalid link")

  const timestamp = new Date().toLocaleString("en-SG", {
    year: "numeric", month: "long", day: "numeric",
    hour: "numeric", minute: "2-digit", second: "2-digit",
    hour12: true, timeZone: "Asia/Singapore"
  })

  person.clickedAt = timestamp
  res.send(`${person.name} clicked at ${timestamp}`)
})

app.get("/", (req, res) => {
  res.send(`
    <h2>Upload & Download</h2>
    <form method="POST" enctype="multipart/form-data" action="/upload">
      <input type="file" name="file" accept=".xlsx" />
      <button type="submit">Upload</button>
    </form>
    <a href="/download">
  <button type="button">Download results</button>
</a>

  `)
})

app.listen(PORT, () => {
  console.log(`Server running on ${PORT}`)
})
