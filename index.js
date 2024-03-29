import XlsxPopulate from "xlsx-populate";

// Using Promises
XlsxPopulate.fromBlankAsync()
.then(fichero => {
  fichero.sheet(0).cell('A1').value('Hello World!')
  // return fichero.toFileAsync('./archivo.xlsx')
})

// Using async-await
async function createBlankFile() {
  const fichero = await XlsxPopulate.fromBlankAsync()
  fichero.sheet(0).cell('A1').value('Hello World!')
  fichero.toFileAsync('./archivo2.xlsx')
}
// createBlankFile()

// Create a DB using async-await
async function createBlankFileDataBase() {
  const workbook = await XlsxPopulate.fromBlankAsync()

  // Row nº1 (headers)
  workbook.sheet(0).cell('A1').value('Name')
  workbook.sheet(0).cell('B1').value('Surname')
  workbook.sheet(0).cell('C1').value('Age')

  // Row nº2 (data)
  workbook.sheet(0).cell('A2').value('Carlos')
  workbook.sheet(0).cell('B2').value('Martínez')
  workbook.sheet(0).cell('C2').value(24)

  // Row nº3 (data)
  workbook.sheet(0).cell('A3').value('Mónica')
  workbook.sheet(0).cell('B3').value('García')
  workbook.sheet(0).cell('C3').value(65)

  workbook.toFileAsync('./archivo3.xlsx')
}
// createBlankFileDataBase()

// Reading a file
async function readACellFromASheet() {
  const workbook = await XlsxPopulate.fromFileAsync('./archivo3.xlsx')
  const a2Value = workbook.sheet('Sheet1').cell('A2').value()
  console.log(a2Value) // === 'Carlos'
}
// readACellFromASheet()

// Reading all the file (range)
async function readARangeFromASheet() {
  const workbook = await XlsxPopulate.fromFileAsync('./archivo3.xlsx')
  const range = workbook.sheet('Sheet1').usedRange().value()
  console.log(range)
}
// readARangeFromASheet()

// Read selecting the range
async function readSpecificRangeFromASheet() {
  const workbook = await XlsxPopulate.fromFileAsync('./archivo3.xlsx')
  const specificRange = workbook.sheet('Sheet1').range('A1:B2').value()
  console.log(specificRange) // === [ [ 'Name', 'Surname' ], [ 'Carlos', 'Martínez' ] ]
}
// readSpecificRangeFromASheet()

// Creating a Sheet, and complete with data using an Array
async function creatingAFileWithVariousRows() {
  const workbook = await XlsxPopulate.fromBlankAsync()
  workbook.sheet(0).cell('A1').value([
    ['Nombre', 'Apellido', 'Edad'],
    ['Juan', 'Perez', 34],
    ['Laia', 'Gomez', 47]
  ])

  workbook.toFileAsync('./archivo4.xlsx')
}
// creatingAFileWithVariousRows()

// Creating a Sheet, and complete with data using JS code
async function creatingAFileUsingJSCode() {
  const workbook = await XlsxPopulate.fromBlankAsync()
  workbook.sheet(0).cell('A1').value([
    [new Date().getDate(), new Date().getMonth()+1, new Date().getFullYear()]
  ])

  workbook.toFileAsync('./archivo5.xlsx')
}
// creatingAFileUsingJSCode()

// Read the file nº5, copying his content, and paste it in another file nº6 with other Sheet called 'Hoja 2'
async function copyingAFileAndCreateOtherFileWithAnotherSheet() {
  const workbook = await XlsxPopulate.fromFileAsync('./archivo5.xlsx')

  // const sheetName = workbook.sheet(0)
  // console.log(sheetName) // <-- sheet data
  // console.log(sheetName.name()) // === 'Sheet1'

  workbook.addSheet('Hoja 2')
  workbook.toFileAsync('./archivo6.xlsx')
}
// copyingAFileAndCreateOtherFileWithAnotherSheet()

// Reading the name of the sheets from a file
async function readVariousSheetsNamesFromAFile() {
  const workbook = await XlsxPopulate.fromFileAsync('./archivo6.xlsx')

  const sheetsNames = workbook.sheets().map(sheet => sheet.name())
  console.log(sheetsNames) // === [ 'Sheet1', 'Hoja 2' ]
}
// readVariousSheetsNamesFromAFile()

// Change sheet name
async function changeSheetName() {
  const workbook = await XlsxPopulate.fromFileAsync('./archivo6.xlsx')

  workbook.sheet('Sheet1').name('Hoja 1')

  workbook.toFileAsync('./archivo7.xlsx') // and save it in another file
}
// changeSheetName()

// Delete a sheet
async function deleteSheet() {
  const workbook = await XlsxPopulate.fromFileAsync('./archivo7.xlsx')

  workbook.deleteSheet('Hoja 2')

  workbook.toFileAsync('./archivo8.xlsx') // and save it in another file
}
// deleteSheet()

// Creating a XLSX file with a password
async function createAFileWithPassword() {
  const workbook = await XlsxPopulate.fromFileAsync('./archivo8.xlsx')

  workbook.toFileAsync('./archivo8-password.xlsx', {
    password: '12345'
  })
}
createAFileWithPassword()