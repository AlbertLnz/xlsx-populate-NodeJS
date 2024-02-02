import XlsxPopulate from "xlsx-populate";

// Using Promises
XlsxPopulate.fromBlankAsync()
.then(fichero => {
  fichero.sheet(0).cell('A1').value('Hello World!')
  return fichero.toFileAsync('./archivo.xlsx')
})

// Using async-await
async function createBlankFile() {
  const  fichero = await XlsxPopulate.fromBlankAsync()
  fichero.sheet(0).cell('A1').value('Hello World!')
  fichero.toFileAsync('./archivo2.xlsx')
}
createBlankFile()