import XlsxPopulate from "xlsx-populate";

XlsxPopulate.fromBlankAsync()
.then(fichero => {
  fichero.sheet(0).cell('A1').value('Hello World!')
  return fichero.toFileAsync('./archivo.xlsx')
})
// And execute the command: node index.js