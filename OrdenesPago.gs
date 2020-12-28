/**
 * Adapted from:
 * https://yagisanatode.com/2020/06/13/google-apps-script-extract-specific-data-from-a-pdf-and-insert-it-into-a-google-sheet/
 * https://gist.github.com/sparkalow/8113e8b7e5c518569c19684ed1d786fb
 */
 

// Globals
var ss = SpreadsheetApp.getActiveSpreadsheet()
var mainFolderId = Browser.inputBox("Ingrese el Id de la carpeta que contiene los archivos")
var sheetCancelaciones = ss.getSheetByName(Browser.inputBox("Ingrese el nombre de la hoja donde desee agregar los datos de las cancelaciones"))
var sheetDebitos = ss.getSheetByName(Browser.inputBox("Ingrese el nombre de la hoja donde desee agregar los datos de los debitos"))
// Stores all data as bidimentional arrays.
var allDataCancelaciones = []
var allDataDebitos = []


/**
 * Main function.
 */
function extractData(){

  // Process files on main folder.
  processFiles(mainFolderId)

  var folders = DriveApp.getFolderById(mainFolderId).getFolders()
  
  // In case main folder has sub folders, process them as well.
  if (folders.hasNext()) {
    processFolders(folders)
  } 
  
  // Once all data is concattenated on allData... variables, imports it to sheets.
  importToSpreadsheet(allDataCancelaciones, sheetCancelaciones)
  importToSpreadsheet(allDataDebitos, sheetDebitos)
  
}


/**
 * Process files on folders and subfolders.
 * @param {FolderIterator}
 */
function processFolders(folders) {
  while (folders.hasNext()) {
    var folder = folders.next()
    var folderId = folder.getId()
    processFiles(folderId)

    // If folder contains subfolders, process them as well
    var subFolder = folder.getFolders()
    processFolders(subFolder)
  }
}

/**
 * Process files on subfolders.
 * @param {string} folderId
 */
function processFiles(folderId){
  var files = DriveApp.getFolderById(folderId).getFilesByType("application/pdf")  
  
  while(files.hasNext()){
    var file = files.next()
    var fileId = file.getId()
    // Uses custom functions to extract data.
    var doc = getTextFromPDF(fileId)
    var data = extractFields(doc.text, doc.name)
    
    // Stores extracted data on array.
    allDataCancelaciones = allDataCancelaciones.concat(data.cancelaciones)
    allDataDebitos = allDataDebitos.concat(data.debitos)
  }
}

/**
 * Extract text from pdf.
 * @param {string} fileId
 * @return {object}
 */
function getTextFromPDF(fileId) {
  var blob = DriveApp.getFileById(fileId).getBlob()
  var resource = {
    title: blob.getName(),
    mimeType: blob.getContentType()
  }
  var options = {
    ocr: true, 
    ocrLanguage: "es"
  }
  // Convert the pdf to a Google Doc with ocr.
  var file = Drive.Files.insert(resource, blob, options);
  
  // Get the texts from the newly created text.
  var doc = DocumentApp.openById(file.id)
  var text = doc.getBody().getText()
  var fileName = doc.getName()
  
  // Deletes the document once the text has been stored.
  Drive.Files.remove(doc.getId())
  
  return {
    name:fileName,
    text:text
  }
}

/**
 * Extact specific fields by matching regular expressions.
 * @param {string} text
 * @param {string} fileName
 * @return bidimentional arrays as object
 */
function extractFields(text, fileName){
  // creates Regex formulas.
  var regexOrdenPago = /(?<=Orden de.+[\r\n])([0-9]{1,})/
  var regexFechaEmision = /(?<=FPD )([\d]{2}\/[\d]{2}\/[\d]{4})/
  var regexCodigoMedico = /(?<=Codigo de medico:.+?)([\d]{1,})/
  var regexBeneficiario = /(?<=Beneficiario:\s+)(.+\b)/
  var regexCancelaciones = /(?<!D e b i t o s.+[\r\n].+)([\d ]{4,})\s*([^0-9\/]+)\s*(\d{2},\d{2})?\s*(\d{1,2}\/\d{4})\s* (\d{1,}\.*\d{1,3},\d{2})/g
  var regexDebitos = /(?<!C a n c e l a c i o n.+[\r\n].+)([\d]{4,})?\s*([^0-9]+|Retencion de ingresos brutos prov Bs As [0-9]{8})\s*(\d{1,2}\/\d{4})?\s*((?:\d{1,}\.*)?\d{1,3},\d{2})(?=.*Total debito)/g
  // extract desired fields from text using Regex formulas
  var ordenPago = regexOrdenPago.exec(text)
  var fechaEmision = regexFechaEmision.exec(text)
  var codigoMedico = regexCodigoMedico.exec(text)
  var beneficiario = regexBeneficiario.exec(text)
  // multiple match fields need to be converted to objects and then to arrays
  var cancelacionesToObject = text.matchAll(regexCancelaciones)
  var cancelacionesToArray = Array.from(cancelacionesToObject)
  var debitosToObject = text.matchAll(regexDebitos)
  var debitosToArray = Array.from(debitosToObject)
  
  // Concatenates elements into a new array
  try {
    var cancelaciones = cancelacionesToArray.map(function(array){
      return [fileName, ordenPago, fechaEmision, codigoMedico, beneficiario,
              array[1], array[2], array[3], array[4], array[5]]
    })
    }catch(e){
      var cancelaciones = [["Error"]] 
      }
  
  try {
    var debitos = debitosToArray.map(function(array){
      return [fileName,ordenPago, fechaEmision, codigoMedico, beneficiario,
              array[1], array[2], array[3], array[4]]
    })
    }catch(e){
      var debitos = [["Error"]]
      }
  
  return {
    cancelaciones: cancelaciones,
    debitos: debitos
  }
}

/**
 * Move data to Google SpreadSheet.
 * @param {bidimentional array} data
 * @param {string} sheetName
 */
function importToSpreadsheet(data, sheetName){
  const sheet = sheetName
  // Paste data under last non empty row
  const range = sheet.getRange(sheet.getLastRow()+1,1,data.length, data[0].length)
  range.setValues(data)
}