//Inspirado en https://yagisanatode.com/2020/06/13/google-apps-script-extract-specific-data-from-a-pdf-and-insert-it-into-a-google-sheet/
// y en https://gist.github.com/sparkalow/8113e8b7e5c518569c19684ed1d786fb

var ss = SpreadsheetApp.getActiveSpreadsheet()
var allDataCancelaciones = []
var allDataDebitos = []

function extractData(){
  const folderId = Browser.inputBox("Ingrese el Id de la carpeta que contiene los archivos")
  var folder = DriveApp.getFolderById(folderId).getFolders()  
  var sheetCancelaciones = ss.getSheetByName(Browser.inputBox("Ingrese el nombre de la hoja donde desee agregar los datos de las cancelaciones"))
  var sheetDebitos = ss.getSheetByName(Browser.inputBox("Ingrese el nombre de la hoja donde desee agregar los datos de los debitos"))
  
  Logger.log(folder)
  Logger.log(folder.hasNext())
  
  if (folder.hasNext()) {
    processFolder(folder);
  } else {
    Browser.msgBox('Folder not found!');
    var files = DriveApp.getFolderById(folderId).getFilesByType("application/pdf")
    Logger.log(files)
    
    while(files.hasNext()){
      var file = files.next()
      var fileId = file.getId()
      Logger.log(fileId)
      var doc = getTextFromPDF(fileId)
      var data = extractFields(doc.text, doc.name)
      
      allDataCancelaciones = allDataCancelaciones.concat(data.cancelaciones)
      allDataDebitos = allDataDebitos.concat(data.debitos)
    }
    
    importToSpreadsheet(allDataCancelaciones, sheetCancelaciones)
    importToSpreadsheet(allDataDebitos, sheetDebitos)
    
  }
  
}

function processFolder(folder) {
  while (folder.hasNext()) {
    var f = folder.next();
    
    var files = f.getFilesByType("application/pdf")
    Logger.log(files)
    
    while(files.hasNext()){
      var file = files.next()
      var fileId = file.getId()
      Logger.log(fileId)
      var doc = getTextFromPDF(fileId)
      var data = extractFields(doc.text, doc.name)
      
      allDataCancelaciones = allDataCancelaciones.concat(data.cancelaciones)
      allDataDebitos = allDataDebitos.concat(data.debitos)
    }
    var subFolder = f.getFolders();
    processFolder(subFolder);
  }
}


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
  
  // Deleted the document once the text has been stored.
  Drive.Files.remove(doc.getId())
  
  return {
    name:fileName,
    text:text
  }
}

function extractFields(text, fileName){
  var regexOrdenPago = /(?<=Orden de.+[\r\n])([0-9]{1,})/
  var regexFechaEmision = /(?<=FPD )([\d]{2}\/[\d]{2}\/[\d]{4})/
  var regexCodigoMedico = /(?<=Codigo de medico:.+?)([\d]{1,})/
  var regexBeneficiario = /(?<=Beneficiario:\s+)(.+\b)/
  var regexCancelaciones = /(?<!D e b i t o s.+[\r\n].+)([\d ]{4,})\s*([^0-9\/]+)\s*(\d{2},\d{2})?\s*(\d{1,2}\/\d{4})\s* (\d{1,}\.*\d{1,3},\d{2})/g
  var regexDebitos = /(?<!C a n c e l a c i o n.+[\r\n].+)([\d]{4,})\s*([^0-9]+)\s*(\d{1,2}\/\d{4})\s* (\d{1,}\.*\d{1,3},\d{2})/g
  //var ordenPago = text.match(regexOrdenPago)
  var ordenPago = regexOrdenPago.exec(text)
  //var fechaEmision = text.match(regexFechaEmision)
  var fechaEmision = regexFechaEmision.exec(text)
  //var codigoMedico = text.match(regexCodigoMedico)
  var codigoMedico = regexCodigoMedico.exec(text)
  //var beneficiario = text.match(regexBeneficiario)
  var beneficiario = regexBeneficiario.exec(text)
  var cancelacionesToObject = text.matchAll(regexCancelaciones)
  var cancelacionesToArray = Array.from(cancelacionesToObject)
  
  try {
    var cancelaciones = cancelacionesToArray.map(function(array){
      return [fileName, ordenPago, fechaEmision, codigoMedico, beneficiario,
              array[1], array[2], array[3], array[4], array[5]]
    })
    }catch(e){
      var cancelaciones = [["Error"]] 
      }
  
  var debitosToObject = text.matchAll(regexDebitos)
  var debitosToArray = Array.from(debitosToObject)
  
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

function importToSpreadsheet(data, sheetName){
  const sheet = ss.getSheetByName(sheetName)
  const range = sheet.getRange(sheet.getLastRow()+1,1,data.length, data[0].length)
  range.setValues(data)
}