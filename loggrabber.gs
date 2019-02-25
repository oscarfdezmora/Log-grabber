function lastBlankRow(link,tab){
// Dado un link y una pestaña, devuelve la primera línea en completamente en blanco de un documento;
  var values = SpreadsheetApp.openByUrl(link).getSheetByName(tab).getDataRange().getValues();
  var row = 0;
  for (var row=0; row<values.length; row++) {
    if (!values[row].join("")) break;
  }
    return row+1;
}

function lastBlankLOG(link,tab){
// Dado un link y una pestaña, devuelve la primera línea en blanco para los Logs (columnas C y D)
  var values = SpreadsheetApp.openByUrl(link).getSheetByName(tab).getRange(13, 4, 1000, 44).getValues();
  var row = 0;
  for (var row=0; row<values.length; row++) {
    if ((values[row][3]=="")&&(values[row][2]=="")) break;
  }
    return row; 
}

function cogerDato(origenUrl,origenTab,destinoUrl,destinoTab,pais,area){
// Toma datos de un link y pestaña (Origen), y los inserta en otro link y pestaña (Destino)
// En destino añade las columnas con los valores de pais y área a modo de identificación
  var origenFichero = SpreadsheetApp.openByUrl(origenUrl);
  var origenPestana = origenFichero.getSheetByName(origenTab);
  
  //En caso de que el documento no contenga datos, lo salta
  if (lastBlankLOG(origenUrl,origenTab)==0) {return}
  
  //Se posiciona para coger los datos en 13:4
  var datos = origenPestana.getRange(13,4,lastBlankLOG(origenUrl,origenTab),60).getValues();
  
  var destinoFichero = SpreadsheetApp.openByUrl(destinoUrl);
  var destinoPestana = destinoFichero.getSheetByName(destinoTab);
  var fila = lastBlankRow(destinoUrl,destinoTab); //comprueba cuál es la primera línea blanca donde escribir
  destinoPestana.getRange(fila,3,lastBlankLOG(origenUrl,origenTab),60).setValues(datos);
  destinoPestana.getRange(fila,1,lastBlankLOG(origenUrl,origenTab)).setValue(pais);
  destinoPestana.getRange(fila,2,lastBlankLOG(origenUrl,origenTab)).setValue(area);
  
}

function backup(url,tab){
  // Realiza backup de la pestaña dada, renombrando con la fecha
  var newName = tab + Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd'T'HH:mm:ss'Z'");
  var ss = SpreadsheetApp.openByUrl(url);
  ss.getSheetByName(tab).copyTo(ss).setName(newName);
}

function createTemp(url,tab){
  // Genera un fichero temporal para el tratamiento de los datos
  backup(url,tab);
  var tempTab = "temp."+tab;
  // Toma las cabeceras y las pega de nuevo
  var datos = SpreadsheetApp.openByUrl(url).getSheetByName(tab).getRange("a1:az1").getValues();
  SpreadsheetApp.openByUrl(url).insertSheet(tempTab).getRange("a1:az1").setValues(datos);

}

function logGrabber(){
  // Dados un listado de URL de logs, genera dos pestañas con la info de Personas y UC, versionando la información anterior. 
  // Tras generarlo, manda la información a un documento donde se trata el consolidado. 
  
  // Pestaña con los datos de origen
  var logsUrl = 'https://docs.google.com/spreadsheets/d/1F57Ubru6sEwlj59keZ5RB_DaGVeevvn0dvQ016Lbu9E/edit#gid=827860839';
  var logsTab = "URLs";
  
  // Pestañas donde se va a generar la información
  var destinoUrl = 'https://docs.google.com/spreadsheets/d/1ZVJWCJcdebmlUMq-sPTlv_ns8D6wFiT8jZwKjFiVbM0/edit#gid=0';
  var destinoTabLog = "Log - UC";
  var destinoTabStaff = "Log - Staff";
  
  // Creación de temporales
  createTemp(destinoUrl,destinoTabLog);
  createTemp(destinoUrl,destinoTabStaff);
  
  var datosLogs = SpreadsheetApp.openByUrl(logsUrl).getSheetByName(logsTab).getRange("a:z").getValues();
  // Valores tomados como referencia de la hoja de URLs; se los recorre tomando la información de cada BU
  for (var row=0; row<datosLogs.length; row++) {
    var pais = datosLogs[row][0];
    var area = datosLogs[row][3];
    var link = datosLogs[row][4];
    var tabUC = "Líneas de Trabajo";
    var tabStaff = "Staff - DS Dsp + Área";
    
    if (pais=="USA") {tabUC = "Lines of Work";tabStaff = "Staff - DS Dsp + Area"}; // Excepciones por USA
    
    if (pais=="") {break}  // Excepción si la hoja está vacía
    
    cogerDato(link,tabUC,destinoUrl,"temp." + destinoTabLog,pais,area);
    cogerDato(link,tabStaff,destinoUrl,"temp." + destinoTabStaff,pais,area);
    
  }
  
  SpreadsheetApp.openByUrl(destinoUrl).deleteSheet(SpreadsheetApp.openByUrl(destinoUrl).getSheetByName(destinoTabLog));
  SpreadsheetApp.openByUrl(destinoUrl).getSheetByName("temp." + destinoTabLog).setName(destinoTabLog);
  
  SpreadsheetApp.openByUrl(destinoUrl).deleteSheet(SpreadsheetApp.openByUrl(destinoUrl).getSheetByName(destinoTabStaff));
  SpreadsheetApp.openByUrl(destinoUrl).getSheetByName("temp." + destinoTabStaff).setName(destinoTabStaff);
  
  // Llevar dato al consolidado
  var consolidateUrl = 'https://docs.google.com/spreadsheets/d/1F57Ubru6sEwlj59keZ5RB_DaGVeevvn0dvQ016Lbu9E/edit#gid=1092206245';
  var consolidatedSS = SpreadsheetApp.openByUrl(consolidateUrl);
  
  consolidatedSS.deleteSheet(consolidatedSS.getSheetByName(destinoTabLog));
  SpreadsheetApp.openByUrl(destinoUrl).getSheetByName(destinoTabLog).copyTo(consolidatedSS).setName(destinoTabLog);
  
  consolidatedSS.deleteSheet(consolidatedSS.getSheetByName(destinoTabStaff));
  SpreadsheetApp.openByUrl(destinoUrl).getSheetByName(destinoTabStaff).copyTo(consolidatedSS).setName(destinoTabStaff);
    
}