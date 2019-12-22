// Algunas variables generales que identifican hojas,rangos y celdas

var hoja = SpreadsheetApp.getActiveSpreadsheet();
var ui = SpreadsheetApp.getUi();

// Hoja 0.Alumnos
var hojaAlumnos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('0.Alumnos');
var filAlumnos = 2;
var colNombre = 1;
var colApellidos = 2;
var colEmail = 3;
var colFecNotificado = 4;
var colComentarios = 5;

// Hoja 1.Parámetros
var hojaParametros = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('1.Parámetros');
var nombreTarea = 'B2';
var filAspectoInicial = 3;
var filAspectoFinal = hojaParametros.getLastRow();
var colAspectos = 2;

// Hoja 4.Resultados
var hojaNotas = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('4.Resultados');
var filResInicial = 3;
var filResFinal = hojaNotas.getLastRow();
var colResInicial = 1;
var colResFinal = 2;
var colAspectosInicial = 3;
var puntuacionMedia = 'B1';

// Añade comando en el menú
function onOpen() {
  var menu = [{name:'Enviar calificaciones 📧', functionName:'enviarCalificaciones'}];
  hoja.addMenu('Evaluación', menu);
};

// Iniciar envío de calificaciones y abrir interfaz html
function enviarCalificaciones() {
  
   // Identificar última fila con *aspectos* a evaluar en la matriz de evaluación
   while (hojaParametros.getRange(filAspectoFinal,colAspectos).getValue() == '' && filAspectoFinal >= filAspectoInicial) {filAspectoFinal--;}

   // Si no los hay, terminar
   if (filAspectoInicial > filAspectoFinal) {
     ui.alert('¡No hay aspectos a evaluar en la matriz!');
   }
   else {
    
     // Identificar última fila con datos de *alumnos* en hoja de resultados
     while (hojaNotas.getRange(filResFinal,colResFinal).getValue() == '' && filResFinal >= filResInicial) {filResFinal--;}
    
       // Si no los hay, terminar
       if (filResInicial > filResFinal) {    
       ui.alert('¡No hay calificaciones en la pestaña de resultados!');
   }
     else {
      
       // Lanzamos panel de selección de destinatarios y comentarios
       var panel=HtmlService.createHtmlOutputFromFile('Panel')
        .setWidth(700)
        .setHeight(550);
       ui.showModalDialog(panel,'Enviar notas por email');
     }
  }
}

// Enviar a interfaz html datos de destinatarios
// filResFinal ya ha sido calculada en función enviarCalificaciones()
// Se devuelven nombres y notas de alumnos en hoja 4.Resultados
function obtenerDatosHoja(){
  return hojaNotas.getRange(filResInicial,colResInicial,filResFinal-filResInicial+1,colResFinal-colResInicial+1).getDisplayValues();
}

// Envío de emails con calificaciones (invocada desde panel modal)
// Recibe una lista de objetos {ID, comentario} ID = cardinal alumno
// los alumnos que disponen de calificación + comentarios
function enviarEmails(alumnosComentarios) {  
    
  // Aquí está la acción >> construir y enviar emails (si tenemos destinatarios)
 
  if (alumnosComentarios.length == 0) {
    SpreadsheetApp.getUi().alert('No se han seleccionado destinatarios, nada que hacer.');
  }
  else {
    
    hoja.toast('Procesando envío...');
    
    // Elementos comunes
    
    var comentarioGeneral = SpreadsheetApp.getUi().prompt('Introduce instrucciones o comentarios generales de cierre del correo electrónico o ACEPTAR:');
    var asunto = 'Calificación: ' + hojaParametros.getRange(nombreTarea).getValue();
    var alumnosNoEmail = '';
  
    // Variables globales reescritas, volver a identificar rangos de aspectos y calificaciones
    while (hojaParametros.getRange(filAspectoFinal,colAspectos).getValue() == '' && filAspectoFinal >= filAspectoInicial) {filAspectoFinal--;}   
    while (hojaNotas.getRange(filResFinal,colResFinal).getValue() == '' && filResFinal >= filResInicial) {filResFinal--;}
    var numAspectos = filAspectoFinal - filAspectoInicial + 1;
    
    // Vamos con cada alumno
    
    // Pasamos a la hoja 0 para que se aprecie el registro de fechas de envío y comentarios
    hoja.getSheetByName('0.Alumnos').activate();
    
    for (i in alumnosComentarios) {
    
      // Mensaje: Encabezado
      mensaje  = 'Hola, ' + hojaNotas.getRange(filResInicial+alumnosComentarios[i].ID, colResInicial).getValue() + ':\n\n';
      mensaje += 'Esta es tu puntuación en la actividad (todos los aspectos sobre 10):\n\n>> ' + hojaParametros.getRange(nombreTarea).getValue() + ' <<\n\n';
   
      // Mensaje: Evaluación de cada aspecto de la matriz de comprobación
      mensaje += 'Aspectos valorados ' + '[' + numAspectos + ']:\n\n';
      for (j=0; j<numAspectos; j++) {
        mensaje += '[' + hojaParametros.getRange(filAspectoInicial + j, colAspectos-1).getValue() + '] ';
        mensaje += hojaParametros.getRange(filAspectoInicial + j, colAspectos).getValue() + '\n';
        mensaje += '>> Puntuación: ' + hojaNotas.getRange(filResInicial + +alumnosComentarios[i].ID, colAspectosInicial+j).getDisplayValue() + ' <<\n\n';
      }
      
      // Mensaje: Nota final, media de la clase y comentarios específico y general (si existen)
      mensaje += '>> PUNTUACIÓN FINAL: ' + hojaNotas.getRange(filResInicial + alumnosComentarios[i].ID, colResFinal).getDisplayValue() + ' <<';
      mensaje += ' (media del grupo: ' + hojaNotas.getRange(puntuacionMedia).getDisplayValue() + ')';
      if (alumnosComentarios[i].comentario != "") { mensaje += '\n\n' + alumnosComentarios[i].comentario; }
      if (comentarioGeneral != "") { mensaje += '\n\n' + comentarioGeneral.getResponseText(); }
          
      // Email del destinatario
      emailAlumno = hojaAlumnos.getRange(filAlumnos+alumnosComentarios[i].ID, colEmail).getValue();
      
      // Si no existe en la tabla de alumnos nos anotamos el alumno afectado
      if (emailAlumno == "") {
        alumnosNoEmail +='❌ ' + hojaAlumnos.getRange(filAlumnos+alumnosComentarios[i].ID, colNombre).getValue() + ' ' +
                       hojaAlumnos.getRange(filAlumnos+alumnosComentarios[i].ID, colApellidos).getValue() + '\n';
      }
      else {
        // Por fin, enviar email
        try {
          MailApp.sendEmail(emailAlumno,asunto,mensaje);
        }      
        catch(e) {
          SpreadsheetApp.getUi().alert('¡Error!','❌ Se ha producido el error:\n\n'+e); }
          // Actualizar comentarios específicos y fecha de envío en hoja de alumnos
          hojaAlumnos.getRange(filAlumnos+alumnosComentarios[i].ID, colComentarios).setValue(alumnosComentarios[i].comentario);
          hojaAlumnos.getRange(filAlumnos+alumnosComentarios[i].ID, colFecNotificado).setValue(new Date()).setNumberFormat('dd/mm/yy HH:mm');
        }
      }
    }
    if (alumnosNoEmail.length > 0) { mensaje = 'Se han omitido alumnos (email no disponible):\n' + alumnosNoEmail; }
    else { mensaje = '';}
    ui.alert('Proceso terminado.\n\nLas notificaciones deberían aparecer en la carpeta\nde elementos enviados de tu buzón de Gmail.\n\n' + mensaje);
}
