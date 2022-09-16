function SendEmail(){

  var u1 = SpreadsheetApp.getUi();
  u1.alert("Script de vencimientos activado.");

  // Obtener dirección de correo
  var emailRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Documento").getRange("M2");
  var emailAddress = emailRange.getValue();

  var fechaLimiteRango = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Documento");
  fechaLimiteRango = fechaLimiteRango.getRange("H2:H36").getValues();
  var diasPrevios = 30;

  if (fechaLimiteRango != null) {
    for (var i = 1; i < fechaLimiteRango.length; i++) {
      var fecha = fechaLimiteRango[i];
      var fechaLimite = new Date(fecha);
      var hoy = new Date();
      var diasRestantes = Math.round((fechaLimite - hoy) / (1000 * 60 * 60 * 24));
      if (diasRestantes <= diasPrevios) {
        //Logger.log("Falta menos de " + diasPrevios + " dias para la fecha limite");
        //Logger.log("Por Vencer");

        // Enviar alerta
        var message = 'Alerta, Servicios por vencer!'; // Second column
        var subject = 'Importante - Se notifica que hay servicios por vencer!';
        MailApp.sendEmail(emailAddress, subject, message);

      } else{
        Logger.log("Falta mucho");
      }
    }
  } else{
      Logger.log("No hay fechas establecidas");
  }


  // // Enviar alerta
  // var message = 'PruebaUno!'; // Second column
  // var subject = 'Alerta desde Google Sheets';
  // MailApp.sendEmail(emailAddress, subject, message);

}
