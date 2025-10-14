const IS_TEST = true;

const ALLOWED_EMAILS = [
  'auxiliar1.oasic@unicatolicadelsur.edu.co',
  'auxiliar2.oasic@unicatolicadelsur.edu.co',
  'coord.tic@unicatolicadelsur.edu.co',
  'asistente.oasic@unicatolicadelsur.edu.co',
  'sena.tic@unicatolicadelsur.edu.co',
  ];

function doGet(e) {
  const email = IS_TEST
    ? 'auxiliar1.oasic@unicatolicadelsur.edu.co'
    : Session.getActiveUser().getEmail();

  if(!ALLOWED_EMAILS.includes(email)) {
    return HtmlService.createHtmlOutput(`<h2>Access denied</h2><p>Your email (${email}) is not authorized to access this form.</p>`);
  }

  return HtmlService
    .createHtmlOutputFromFile('index')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .append(`<script> var prefill = ${JSON.stringify(e.parameter)};</script>`);
}

function submitForm(data) {
  const formId = data.formId;
  delete data.formId;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(formId);

  if(!sheet) {
    throw new Error(`La hoja con nombre '${formId}' no existe.`);
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const email = IS_TEST
    ? 'auxiliar1.oasic@unicatolicadelsur.edu.co'
    : Session.getActiveUser().getEmail();

  const row = [new Date(), email];
  switch (formId) {
    
    case 'Inventario':
      inventarioForm(sheet, data, headers, row);
      break;

    case 'Préstamo':
      prestamoForm(sheet, data, headers, row);
      break;

    case 'Devolución':
      devolucionForm(sheet, data, headers, row);
      break;

    case 'Soporte':
      soporteForm(sheet, data, headers, row);
      break;
      
    default:
      return;
  }

}

function inventarioForm(sheet, data, headers, row){

  const tipoDePlaca = data['Tipo de Placa (I)'];
  const placaInventario = data['Placa Inventario (I)'];
  const placaCompleta = `${tipoDePlaca}-${placaInventario}`;
  const placaCompletaIndex = headers.indexOf('Placa Completa (I)');
  const sheetData = sheet.getDataRange().getValues();

  let rowIndex = sheetData.findIndex((sheetRow, i) => i > 0 && sheetRow[placaCompletaIndex] === placaCompleta);

  data.URL = (IS_TEST
    ? 'https://script.google.com/a/macros/unicatolicadelsur.edu.co/s/AKfycbwO0AZLisAXEwLucGd0MvqsAgwRQicaMy87BlMnM_Wp/dev?'
    : 'https://script.google.com/a/macros/unicatolicadelsur.edu.co/s/AKfycbwO0AZLisAXEwLucGd0MvqsAgwRQicaMy87BlMnM_Wp?')
    + data.URL;

  data['Placa Completa (I)'] = placaCompleta;

  for(let i = 2; i < headers.length - 1; i++) {
    const key = headers[i];
    row.push(data[key] || "");
  }

  if(rowIndex !== -1) {
    sheet.getRange(rowIndex + 1, 1, 1, row.length).setValues([row]);
    rowIndex = rowIndex + 1;
  } else {
    sheet.appendRow(row);
    rowIndex = sheet.getLastRow();
  }

  // Handle 'Disponibilidad' sheet logic
  if(data['Estado (I)'] === 'Para préstamo') {
    const disponibilidadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Disponibilidad');
    
    if(disponibilidadSheet) {
      const disponibilidadData = disponibilidadSheet.getDataRange().getValues();
      const disponibilidadHeaders = disponibilidadData[0];
      const nombreColumnIndex = disponibilidadHeaders.indexOf('NOMBRE');
      
      // Check if placaCompleta exists in 'NOMBRE' column
      const exists = disponibilidadData.findIndex((dispRow, i) => 
        i > 0 && dispRow[nombreColumnIndex] === placaCompleta
      );
      
      // If it doesn't exist, create a new row
      if(exists === -1) {
        const tipoColumnIndex = disponibilidadHeaders.indexOf('TIPO');
        const disponibleColumnIndex = disponibilidadHeaders.indexOf('DISPONIBLE');
        const activoColumnIndex = disponibilidadHeaders.indexOf('ACTIVO');
        
        const newRow = new Array(disponibilidadHeaders.length).fill('');
        newRow[nombreColumnIndex] = placaCompleta;
        newRow[tipoColumnIndex] = data['Tipo de Recurso (I)'];
        newRow[disponibleColumnIndex] = true;
        newRow[activoColumnIndex] = true;
        
        disponibilidadSheet.appendRow(newRow);
      }  else {
        // If it exists, ensure ACTIVO is TRUE IT DIDN'T WORK
        const activoColumnIndex = disponibilidadHeaders.indexOf('ACTIVO');
        disponibilidadSheet.getRange(exists + 1, activoColumnIndex + 1).setValue(true);
      }
    }
  } else {
    // When Estado (I) is NOT 'Para préstamo'
    const disponibilidadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Disponibilidad');
    
    if(disponibilidadSheet) {
      const disponibilidadData = disponibilidadSheet.getDataRange().getValues();
      const disponibilidadHeaders = disponibilidadData[0];
      const nombreColumnIndex = disponibilidadHeaders.indexOf('NOMBRE');
      const activoColumnIndex = disponibilidadHeaders.indexOf('ACTIVO');
      
      // Find the row with placaCompleta in 'NOMBRE' column
      const rowIndexDisp = disponibilidadData.findIndex((dispRow, i) => 
        i > 0 && dispRow[nombreColumnIndex] === placaCompleta
      );
      
      // If it exists, set ACTIVO to FALSE
      if(rowIndexDisp !== -1) {
        disponibilidadSheet.getRange(rowIndexDisp + 1, activoColumnIndex + 1).setValue(false);
      }
    }
  }

  const urlYeshua = 'http://yeshua.unicatolicadelsur.edu.co:4200/qr.php?code=';

  const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=150x150&data=${encodeURIComponent(urlYeshua + placaCompleta)}`;

  const response = UrlFetchApp.fetch(qrUrl);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd_HHmmss");
  const blob = response.getBlob().setName(`QR-${placaCompleta}-${timestamp}.png`);

  const folder = DriveApp.getFolderById('19zlkq_wNZ8nKJ5bi5uUaDhgjnuQQg9y0');
  const file = folder.createFile(blob);

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const fileId = file.getId();
  const publicUrl = `https://drive.google.com/uc?id=${fileId}`;

  const numColumns = sheet.getLastColumn();
  sheet.getRange(rowIndex, numColumns).setFormula(`=IMAGE("${publicUrl}";4;150;150)`);
  sheet.setRowHeight(rowIndex, 150);
  sheet.setColumnWidth(numColumns, 150);

}

function soporteForm(sheet, data, headers, row){

  const tipoDePlaca = data['Tipo de Placa (S)'];
  const placaInventario = data['Placa Inventario (S)'];
  const placaCompleta = `${tipoDePlaca}-${placaInventario}`;
  data['Placa Completa (S)'] = placaCompleta;

  const inventarioSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventario');

  if(inventarioSheet) {
    const inventarioData = inventarioSheet.getDataRange().getValues();
    const inventarioHeaders = inventarioData[0];
    const placaCompletaIndex = inventarioHeaders.indexOf('Placa Completa (I)');

    const rowIndexPlacaCompleta = inventarioData.findIndex((invRow, i) => 
      i > 0 && invRow[placaCompletaIndex] === placaCompleta
    );

    const solicitante = rowIndexPlacaCompleta !== -1
      ? inventarioData[rowIndexPlacaCompleta][inventarioHeaders.indexOf('Responsable (I)')]
      : '';

    data['Solicitante (S)'] = solicitante === '' ? 'TIC' : solicitante;
  }

  for(let i = 2; i < headers.length; i++) {
    const key = headers[i];
    row.push(data[key] || "");
  }

  sheet.appendRow(row);

}

function prestamoForm(sheet, data, headers, row){
  const disponibilidadSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Disponibilidad');

  if(disponibilidadSheet) {
    cualControlHDMI = data['Cuál control HDMI? (D)'];
    cualPortatil = data['Cuál portátil? (D)'];
    cualOtro = data['Cuál otro? (D)'];

    const disponibilidadData = disponibilidadSheet.getDataRange().getValues();
    const disponibilidadHeaders = disponibilidadData[0];
    const nombreColumnIndex = disponibilidadHeaders.indexOf('NOMBRE');
    //TODO end this
  }

  for(let i = 2; i < headers.length; i++) {
    const key = headers[i];
    row.push(data[key] || "");
  }
  sheet.appendRow(row);
}

function devolucionForm(sheet, data, headers, row){

}

function getDeviceOptions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Disponibilidad");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const nombreIdx = headers.indexOf("NOMBRE");
  const tipoIdx = headers.indexOf("TIPO");
  const disponibleIdx = headers.indexOf("DISPONIBLE");
  const activoIdx = headers.indexOf("ACTIVO");

  const result = {
    hdmi: [],
    portatil: [],
    otro: [],
  };

  for(let i = 1; i < data.length; i++) {
    const row = data[i];
    const activo = row[activoIdx];

    if (activo !== true) {
      continue;
    }

    const nombre = row[nombreIdx];
    const tipo = row[tipoIdx];
    const disponible = row[disponibleIdx];

    const option = {
      value: nombre,
      disabled: disponible !== true
    };

    switch (tipo) {
      case "Control HDMI":
        result.hdmi.push(option);
        break;
      case "Portátil":
        result.portatil.push(option);
        break;
      default:
        result.otro.push(option);
        break;
    }
  }

  return result;

}

function getAllDevicesByTypeOfPlate() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventario");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  tipoDePlacaIdx = headers.indexOf("Tipo de Placa (I)");
  placaInventarioIdx = headers.indexOf("Placa Inventario (I)");
  estadoIdx = headers.indexOf("Estado (I)");

  const result = {
    FUCS: [],
    TMS: [],
  };

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tipoDePlaca = row[tipoDePlacaIdx];
    const placaInventario = row[placaInventarioIdx];
    const estado = row[estadoIdx];

    if (estado === "Dado de baja") {
      continue; // Skip items that are "Dado de baja",
    }
    switch (tipoDePlaca) {
      case "FUCS":
        result.FUCS.push({ value: placaInventario });
        break;
      case "TMS":
        result.TMS.push({ value: placaInventario });
        break;
    }
  }
  return result;
}

function getPlaceOptions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Lugares");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const sedeIdx = headers.indexOf("SEDE");
  const nombreIdx = headers.indexOf("NOMBRE");

  const result = {
    Torobajo: [],
    Centro: [],
  };
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const nombre = row[nombreIdx];
    const sede = row[sedeIdx];

    switch(sede){
      case "Torobajo":
        result.Torobajo.push({ value: nombre });
        break;
      case "Centro":
        result.Centro.push({ value: nombre });
        break;
    }
  }
  return result;
}

function initializePlainTextFormat() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Soporte");
  const range = sheet.getDataRange();
  range.setNumberFormat("@STRING@");
}

function downloadRedirectorPHP() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Inventario');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const keyColIndex = headers.indexOf('Placa Inventario (I)');
  const valueColIndex = headers.indexOf('URL');

  if (keyColIndex === -1 || valueColIndex === -1) {
    SpreadsheetApp.getUi().alert("No se encontraron las columnas 'Placa Inventario (I)' y 'URL'");
    return;
  }

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();

  let phpContent = `<?php

$map = [\n`;

  data.forEach(row => {
    const key = row[keyColIndex]?.toString().trim();
    const value = row[valueColIndex]?.toString().trim();
    if (key && value) {
      phpContent += `    '${key}' => '${value}',\n`;
    }
  });

  phpContent += `];

$code = $_GET['code'] ?? '';
$url = $map[$code] ?? 'https://my.link.com/error.html';

?>
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Redireccionando...</title>
  <script>
    setTimeout(() => {
      window.location.href = <?= json_encode($url) ?>;
    }, 100);
  </script>
</head>
<body>
  <p>Redireccionando...</p>
</body>
</html>
`;

  const blob = Utilities.newBlob(phpContent, 'application/x-httpd-php', 'qr.php');

  // Find or create folder "QR Files"
  const folderName = 'QR Files';
  let folder;
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }

  // Remove existing "qr.php" files in the folder
  const files = folder.getFilesByName('qr.php');
  while (files.hasNext()) {
    files.next().setTrashed(true);
  }

  // Create the new file in the folder
  const file = folder.createFile(blob);

  SpreadsheetApp.getUi().alert(`Archivo "qr.php" guardado exitosamente en tu carpeta "${folderName}" en Google Drive.`);
}



