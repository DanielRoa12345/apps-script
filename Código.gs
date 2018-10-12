function myFunction(){
  var doc = DocumentApp.create("documento");
  doc.getBody().appendParagraph("Bienvenido al mundo de los script");
   doc.getBody().appendParagraph("Bienvenido Google App Script ");
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('principal');

 }

function getIndicadores() {
  var html = HtmlService.createHtmlOutputFromFile('Indicadores').getContent();
  return html
}
function getGraf() {
  var html = HtmlService.createHtmlOutputFromFile('Graf').getContent();
  return html
}

function getConsultarIndicadores() {
  return HtmlService.createHtmlOutputFromFile('ListaIndicadores')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
}




function getEmpresas(){
  var html = HtmlService.createHtmlOutputFromFile('empresas').getContent();
  return html
}

function getEmpresas_Tecnologias(){
  var html = HtmlService.createHtmlOutputFromFile('Empresas_Tecnologias').getContent();
  return html
}

// Contiene el Xhtml de boton registrar indicadores que abre dindicadores empresas
function getIndicadores_Empresas(){
  var html = HtmlService.createHtmlOutputFromFile('Indicadores_Empresas').getContent();
  
  return html
}

function getLecciones(){
  var html = HtmlService.createHtmlOutputFromFile('LeccionesAprendidas').getContent();
  return html
}

function getGraficos(){
  var html = HtmlService.createHtmlOutputFromFile('Graficos').getContent();
  
  return html
}

function getGrafico(){
   var mysheet = SpreadsheetApp.openById('18EARPeky_C4Er3KEb3L0-sbP7S7ClWhtw7exVYPQz2c');
  var sheet = mysheet.getSheets()[0];
  var data = sheet.getRange('C1:E9').getValues();
  
  var dataTable = Charts.newDataTable();
  dataTable.addColumn(Charts.ColumnType.STRING, data[0][0]);
  for(var i=1; i<data[0].length;i++){
    dataTable.addColumn(Charts.ColumnType.NUMBER, data[0][i]);
  }
  
  for(var j=1; j<data.length; j++){
    dataTable.addRow(data[j]);
  }
  
  var chart=Charts.newPieChart()
  .setDataTable(dataTable)
  .setTitle("Sales")
  .build();
  
  var app= UiApp.createApplication().setTitle("My ");
  app.add(chart)
  return app;

}
    
      
      
function Init()
{
  var spreadsheet  = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1nZk1Pv5EZPLdd13tZv9JKjWNZsEW6bRxaqVM_NDEhFA/edit#gid=0");
  var sheet        = spreadsheet.getActiveSheet();
  var rows         = sheet.getDataRange();
  var numRows      = rows.getNumRows();
  var values       = rows.getValues();
  var string = "";
  var string1 = "";
  
  for(var i = 0 ; i < numRows ; ++i)
  {
    var row = values[i];
  
    string += "<p>" + row[1] +"</p>";
  
    string1 += "<p>" + row[2] +"</p>";
  }
 
 

  return "<tr>"+"<td>"+string+"</td>"+"<td>"+string1+"</td>"+"</tr>";
}

// Metodo de consultar empresas en el xhtml de empresas del boton consultar empresas.
function getConsultaEmpresas(){
  
  // var html = HtmlService.createHtmlOutputFromFile('TablaEmpresas').getContent();
  //return html
  
   

  
  var spreadsheet  = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1nZk1Pv5EZPLdd13tZv9JKjWNZsEW6bRxaqVM_NDEhFA/edit?usp=drive_web&ouid=101628587313676392748");  //Esta ruta será la que tengas a tu fichero que uses como BBDD
  var sheet        = spreadsheet.getActiveSheet();
  var rows         = sheet.getDataRange();
  var numRows      = rows.getNumRows();
  var values       = rows.getValues();
  var string = "";
  var string1="";
  var tableOpen = "<table class='table table-striped' style='margin-top: 50px !important; width:50%; border-radius: 5px; margin: 0px auto; float: none;'>";
  var tableClose = "</table>";
  var headerTable = "<tr><th>Nit</th><th>Nombres</th><th>Gerente</th><th>Ciudad</th><th>Direccion</th><th>Telefono</th><th>Celular</th><th>Email</th><th>Tecnología</th></tr>";

  
  
 for(var i = 1 ; i < numRows ; ++i)
  {
  var row = values[i];
  
    string+="<tr>";
    string+="<td>" + row[1] + "</td>";
    string+="<td>" + row[2] + "</td>";
    string+="<td>" + row[3] + "</td>";
    string+="<td>" + row[4] + "</td>";
    string+="<td>" + row[5] + "</td>";
    string+="<td>" + row[6] + "</td>";
    string+="<td>" + row[7] + "</td>";
    string+="<td>" + row[8] + "</td>";
    string+="<td>" + row[9] + "</td>";
    string+= "</tr>";
  
   
}
 

  
  return tableOpen + headerTable + string + tableClose ;

  }

function getTablaParametrica(sheet)
{
  return SpreadsheetApp.openById(ID_HOJA).getSheetByName(sheet).getDataRange().getValues();
 // return tabla;
}


// Metodo de consulta de indicadores de el xhtml de indicadores.
function getConsultaIndicadores(){
  
  // var html = HtmlService.createHtmlOutputFromFile('TablaEmpresas').getContent();
  //return html
  //google.charts.load('current', {'packages':['corechart']});
  //google.charts.setOnLoadCallback(drawChart);


  
  var spreadsheet  = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1ZIcyLjGc2A16E2Dz-18xit4Jaxolp5yaJ-KjPXUjB6I/edit#gid=0");  //Esta ruta será la que tengas a tu fichero que uses como BBDD
  var sheet        = spreadsheet.getActiveSheet();
  var rows         = sheet.getDataRange();
  var numRows      = rows.getNumRows();
  var values       = rows.getValues();
  var string = "";
  var string1="";

  var tableOpen = "<table class='table table-striped' style='margin-top: 50px !important; width:50%; border-radius: 5px; margin: 0px auto; float: none;'>";
  var tableClose = "</table>";
  var headerTable = "<tr><th>Código</th><th>Nombre</th></tr>";
  
  for(var i = 1 ; i < numRows ; ++i)
  {
  var row = values[i];
  
 //string += "<p>" + row[1]  +row[2]+"</p>";
    string+="<tr>";
    string+="<td>" + row[1] + "</td>";
    string+="<td>" + row[2] + "</td>";
    string+= "</tr>";
  
   
}
  
   
  //drawChart();
  
  var range = sheet.getRange("A1:B23");

  var dataSourceUrl = 'https://docs.google.com/spreadsheet/tq?range=A1%3AG5&key=0Aq4s9w_HxMs7dHpfX05JdmVSb1FpT21sbXd4NVE3UEE&gid=2&headers=-1';
  
 var dataSourceUrl2="https://docs.google.com/spreadsheets/d/1ZIcyLjGc2A16E2Dz-18xit4Jaxolp5yaJ-KjPXUjB6I/edit#gid=0";

  
  var data = Charts.newDataTable()
       .addColumn(Charts.ColumnType.STRING, "Month")
       .addColumn(Charts.ColumnType.NUMBER, "In Store")
       .addColumn(Charts.ColumnType.NUMBER, "Online")
       .addRow(["January", 10, 1])
       .addRow(["February", 12, 1])
       .addRow(["March", 20, 2])
       .addRow(["April", 50, 3])
       .addRow(["May", 30, 4])
       .build();

  
    var chartBuilder = Charts.newLineChart()
       .setTitle('Yearly Rainfall')
       .setXAxisTitle('Month')
       .setYAxisTitle('Rainfall (in)')
       .setDimensions(600, 500)
       .setCurveStyle(Charts.CurveStyle.SMOOTH)
       .setPointStyle(Charts.PointStyle.MEDIUM)
       .setDataSourceUrl(dataSourceUrl);

   var chart = chartBuilder.build();
  
   var htmlOutput = HtmlService.createHtmlOutput().setTitle('My Chart');
    var imageData = Utilities.base64Encode(chart.getAs('image/png').getBytes());
    var imageUrl = "data:image/png;base64," + encodeURI(imageData);
    htmlOutput.append("Render chart server side: <br/>");
    htmlOutput.append("<img border=\"1\" src=\"" + imageUrl + "\">");
    //return htmlOutput;



  Logger.log(htmlOutput.getContent());

  //return tableOpen + headerTable + string + tableClose + "<br/>" + htmlOutput.getContent() + "'/>";
  return tableOpen + headerTable + string + tableClose ;

}


// Se coloca la funcion para el grafico . 
function drawChart() {
        var data = google.visualization.arrayToDataTable([
          ['Year', 'Sales', 'Expenses'],
          ['2013',  20000,      400],
          ['2014',  1170,      460],
          ['2015',  660,       1120],
          ['2016',  1030,      540]
        ]);

        var options = {
          title: 'Company Performance',
          hAxis: {title: 'Year',  titleTextStyle: {color: '#333'}},
          vAxis: {minValue: 0}
        };

        var chart = new google.visualization.AreaChart(document.getElementById('chart_div'));
        chart.draw(data, options);
      }





// Metodo de boton consulta tecnologias Empresas tecnologias 
function getConsultaTecnologias(){
  
  // var html = HtmlService.createHtmlOutputFromFile('TablaEmpresas').getContent();
  //return html
  
  var spreadsheet  = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1ARJk2oZ392YNV5V62XSfwJAwbmRJGwjjYc1beKTLP_I/edit?usp=drive_web&ouid=101628587313676392748");  //Esta ruta será la que tengas a tu fichero que uses como BBDD
  var sheet        = spreadsheet.getActiveSheet();
  var rows         = sheet.getDataRange();
  var numRows      = rows.getNumRows();
  var values       = rows.getValues();
  var string = "";
  var string1="";
  
  var tableOpen = "<table class='table table-striped' style='margin-top: 50px !important; width:50%; border-radius: 5px; margin: 0px auto; float: none;'>";
  var tableClose = "</table>";
  
  var headerTable = "<tr><th>Codigo_T</th><th>TipoTecnologia</th><th>Nombre_T</th><th>Versión_T</th><th>Costo_T</th><th>SoftwareL</th></tr>";


  for(var i = 1 ; i < numRows ; ++i)
  {
  var row = values[i];
  
 //string += "<p>" + row[1]  +row[2] +row[3] +row[4] +row[5] +row[6] +"</p>";
  
       string+="<tr>";
    string+="<td>" + row[1] + "</td>";
    string+="<td>" + row[2] + "</td>";
    string+="<td>" + row[3] + "</td>";
    string+="<td>" + row[4] + "</td>";
    string+="<td>" + row[5] + "</td>";
    string+="<td>" + row[6] + "</td>";
    string+= "</tr>";
  
   
}
  //return string;
  return tableOpen + headerTable + string + tableClose;

  
  
  
}




//Metodo consultar empresas tecnologias del xhtml Tecnologias 

function getConsultaTec_Empr(){
  
  // var html = HtmlService.createHtmlOutputFromFile('TablaEmpresas').getContent();
  //return html
  
  var spreadsheet  = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1W6ZWZi0QFmI2-hJiW6e_hOmJEICOsf4ZQ_atl06KbWM/edit");  //Esta ruta será la que tengas a tu fichero que uses como BBDD
  var sheet        = spreadsheet.getActiveSheet();
  var rows         = sheet.getDataRange();
  var numRows      = rows.getNumRows();
  var values       = rows.getValues();
  var string = "";
  var string1="";

  var tableOpen = "<table class='table table-striped' style='margin-top: 50px !important; width:50%; border-radius: 5px; margin: 0px auto; float: none;'>";
  var tableClose = "</table>";
  var headerTable = "<tr><th>Empresa</th><th>Tecnologia</th></tr>";
  
  for(var i = 1 ; i < numRows ; ++i)
  {
  var row = values[i];
  
 //string += "<p>" + row[1]  +row[2]+"</p>";
  string+="<tr>";
  string+="<td>" + row[1] + "</td>";
  string+="<td>" + row[2] + "</td>";
  string+= "</tr>";
  
    
   
}
  
 return tableOpen + headerTable + string + tableClose;


}



// Metodo que ejecuta el xhtml de Lecciones Aprendidas
function getConsultaLecciones(){
  
  // var html = HtmlService.createHtmlOutputFromFile('TablaEmpresas').getContent();
  //return html
  
  var spreadsheet  = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1j-w57veqIzZf0_n1gBC-GenHd3omHre25M0i6YVHR3A/edit#gid=0");  //Esta ruta será la que tengas a tu fichero que uses como BBDD
  var sheet        = spreadsheet.getActiveSheet();
  var rows         = sheet.getDataRange();
  var numRows      = rows.getNumRows();
  var values       = rows.getValues();
  var string = "";
  var string1="";
  var tableOpen = "<table class='table table-striped' style='margin-top: 50px !important; width:50%; border-radius: 5px; margin: 0px auto; float: none;'>";
  var tableClose = "</table>";
  var headerTable = "<tr><th>Empresa</th><th>LECCION</th></tr>";
  
  for(var i = 1 ; i < numRows ; ++i)
  {
  var row = values[i];
  
      string+="<tr>";
  string+="<td>" + row[1] + "</td>";
  string+="<td>" + row[2] + "</td>";
  string+= "</tr>";
  
   
}
  
 return tableOpen + headerTable + string + tableClose;


}

//FRAN

function queryHoja(tabla, campos, filtros)
{
  var hoja = SpreadsheetApp.openById(ID_HOJA).getSheetByName(tabla);
  var cols = campos.split(":");
  var criterios = filtros.split(":");
  var flag = true;
  
  var nFilas = hoja.getLastRow();
  var DATOS = new Array();
  var nReg = 0;

  for(var fila = 1; fila<=nFilas; fila++){
    for(var i=0; i<criterios.length; i++){
      var criterio = criterios[i].split("=");
      var valor = hoja.getRange(fila, criterio[0]).getValue();
      if(valor != criterio[1]){
        flag = false;
        break;
      }
    }
    if(flag == true){
      var Registro = new Array();
      for(var i=0; i<cols.length; i++){
        Registro[i] = hoja.getRange(fila, cols[i]).getValue();
      }
      DATOS[nReg++] = Registro;
    }else{
      flag = true;
    }
  }
  return DATOS;    
}



//////////////////////////////////////77
function getTecnologias(){
  var html = HtmlService.createHtmlOutputFromFile('Tecnologias').getContent();
  return html
}

function getPagina2() {
  var html = HtmlService.createHtmlOutputFromFile('Pagina2').getContent();
  return html
}

function createAndSendDocument() {
    //Crear un nuevo Documento de Nombre Hola Mundo de AppScript
  var doc = DocumentApp.create('Hola Mundo de AppScript');
  //Obtenemos el Body del Documento y agregamos un Parrafo
  doc.getBody().appendParagraph('Este Documento fue creado a Partir de AppScript');
  //URL del Documento Generado
  var url = doc.getUrl();
  //Obtenemos nuestro Correo Electronico
  var email = Session.getActiveUser().getEmail();
  //El asunto es el nombre del Documento
  var subject = doc.getName();
  //El cuerpo del correo max 20kb indica la URL de nuestro documento
  var body = 'Link con tu Documento: ' + url;
  //Enviamos el correo (:
  GmailApp.sendEmail(email, subject, body);
}


function getListOptions() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Base_Datos_empresas');
  var lastRow = sheet.getLastRow();  
  var myRange = sheet.getRange("B2"+lastRow); 
  var id = myRange.getValues(); 
  Logger.log(id);
  return(id);

}

// Metodo de Resultado de indicadores  en el xhtml de empresas del boton Registrar indicadores empresas.
function getResultadoEmpresas(){
  
  // var html = HtmlService.createHtmlOutputFromFile('TablaEmpresas').getContent();
  //return html
  
  var spreadsheet  = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/18EARPeky_C4Er3KEb3L0-sbP7S7ClWhtw7exVYPQz2c/edit#gid=0");  //Esta ruta será la que tengas a tu fichero que uses como BBDD
  var sheet        = spreadsheet.getActiveSheet();
  var rows         = sheet.getDataRange();
  var numRows      = rows.getNumRows();
  var values       = rows.getValues();
  var string = "";
  var string1="";
  var tableOpen = "<table class='table table-striped' style='margin-top: 50px !important; width:50%; border-radius: 5px; margin: 0px auto; float: none;'>";
  var tableClose = "</table>";
  var headerTable = "<tr><th>Empresa</th><th>Indicador</th><th>Valor_Inicial</th><th>Valor_Actual</th><th>Notas</th></tr>";

  
  
 for(var i = 1 ; i < numRows ; ++i)
  {
  var row = values[i];
  
    string+="<tr>";
    string+="<td>" + row[1] + "</td>";
    string+="<td>" + row[2] + "</td>";
    string+="<td>" + row[3] + "</td>";
    string+="<td>" + row[4] + "</td>";
    string+="<td>" + row[5] + "</td>";
    string+= "</tr>";
  
   
}
  
  return tableOpen + headerTable + string + tableClose;

  }


function grafica(){
var ss = SpreadsheetApp.openById("18EARPeky_C4Er3KEb3L0-sbP7S7ClWhtw7exVYPQz2c");
  var data = ss.getDataRange();
  
  var ageFilter = Charts.newNumberRangeFilter().setFilterColumnIndex(4).build();
  var transporFiler = Charts.newCategoryFilter().setFilterColumnIndex(3);
  var nameFiler = Charts.newStringFilter().setFilterColumnIndex(1).build();
  
  var tableChart = Charts.newTableChart().setDataViewDefinition(Charts.newDataViewDefinition().setColumns([1,2,3,4])).build();
  
  var pieChart = Charts.newPieChart().setDataViewDefinition(Charts.newDataViewDefinition().setColumns([1,2,3,4])).build();
  
  var pieChart = Charts.newPieChart().setDataViewDefinition(Charts.newDataViewDefinition().setColumns([1,2,3,4])).build();
  
  var dashboard = Charts.newDashboardPanel().setDataTable(data).bind([ageFilter , transporFiler , nameFiler],
  [tableChart , pieChart]).build();
  
  var app = UiApp.createApplication();
  var filterPanel = app.createVerticalPanel();
  var chartPanel = app.createHorizontalPanel();
  filterPanel.add(ageFilter).add(transporFiler).add(nameFiler).setSpacing(10);
  
  chartPanel.add(app.createHorizontalPanel().add(filterPanel).add(chartPanel));
  
  app.add(dashboard);
  
  return app ;
  
  
}

// Metodo de consulta tecnologias List para asociar a empresa al momento de crearla 
function getConsultaTecnologiasList(){
  
  // var html = HtmlService.createHtmlOutputFromFile('TablaEmpresas').getContent();
  //return html
  
  var spreadsheet  = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1ARJk2oZ392YNV5V62XSfwJAwbmRJGwjjYc1beKTLP_I/edit?usp=drive_web&ouid=101628587313676392748");  //Esta ruta será la que tengas a tu fichero que uses como BBDD
  var sheet        = spreadsheet.getActiveSheet();
  var rows         = sheet.getDataRange();
  var numRows      = rows.getNumRows();
  var values       = rows.getValues();
  var string = "";
  var string1="";
  
  var tableOpen = "<table class='table table-striped' style='margin-top: 50px !important; width:50%; border-radius: 5px; margin: 0px auto; float: none;'>";
  var tableClose = "</table>";
  
  var headerTable = "<tr><th>Codigo_T</th><th>TipoTecnologia</th><th>Nombre_T</th><th>Versión_T</th><th>Costo_T</th><th>SoftwareL</th></tr>";


  for(var i = 1 ; i < numRows ; ++i)
  {
  var row = values[i];
  
 //string += "<p>" + row[1]  +row[2] +row[3] +row[4] +row[5] +row[6] +"</p>";
  
       string+="<tr>";
    string+="<td>" + row[2] + "</td>";
    string+= "</tr>";
  
   
}
  //return string;
  return tableOpen + headerTable + string + tableClose;

  
  
  
}


function getConsultaEmpresasList(){
  
  // var html = HtmlService.createHtmlOutputFromFile('TablaEmpresas').getContent();
  //return html
  
   

  
  var spreadsheet  = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1nZk1Pv5EZPLdd13tZv9JKjWNZsEW6bRxaqVM_NDEhFA/edit?usp=drive_web&ouid=101628587313676392748");  //Esta ruta será la que tengas a tu fichero que uses como BBDD
  var sheet        = spreadsheet.getActiveSheet();
  var rows         = sheet.getDataRange();
  var numRows      = rows.getNumRows();
  var values       = rows.getValues();
  var string = "";
  var string1="";
  var tableOpen = "<table class='table table-striped' style='margin-top: 50px !important; width:50%; border-radius: 5px; margin: 0px auto; float: none;'>";
  var tableClose = "</table>";
  var headerTable = "<tr><th>Nit</th><th>Nombres</th><th>Gerente</th><th>Ciudad</th><th>Direccion</th><th>Telefono</th><th>Celular</th><th>Email</th><th>Tecnología</th></tr>";

  
  
 for(var i = 1 ; i < numRows ; ++i)
  {
  var row = values[i];
  
    string+="<tr>";
    string+="<td>" + row[2] + "</td>";
    string+= "</tr>";
  
   
}
 

  
  return tableOpen + headerTable + string + tableClose ;

  }







