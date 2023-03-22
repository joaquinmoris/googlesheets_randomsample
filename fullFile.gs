/////////////////////////////////////////////////////
//
// FUNCIÓN PARA MENÚ Y FUNCIÓN PARA PROMPT
//
/////////////////////////////////////////////////////


function onOpen(e) {
  //Crea en el menú de la hoja un submenú desde el que lanzar la función de generación y presentación
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .createMenu('Resultados')
      .addItem('Genera Resultados', 'showPrompt')
      .addToUi();
}

//Crea el popup desde el que se pide el DNI
function showPrompt() {

  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
      'Genera tus resultados',
      'Introduce los dígitos de tu DNI/NIE/Pasaporte, SIN NINGUNA LETRA:',
      ui.ButtonSet.OK);

  // Se lee el botón que se ha presionado y el texto que se ha guardado.
  var button = result.getSelectedButton();
  var text = result.getResponseText();
  if (button == ui.Button.OK) {
    // Han seleccionado aceptar.
    // Eliminamos cualquier caracter que haya metido que no sean números
    let textoSaneado = parseInt(text.replaceAll(RegExp("[^0-9]","g"), ''));
    // Lanzamos la función que aleatoriza y muestra los datos
    generateFromSeed(textoSaneado)
  } else if (button == ui.Button.CLOSE) {
    // Han cerrado la ventana presionando la X.
    ui.alert('Has cerrado el diálogo sin generar datos.');
  }
}

/////////////////////////////////////////////////////
//
// PARÁMETROS
//
/////////////////////////////////////////////////////

// Parámetros de ejecución del script
var nombreHojaDatos = "Resultados"
var muestraElegible = 10; //Número de filas elegibles para ser seleccionadas
var numeroFilas = 3; //Número de filas que se van a escoger aleatoriamente dentro de la muestra elegible
var primeraColumnaDatos = 1; //Índice de la primera columna elegible
var numeroMedidas = 6; //Número de variables que queremos extraer
var filaNombres = 1; // Fila en la que están los nombres de las variables

/////////////////////////////////////////////////////
//
// FUNCIONES ALEATORIZACIÓN Y SELECCIÓN
//
/////////////////////////////////////////////////////

function lehmer (seed){
  // Genera valores pseudoaleatorios
  // https://en.wikipedia.org/wiki/Lehmer_random_number_generator
	return(seed*48271 % (2**31-1))
}

function generateOrder(seed, numValues){
  // Crea una lista pseudoaleatoria de numValues valores partiendo de una semilla (seed)
  // y usando cada valor como semilla para el siguiente
	var temp_seed = seed;
	var values = [];
  var curValue = 0;
  for (var curElement = 0; curElement < numValues; ++curElement) {
  	curValue = lehmer(temp_seed)
    values.push(curValue);
    temp_seed = curValue;
  }
  return values;
}

function rankings(arr) {
  //Saca el ranking de la puntuaciones de un array
  const sorted = [...arr].sort((a, b) => b - a);
  return arr.map((x) => sorted.indexOf(x) + 1);
};


function generateFromSeed(seed) {

  // Cargamos las hojas de datos y de parámetros
  var hojaDatos=SpreadsheetApp.getActive().getSheetByName(nombreHojaDatos);

  // Hacemos un array con las filas elegidas, las filas se miran como offset desde filaNombres
  // Al principio ponemos la fila 0 para poner primero los nombres de las variables
  var filasElegidas = [0];
  var aleatorizados = rankings(generateOrder(seed, muestraElegible));

  // Buscamos los índices de fila de las filas elegidas
  for (var curValue = 1; curValue <= numeroFilas; ++curValue){
    filasElegidas.push(aleatorizados.indexOf(curValue)+1);
  }

  // Inicializamos el documento html
  var html='<html><head></head><body>';
  html+='<style>th,td{border:1px solid black;}</style><table>';
  // Añadimos el código del DNI para que se pueda verificar que se ha metido el correcto
  html+=Utilities.formatString(`<p>DNI: %s</p>
  <input id='copy_btn' type='button' value='Copiar'>
  <button type="button" onclick="tableToCSV()">Descarga CSV</button>`,seed);

  // Añadimos las filas, y dentro de cada fila la columna
  for (var fila = 0; fila < numeroFilas + 1; ++fila) {
    html+='<tr>'; //Inicia la fila
    for (var columna = 0; columna < numeroMedidas; ++columna){
        html+=Utilities.formatString('<td>%s</td>', hojaDatos.getRange(filaNombres + filasElegidas[fila], primeraColumnaDatos + columna).getValue());
    }
    html+='</tr>';  
  }
  //Cerramos el html
  html+='</table>'+Utilities.formatString(codeForCsvFunctions,seed)+codeForTableCopy+'</body></html>';
  //Lanzamos la página
  var userInterface=HtmlService.createHtmlOutput(html).setWidth(1000);
  SpreadsheetApp.getUi().showModelessDialog(userInterface, "Datos para DNI/NIE/Pasaporte " + seed);
}

/////////////////////////////////////////////////////
//
// FUNCIONES AUXILIARES (StackOverflow y similares)
//
/////////////////////////////////////////////////////

// Esta función busca el primer elemento de html de tipo table, lo selecciona y lo copia al cortapapeles
var codeForTableCopy = `
<script>
  var copyBtn = document.querySelector("#copy_btn");
  copyBtn.addEventListener("click", function () {
    var urlField = document.querySelector("table");
    var range = document.createRange();
    range.selectNode(urlField);
    window.getSelection().addRange(range);
    document.execCommand("copy")}, false);
</script>`

// Estas dos funciones crean un csv a partir de todos los elementos de tabla existente en el html
// y luego se descarga el fichero creado
var codeForCsvFunctions = `
<script type="text/javascript">
        function tableToCSV() {
 
            // Variable to store the final csv data
            var csv_data = [];
 
            // Get each row data
            var rows = document.getElementsByTagName('tr');
            for (var i = 0; i < rows.length; i++) {
 
                // Get each column data
                var cols = rows[i].querySelectorAll('td,th');
 
                // Stores each csv row data
                var csvrow = [];
                for (var j = 0; j < cols.length; j++) {
 
                    // Get the text data of each cell
                    // of a row and push it to csvrow
                    csvrow.push(cols[j].innerHTML);
                }
 
                // Combine each column value with comma
                csv_data.push(csvrow.join(","));
            }
 
            // Combine each row data with new line character
            csv_data = csv_data.join('\\n');
 
            // Call this function to download csv file 
            downloadCSVFile(csv_data);
         }
 
        function downloadCSVFile(csv_data) {
 
            // Create CSV file object and feed
            // our csv_data into it
            CSVFile = new Blob([csv_data], {
                type: "text/csv"
            });
 
            // Create to temporary link to initiate
            // download process
            var temp_link = document.createElement('a');
 
            // Download csv file
            temp_link.download = "DatosID%s.csv";
            var url = window.URL.createObjectURL(CSVFile);
            temp_link.href = url;
 
            // This link should not be displayed
            temp_link.style.display = "none";
            document.body.appendChild(temp_link);
 
            // Automatically click the link to
            // trigger download
            temp_link.click();
            document.body.removeChild(temp_link);
        }
</script>`
