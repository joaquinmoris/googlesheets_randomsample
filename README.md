# googlesheets_randomsample


Crea un nuevo menú que permite seleccionar una muestra pseudoaleatorizada a partir de un índice (DNI) seleccionando filas que estén en la hoja seleccionada. Se pueden personalizar los siguientes parámetros:

Parámetros de ejecución del script

+ var nombreHojaDatos = "Resultados" // Nombre de la página en la que se han guardado los datos
+ var muestraElegible = 10; //Número de filas elegibles para ser seleccionadas
+ var numeroFilas = 3; //Número de filas que se van a escoger aleatoriamente dentro de la muestra elegible
+ var primeraColumnaDatos = 1; //Índice de la primera columna elegible
+ var numeroMedidas = 6; //Número de variables que queremos extraer
+ var filaNombres = 1; // Fila en la que están los nombres de las variables, el resto de las filas se seleccionan como offset a partir de esta
