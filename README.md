# Excel-de-0-a-100-para-Analizar-tus-datos
Resumen para aprender Excel de 0 a 100 si estas comenzando en el mundo de los Datos.

# ¿Qué es Excel? 
Excel es un programa desarrollado por Microsoft y pertenece al paquete de Office, fue creado para el procesamiento de datos.
Excel es una hoja de cálculo que nos permite manipular datos numéricos y de texto en tablas formadas por la unión de filas y columnas. 
En su pantalla principal se muestra una matriz de dos dimensiones, que está formada por columnas y filas, de esta manera se le da forma a una celda, que básicamente es la intersección de una columna y una fila. A cada fila se le llama registro y a cada columna dato, por lo que la celda contiene datos que corresponden a un registro de un campo.En cada celda podemos introducir principalente tres tipos de datos, valores numericos, texto o formulas y entre ellas, podemos realizar cálculos aritméticos básicos, aplicar funciones matemáticas de mayor complejidad y utilizar funciones de estadísticas o funciones de tipo lógica, entre muchas de sus funciones que contare más adelante.

# ¿Por qué es util para analizar datos?
Porque nos permite organizar y  gestiónar los datos y con ellos cálcular valores y análizarlos para obtener información util. Entre sus muchas funcionalidades destacan las herramientas para el análisis de los mismos, es posible analizar datos estadísticos o técnicos ahorrando tiempo y energía. Y no solo nos sirve para la etapa de procesamiento y analisis de los datos sino que nos permite generar reportes y visualizaciones mediante herramientas de gráficos y las tablas dinámicas. 

# Cálculos aritméticos

En Excel una podemos realizar operaciones aritméticas simples como por ejemplo: sumar (+), restar (-), multiplicar (*), dividir (/). Para poder realizar cálculos aritméticos en Excel, solo debemos poner un (=) o el signo (+) al inicio de la celda, seguido de la fórmula que deseamos ejecutar.

+ (suma)
- (resta o negación)
* (multiplicación)
/ (división)
ˆ (potencia)
% (porcentaje)

Ejemplo de suma
=a+b

# Fórmula
Empezaremos por definir una de las herramientas con la que mas trabajaremos en el analisis de nuestros datos, la fórmula. 
¿Qué es una formula?
Una fórmula es un método  práctico convencional que, a partir de determinados símbolos, reglas, pasos y/o valores, permite resolver problemas o ejecutar procesos de manera ordenada y sistemática, a fin de obtener un resultado específico y controlado.

En Excel debemos escribir las fórmulas siempre comenzando con el signo (=) seguido de la operacion que querramos realizar y podemos hacerlo introduciendo a mano los valores que deseamos utilizar, o referenciando celdas que contienen los valores, se referecian apretando sobre la celda que queremos utilizar, o colocando la letra de la columna seguido del numero de fila. Existen funciones integradas, que han sido predefinidas con la fabricacion de Excel y nos facilitan a la hora de realizar calculos ya que solo debemos introducir los valores que deseamos calcular como los argumentos de la formula. Estan divididas en categorias segun las operaciones que necesitemos realizar. 
Estas son: 
Funciones de búsqueda y referencia
Funciones de texto
Funciones lógicas
Funciones de fecha y hora
Funciones de base de datos
Funciones matemáticas y trigonométricas
Funciones financieras
Funciones estadísticas
Funciones de información
Funciones de ingeniería
Funciones de cubo
Funciones web

SUMAR
SUMAR.SI
SUMAR.SI.CONJUNTO
CONTAR
CONTAR.SI
CONTAR.SI.CONJUNTO
CONTARA
PROMEDIO
PROMEDIO.SI
PROMEDIO.SI.CONJUNTO

# Funciones Lógicas
SI
Y
O

# Funciones de búsqueda y referencia

BUSCAR

BUSCARV: Nos permite hacer una búsqueda de un valor dentro de la primera columna de un rango de datos.
Sintaxis : BUSCARV(valor_deseado,rango_de_busqueda,columna_de_resultado, V/F) 
-valor_deseado: Lo que desea buscar
-rango_de_busqueda: donde se desea buscarlo, el rango debe empezar en la columna donde se desea buscar el valor. 
-columna_de_resultad: el número de columna en el rango que contiene el valor a devolver.
-V/F: devuelve una Coincidencia exacta o Coincidencia aproximada, indicada como 1/VERDADERO, 1/FALSO)
Debemos recordar que el valor del primer argumento de la función será buscado siempre en la primera columna de la tabla de datos. No es posible buscar en una columna diferente que no sea la primera columna. El segundo argumento de la función indica la totalidad del rango que contiene los datos. En este rango es importante asegurase de incluir la columna que vamos a necesitar como resultado.El último argumento de la función es opcional, pero si no proporcionamos un valor, la función BUSCARV hará una búsqueda aproximada. Para que la función realice una búsqueda exacta debemos colocar el valor falso y obtendremos como resultado el valor de la columna que hayamos indicado.
Si la función BUSCARV no encuentra el valor en la columna uno, devolverá el error #N/A.

BUSCARH: busca un valor dentro de una fila y devuelve el valor que ha sido encontrado o un error #N/A en caso de no haberlo encontrado. 
Sintaxis: igual a la sintaxis de BUSCARV. pero la busqueda se realiza en la primera fila del rango, y en el tercer argumento se debe aclarar la posicion de la fila que contiene el resultado a devolver. 
Si la función BUSCARH no encuentra el valor en la fila uno, devolverá el error #N/A.

REEMPLAZAR: reemplaza parte de una cadena de texto, en función del número de caracteres que especifique, por una cadena de texto diferente.
Sintaxis: =REEMPLAZAR(texto_original, núm_inicial, núm_de_caracteres, texto_nuevo)
texto_original:Obligatorio. Es el texto en el que desea reemplazar algunos caracteres.
núm_inicial:Obligatorio. Es la posición del carácter dentro de texto_original que desea reemplazar por texto_nuevo.
núm_de_caracteres:Obligatorio. Es el número de caracteres de texto_original que se desea que REEMPLAZAR reemplace por texto_nuevo.
texto_nuevo:Obligatorio. Es el texto que reemplazará los caracteres de texto_original.

REEMPLAZARB: reemplaza parte de una cadena de texto, en función del número de bytes que especifique, por una cadena de texto diferente.
Sintaxis: =REEMPLAZARB(texto_original, núm_inicial, núm_bytes, texto_nuevo)
Es igual que la funcion reemplazar pero en vez de especificar numero de caracteres, se especifica numero de bytes.

EXTRAE: Nos sirve para extraer determinado número de caracteres de una cadena de texto. 
Sintaxis: = EXTRAE (texto, posición_inicial, núm_de_caracteres)
-texto: La cadena de texto original que contiene el dato que necesitamos extraer.
-posición_inicial: la posición del primer carácter que se desea extraer.
-núm_de_caracteres: número de caracteres a extraer.

CONCATENAR: Nos permite unir dos o más cadenas de texto en una misma celda lo cual es muy útil cuando nos encontramos manipulando bases de datos y necesitamos hacer una concatenación.
Sintaxis: = CONCATENAR (texto1, texto2).
Texto 1 y texto 2 pueden ser celdas referenciadas. Tambien se puede poner por ejemplo: = CONCATENAR (A1," ", B1),  con las comillas separadas por un espacio estamos indicando que los dos textos se separen por un espacio. 



IR A ESPECIAL





# Funciones anidadas

# Funciones de fecha

# Caracteres comodin


# Auto llenado

# Referencia absoluta, referencia relativa y mixta. 
Las referencias de una celda en una hoja de Excel, es la direccion dentro de la hoja y  siempre constará de dos partes: la primera parte indicará la letra (o letras) de la columna a la que pertenece y la segunda parte indicará su número de fila.
Cuando hablamos de los tipos de referencia, estamos hablando de los tipos de comportamiento que tienen las referencias al ser copiadas o trasladadas a otra celda. 

REFERENCIA RELATIVA: De manera predeterminada, las referencias en Excel son relativas. El término relativo significa que al momento de copiar una fórmula, Excel modificará las referencias en relación a la nueva posición donde se está haciendo la copia de la fórmula. Por ejemplo si en la celda C1, escribimos =A1+B1. Cuando copiemos la formula en la celda C2, no realizara la operacion referenciando las celdas A1 y B1, sino que realizara la operacion con las celdas A2 y B2, porque cambia la referencia segun la posicion con las celdas referenciadas. 

REFERENCIA ABSOLUTA: Hay ocasiones en las que necesitamos “fijar” la referencia a una celda de manera que permanezca igual aún después de ser copiada. Si queremos impedir que Excel modifique las referencias de una celda al momento de copiar la fórmula, entonces debemos convertir una referencia relativa en absoluta y eso lo podemos hacer anteponiendo el símbolo “$” a la letra de la columna y al número de la fila de la siguiente manera: si seguimos con el ejemplo anterior si en la celda C1 escribimos la formula = $A$1 + $B$1, y copiamos la formula a la celda C2, la operacion que se va a realizar va a seguir siendo con las celdas A1 y B1 ya que las hemos "fijado". 

REFERENCIA MIXTA: Es similar a la referencia absoluta ya que fijamos una parte de la celda referenciada, por ejemplo su numero de fila o numero de columna. Por ejemplo: Si en la celda C1 escribimos =$A1 + $B1, lo que estamos fijando es el numero de columna a utilizar, por lo que si copiamos la formula a la celda C2, se realizara la formula con $A2 + $B2.

# Errores comunes y como solucionarlos

ERROR #¿NOMBRE?
El tipo de error #¿NOMBRE? se genera cuando una celda hace referencia a una función que no existe. Por ejemplo, si introducimos la fórmula =FORMATOFINAL() obtendremos este tipo de error porque es una función inexistente.

Cuando veas desplegado el error #¿NOMBRE? debes asegurarte de que has escrito correctamente el nombre de la función. Y si estás acostumbrado a utilizar el nombre de las funciones en inglés, pero te encuentras utilizando una versión de Excel en español, debes utilizar su equivalente en español o de lo contrario obtendrás este tipo de error.

ERROR #¡REF!
Cuando una celda intenta hacer referencia a otra celda que no puede ser localizada porque tal vez fue borrada o sobrescrita, entonces obtendremos un error del tipo #¡REF!.

Si obtienes este tipo de error  debes revisar que la función no esté haciendo referencia a alguna  celda que fue eliminada. Este tipo de error es muy común cuando eliminamos filas o columnas que contienen datos que estaban relacionados a una fórmula y al desaparecer se ocasiona que dichas fórmulas muestren el error #¡REF!

ERROR #¡DIV/0!
Cuando Excel detecta que se ha hecho una división entre cero muestra el error #¡DIV/0! Para resolver este error copia el denominador de la división a otra celda e investiga lo que está causando que sea cero.

ERROR #¡VALOR!
El error #¡VALOR! sucede cuando proporcionamos un tipo de dato diferente al que espera una función. Por ejemplo, si introducimos la siguiente función =SUMA(1, “a”) obtendremos el error #¡VALOR! porque la función SUMA espera argumentos del tipo número pero hemos proporcionado un carácter.

Para resolver este error debes verificar que has proporcionado los argumentos del tipo adecuado tal como los espera la función ya sean del tipo texto o número. Tal vez tengas que consultar la definición de la función para asegurarte de que estás utilizando el tipo de datos adecuado.

ERROR #¡NUM!
El error #¡NUM! es el resultado de una operación en Excel que ha sobrepasado sus límites y por lo tanto no puede ser desplegado. Por ejemplo, la fórmula =POTENCIA(1000, 1000) resulta en un número tan grande que Excel muestra el error #¡NUM!

ERROR #¡NULO!
El error #¡NULO! se genera al especificar incorrectamente un rango en una función. Por ejemplo, si tratamos de hacer una suma =A1 + B1 B5 Excel mostrará este tipo de error. Observa que en lugar de especificar el rango B1:B5 he omitido los dos puntos entre ambas celdas.

Este error se corrige revisando que has especificado correctamente los rangos dentro de la fórmula.

ERROR #N/A
Este tipo de error indica que el valor que estamos intentando encontrar no existe. Por esta razón el error #N/A es muy común cuando utilizamos funciones de búsqueda como BUSCARV o BUSCARH. Cuando la función BUSCARV no encuentra el valor que estamos buscando, regresa el error de tipo #N/A.

# ¿Cómo saber si un error, es efectivamente un error? 
Existen algunas funciones de información que nos permiten saber si un valor es efectivamente un error. Las funciones que nos ayuda en esta tarea son: la función ESNOD, la función ESERR y la función ESERROR. 

# Funcion util: ¿Cómo separar texto en filas y columnas? 
La funcion se llama "Texto en columnas" nos va a permitir separar texto en filas y columnas los archivos de texto que pueden venir en formatos .csv o para separa texto que pueden venir juntos en una columna como NOMBRE Y APELLIDO.  Primero seleccionamos la columna que contiene el texto, luego nos dirigimos a la pestaña "Datos", se encuentra la seccion "Herramienta de datos" donde esta la opcion "Texto en columnas". Si nuestro texto esta por ejemplo separado por comas, se despliega una ventana donde eberemos seleccion la opcion Delimitados,  en el que podemos elegir que tipo de separados utilizar, en este caso, comas.



