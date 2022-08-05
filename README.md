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

# Auto llenado

# Referencia absoluta $

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



