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

- Suma (+)
- Resta o negación (-)
- Multiplicación (*)
- División (/)
- Potencia (^)
- Porcentaje (%)
Ejemplo de suma
=a+b

# Fórmula
Empezaremos por definir una de las herramientas con la que mas trabajaremos en el analisis de nuestros datos, la fórmula. 
¿Qué es una formula?
Una fórmula es un método  práctico convencional que, a partir de determinados símbolos, reglas, pasos y/o valores, permite resolver problemas o ejecutar procesos de manera ordenada y sistemática, a fin de obtener un resultado específico y controlado.

En Excel debemos escribir las fórmulas siempre comenzando con el signo (=) seguido de la operacion que querramos realizar y podemos hacerlo introduciendo a mano los valores que deseamos utilizar, o referenciando celdas que contienen los valores, se referecian apretando sobre la celda que queremos utilizar, o colocando la letra de la columna seguido del numero de fila. Existen funciones integradas, que han sido predefinidas con la fabricacion de Excel y nos facilitan a la hora de realizar calculos ya que solo debemos introducir los valores que deseamos calcular como los argumentos de la formula. Estan divididas en categorias segun las operaciones que necesitemos realizar. 
Estas son: 
- Funciones de búsqueda y referencia
- Funciones de texto
- Funciones lógicas
- Funciones de fecha y hora
- Funciones de base de datos
- Funciones matemáticas y trigonométricas
- Funciones financieras
- Funciones estadísticas
- Funciones de información
- Funciones de ingeniería
- Funciones de cubo
- Funciones web


# Funciones de base de dato

- SUMAR: suma valores. Puede sumar valores individuales, referencias o rangos de celda o una combinación de las tres.
Sintaxis: =SUMA(rango_de_suma)

- SUMAR.SI: se usa para sumar los valores intervalo de un rango que cumplan los criterios que haya especificado.
Sintaxis: =SUMAR.SI(rango; criterio; [rango_suma])
Rango: Obligatorio. El rango de celdas que se desea evaluar según los criterios. Las celdas de cada rango deben ser números, nombres, matrices o referencias que contengan números. Los valores en blanco y de texto se ignoran. 
criterio: Obligatorio. Es el criterio en forma de número, expresión, referencia de celda, texto o función que determina las celdas que va a sumar. Se pueden incluir caracteres comodín. Por ejemplo, los criterios pueden expresarse como 32, ">32", B5, "3?", "manzanas*", “*~” u HOY().
Importante: Cualquier criterio de texto o cualquier criterio que incluya los símbolos lógicos o matemáticos debe estar entre comillas dobles ("). Si el criterio es numérico, las comillas dobles no son necesarias.
Rango_suma: Opcional. Son las celdas reales que se sumarán, si es que desea sumar celdas a las ya especificadas en el argumento rango. Si omite el argumento rango_suma, Excel suma las celdas especificadas en el argumento rango (las mismas celdas a las que se aplica el criterio).

- SUMAR.SI.CONJUNTO: igual que la funcion SUMAR.SI, pero con mas de un criterio.
Sintaxis: =SUMAR.SI.CONJUNTO(rango_suma; rango_criterios1; criterios1; [rango_criterios2; criterios2];...)

- CONTAR: cuenta la cantidad de celdas que contienen números y cuenta los números dentro de la lista de argumentos, se usa para obtener la cantidad de entradas en un campo de número de un rango o matriz de números.
Sintaxis: =CONTAR(valor1; [valor2]; ...)
valor1: Obligatorio. Primer elemento, referencia de celda o rango en el que desea contar números.
valor2: Opcional. Hasta 255 elementos, celdas de referencia o rangos adicionales en los que desea contar números.

- CONTAR.SI: sirve para contar el número de celdas que cumplen un criterio.
- Sintaxis: =CONTAR.SI(rango;criterios)

- CONTAR.SI.CONJUNTO: Es igual a la funcion CONTAR.SI pero con mas de un criterio, aplica criterios a las celdas de varios rangos y cuenta el número de veces que se cumplen todos los criterios.
Sintaxis: =CONTAR.SI.CONJUNTO(rango_criterios1; criterios1; [rango_criterios2; criterios2];…)

- CONTARA: cuenta la cantidad de celdas que no están vacías en un intervalo.
Sintaxis: =CONTARA(valor1; [valor2]; ...)

- PROMEDIO: Devuelve el promedio (media aritmética) de los argumentos. 
Sintaxis: =PROMEDIO(número1; [número2]; ...)
Número1: Obligatorio. El primer número, referencia de celda o rango para el cual desea el promedio.
Número2, ...: Opcional. Números, referencias de celda o rangos adicionales para los que desea el promedio, hasta un máximo de 255.

- PROMEDIO.SI: Devuelve el promedio (media aritmética) de todas las celdas de un rango que cumplen unos criterios determinados.
Sintaxis: =PROMEDIO.SI(rango; criterios; [rango_promedio])
Rango: Obligatorio. Una o más celdas cuyo promedio se desea obtener que incluyan números, o nombres, matrices o referencias que contengan números.
Criterio: Obligatorio. Criterio en forma de número, expresión, referencia de celda o texto que determina las celdas cuyo promedio se va a obtener. Por ejemplo, los criterios pueden expresarse como 32, "32", ">32", "manzanas" o B4.
Rango_promedio: Opcional. Conjunto real de celdas cuyo promedio se va a calcular. Si se omite, se utiliza un rango.

- PROMEDIO.SI.CONJUNTO: Igual que la funcion PROMEDIO.SI pero con mas de un criterio. 

- MIN: Devuelve el valor mínimo de un conjunto de valores.
Sintaxis: =MIN(número1, [número2], ...)

- MAX: Devuelve el valor máximo de un conjunto de valores.
Sintaxis: =MAX(número1, [número2], ...)

- PRODUCTO: La función PRODUCTO multiplica todos los números dados como argumentos y devuelve el producto.También puede realizar la misma operación con el operador matemático multiplique (*); por ejemplo, =A1 * A2.La función PRODUCTO es útil cuando necesita multiplicar varias celdas juntas. Por ejemplo, la fórmula =PRODUCTO(A1:A3, C1:C3) equivale a =A1 * A2 * A3 * C1 * C2 * C3.
- Sintaxis: =PRODUCTO(número1, [número2], ...)

- DESVESTA: Calcula la desviación estándar de una muestra. La desviación estándar es la medida de la dispersión de los valores respecto a la media (valor promedio).
Sintaxis: =DESVESTA(valor1, [valor2], ...)

- VAR: Calcula la varianza de una muestra.
Sintaxis: =VAR(número1,[número2],...)

- COCIENTE: devuelve la parte entera de una division. 
Sintaxis:COCIENTE(numerador, denominador)

La sintaxis de la función COCIENTE tiene los siguientes argumentos:

Numerador    Obligatorio. Es el dividendo.

Denominador    Obligatorio. Es el divisor.

- FACT: Devuelve el factorial de un número. El factorial de un número es igual a 1*2*3*...* número.

Sintaxis
FACT(número)

Número: es el número no negativo cuyo factorial se desea obtener. Si el valor de número no es un entero, se truncará.

- POTENCIA: (POWER en inglés) Eleva un número a una potencia especificada.

Sintaxis: POTENCIA(número, potencia)

- PRODUCTO: La función PRODUCTO multiplica todos los números dados como argumentos y devuelve el producto.
Sintaxis: PRODUCTO(número1, [número2], ...)

La sintaxis de la función PRODUCTO tiene los siguientes argumentos:

Número1    Obligatorio. Es el primer número o intervalo que desea multiplicar.

Número2, ...    Opcional. Son los números o rangos adicionales que desea multiplicar, hasta un máximo de 255 argumentos.

- RAIZ: Devuelve la raíz cuadrada de un número.

Sintaxis: RAIZ(número)

La sintaxis de la función RAIZ tiene los siguientes argumentos:

Número    Obligatorio. Es el número cuya raíz cuadrada desea obtener.

- RESIDUO: Devuelve el residuo o resto de la división entre número y divisor. El resultado tiene el mismo signo que divisor.

Sintaxis: RESIDUO(número, divisor)

La sintaxis de la función RESIDUO tiene los siguientes argumentos:

Número    Obligatorio. Es el número cuyo resto desea obtener.

Divisor    Obligatorio. Es el número por el cual desea dividir el argumento número.


# Funciones Lógicas

- SI: permite realizar comparaciones lógicas entre un valor y un resultado que espera, una instrucción SI puede tener dos resultados. El primer resultado es si la comparación es Verdadera y el segundo si la comparación es Falsa.
Sintaxis: =SI(Prueba_logica,Valor_si_verdadero,Valor_si_falso)
(Prueba_lógica (obligatorio): Expresión lógica que será evaluada para conocer si el resultado es VERDADERO o FALSO.
Valor_si_verdadero (opcional): El valor que se devolverá en caso de que el resultado de la Prueba_lógica sea VERDADERO.
Valor_si_falso (opcional): El valor que se devolverá si el resultado de la evaluación es FALSO.

- Y: expresion logica para determinar si todas las condiciones de una prueba son VERDADERAS. 
Sintaxis: =SI(Valor_logico1,Valor_logico2)
Valor_lógico1 (obligatorio): Expresión lógica que será evaluada por la función.
Valor_lógico2 (opcional): Expresiones lógicas a evaluar, opcional hasta un máximo de 255.
La función Y solamente regresará el valor VERDADERO si todas las expresiones lógicas evaluadas son verdaderas. Bastará con que una sola expresión sea falsa para que la función Y tenga un resultado FALSO.

- O: Expresion logica que nos muestra verdadero si alguno de los argumentos especificados es verdadero desde el punto de vista lógico, y falso si todos los argumentos son falsos.
Sintaxis: O((Valor_logico1,Valor_logico2)

# Funciones anidadas

# Funciones de búsqueda y referencia

- BUSCAR: Permite buscar en una sola fila o columna y encontrar un valor desde la misma posición en una segunda fila o columna.
Sintaxis: =BUSCAR(valor_buscado; vector_de_comparación; [vector_resultado])
valor buscado: Obligatorio. Es el valor que busca la función BUSCAR en el primer vector, puede ser un número, texto, un valor lógico o un nombre de referencia que se refiere a un valor.
Vector_de_comparación: Obligatorio. Es un rango que solo contiene una fila o una columna, pueden ser texto, números o valores lógicos.
vector_resultado    Opcional. Un rango que solo contiene una fila o una columna. El argumento vector_result debe tener el mismo tamaño que vector_de_comparación. Debe tener el mismo tamaño.
Si la función BUSCAR no puede encontrar el valor_buscado, la función muestra el valor mayor en vector_de_comparación, que es menor o igual que el valor_buscado.
Si el valor_buscado es menor que el menor valor del vector_de_comparación, BUSCAR devuelve el valor de error #N/A.

- BUSCARV: Nos permite hacer una búsqueda de un valor dentro de la primera columna de un rango de datos.
Sintaxis : BUSCARV(valor_deseado,rango_de_busqueda,columna_de_resultado, V/F) 
-valor_deseado: Lo que desea buscar
-rango_de_busqueda: donde se desea buscarlo, el rango debe empezar en la columna donde se desea buscar el valor. 
-columna_de_resultad: el número de columna en el rango que contiene el valor a devolver.
-V/F: devuelve una Coincidencia exacta o Coincidencia aproximada, indicada como 1/VERDADERO, 1/FALSO)
Debemos recordar que el valor del primer argumento de la función será buscado siempre en la primera columna de la tabla de datos. No es posible buscar en una columna diferente que no sea la primera columna. El segundo argumento de la función indica la totalidad del rango que contiene los datos. En este rango es importante asegurase de incluir la columna que vamos a necesitar como resultado.El último argumento de la función es opcional, pero si no proporcionamos un valor, la función BUSCARV hará una búsqueda aproximada. Para que la función realice una búsqueda exacta debemos colocar el valor falso y obtendremos como resultado el valor de la columna que hayamos indicado.
Si la función BUSCARV no encuentra el valor en la columna uno, devolverá el error #N/A.

- BUSCARH: busca un valor dentro de una fila y devuelve el valor que ha sido encontrado o un error #N/A en caso de no haberlo encontrado. 
Sintaxis: igual a la sintaxis de BUSCARV. pero la busqueda se realiza en la primera fila del rango, y en el tercer argumento se debe aclarar la posicion de la fila que contiene el resultado a devolver. 
Si la función BUSCARH no encuentra el valor en la fila uno, devolverá el error #N/A.

- REEMPLAZAR: reemplaza parte de una cadena de texto, en función del número de caracteres que especifique, por una cadena de texto diferente.
Sintaxis: =REEMPLAZAR(texto_original, núm_inicial, núm_de_caracteres, texto_nuevo)
texto_original:Obligatorio. Es el texto en el que desea reemplazar algunos caracteres.
núm_inicial:Obligatorio. Es la posición del carácter dentro de texto_original que desea reemplazar por texto_nuevo.
núm_de_caracteres:Obligatorio. Es el número de caracteres de texto_original que se desea que REEMPLAZAR reemplace por texto_nuevo.
texto_nuevo:Obligatorio. Es el texto que reemplazará los caracteres de texto_original.

- REEMPLAZARB: reemplaza parte de una cadena de texto, en función del número de bytes que especifique, por una cadena de texto diferente.
Sintaxis: =REEMPLAZARB(texto_original, núm_inicial, núm_bytes, texto_nuevo)
Es igual que la funcion reemplazar pero en vez de especificar numero de caracteres, se especifica numero de bytes.

- EXTRAE: Nos sirve para extraer determinado número de caracteres de una cadena de texto. 
Sintaxis: = EXTRAE (texto, posición_inicial, núm_de_caracteres)
-texto: La cadena de texto original que contiene el dato que necesitamos extraer.
-posición_inicial: la posición del primer carácter que se desea extraer.
-núm_de_caracteres: número de caracteres a extraer.

- CONCATENAR: Nos permite unir dos o más cadenas de texto en una misma celda lo cual es muy útil cuando nos encontramos manipulando bases de datos y necesitamos hacer una concatenación.
Sintaxis: = CONCATENAR (texto1, texto2).
Texto 1 y texto 2 pueden ser celdas referenciadas. Tambien se puede poner por ejemplo: = CONCATENAR (A1," ", B1),  con las comillas separadas por un espacio estamos indicando que los dos textos se separen por un espacio. 

- LEN: devuelve el número de caracteres de una cadena de texto.
Sintaxis: =LEN(texto)
Texto: es el texto cuya longitud se desea obtener. Los espacios cuentan como caracteres.

- COINCIDIR:  busca un elemento determinado en un intervalo de celdas y después devuelve la posición relativa de dicho elemento en el rango. Por ejemplo, si el rango A1:A3 contiene los valores 5, 25 y 38, la fórmula =COINCIDIR(25,A1:A3,0) devuelve el número 2, porque 25 es el segundo elemento del rango.
Sintaxis: =COINCIDIR(valor_buscado,matriz_buscada, [tipo_de_coincidencia])

- DERECHA: Devuelve un valor de tipo Variant (String) que contiene un número especificado de caracteres del lado derecho de una cadena.
Sintaxis: =Derecha( cadena, longitud )

- IZQUIERDA: Devuelve un valor de tipo Variant (String) que contiene un número específico de caracteres a partir del lado izquierdo de una cadena.
Sintaxis: =Izquierda( cadena, longitud )

- MID: devuelve un número concreto de caracteres de una cadena de texto, empezando en la posición especificada y basándose en el número de caracteres que se especifique.
Sintaxis: MID(texto,posición_inicial,núm_de_caracteres)
Texto: Cadena de texto que contiene los caracteres que se desea extraer.
Posición_inicial: Posición del primer carácter que se desea extraer del texto. La posición_inicial del primer carácter de texto es 1, y así sucesivamente.
Núm_de_caracteres: especifica el número de caracteres del texto que MID debe devolver.

- ESPACIOS: Quita los espacios iniciales, finales y repetidos del texto. 
Sintaxis: =ESPACIOS(texto)
texto: String o referencia a una celda que contiene una string a la que se le van a quitar espacios.

# Funciones de fecha
- AHORA: Devuelve el número de serie de la fecha y hora actuales.
Sintaxis: =AHORA()
La sintaxis de la función AHORA no tiene argumentos.

- FECHA: Use la función FECHA de Excel cuando necesite tomar tres valores diferentes y combinarlos para formar una fecha. Sintaxis: =(dia,mes,año) pueden introducirse manualmente los valores o referenciando celdas.
Sintaxis: FECHA(año; mes; día)

- SIFECHA: Calcula el número de días, meses o años entre dos fechas.
Sintaxis: SIFECHA(fecha_inicial;fecha_final;unidad)
fecha_inicial: Una fecha que representa la primera o la fecha inicial de un período determinado.
fecha_final: Una fecha que representa la última del período o al fecha de finalización.
El tipo de información que desea devolver, donde:
Unidad :Devuelve
"Y": El número de años completos en el período.
"M": El número de meses completos en el período.
"D": El número de días en el período.
"MD": La diferencia entre los días en fecha_inicial y fecha_final. Los meses y años de las fechas se pasan por alto.
"YM": La diferencia entre los meses de fecha_inicial y fecha_final. Los días y años de las fechas se pasan por alto
"YD": La diferencia entre los días de fecha_inicial y fecha_final. Los años de las fechas se pasan por alto.

- AÑO: Sintaxis: AÑO(núm_de_serie)
Núm_de_serie: Obligatorio. Es la fecha del año que desea buscar. Debe especificar las fechas con la función FECHA o como resultado de otras fórmulas o funciones.
- MES: Sintaxis: MES(núm_de_serie) . IDEM AÑO.
- DIA: Sintaxis: DIA(núm_de_serie). IDEM AÑO. 
- HORA: Sintaxis: DIA(núm_de_serie). IDEM AÑO.
- MINUTO: Sintaxis: MINUTO(núm_de_serie). IDEM AÑO.



# Caracteres comodin
Los caracteres comodín son caracteres especiales que pueden representar caracteres desconocidos en un valor de texto y son prácticos para encontrar varios elementos con datos similares pero no idénticos. Los caracteres comodín también le pueden ayudar a obtener datos basados en la coincidencia de un patrón específico.
Sirven para buscar un elemento específico cuando uno no recuerda cómo se escribe. Estos son:
- ? : Sustituye un solo caracter, es decir, hace coincidir un carácter alfabético individual en una posición concreta.
Ejemplo: b?l encuentra bala, billete y bola.
- * : Sustituye cualquier número de caracteres, es decir, puede utilizar el asterisco (*) en cualquier sitio de una cadena de caracteres. 
Ejemplo: qu* encuentra qué, quién y quizás pero no aquellos ni aunque.
- ~ seguido de ? o *: permite incorporar el * o ? en el criterio como caracteres y no como caracteres comodin. 
Ejemplo: ho~*, encuentra hora*, hola*, hoja* 
[ ]: hace coincidir los caracteres incluidos entre los corchetes.
Ejemplo: b[ao]l encuentra bala y bola pero no billete.
- !: Excluye los caracteres incluidos entre los corchetes.
Ejemplo: r[!oc]a encuentra risa y rema pero no roca ni rosa.
Igual que “[!a]*” encuentra todos los elementos que no empiezan con la letra a.
- -: Hace coincidir cualquier intervalo de caracteres. Recuerde que debe especificar los caracteres en orden ascendente (de la A a la Z, no de la Z a la A).
Ejemplo: a[m-s]a encuentra ama, ata y asa
- #: Hace coincidir cualquier carácter numérico.
Ejemplo: 1#3 encuentra 103, 113 y 123.

# Auto relleno: 
Sirve para rellenar celdas con datos que siguen un patrón o que se basan en datos de otras celdas.
¿Cómo realizarlo?
Seleccione las celdas que quiera usar como base para rellenar celdas adicionales.
Para una serie como “1, 2, 3, 4, 5…”, escriba 1 y 2 en las primeras dos celdas. Para la serie “2, 4, 6, 8…”, escriba 2 y 4.
Para la serie “2, 2, 2, 2…”, escriba 2 solo en la primera celda.
En la esquina inferior derecha aparece un cuadrado negro que es el controlador de relleno, haga click y arrastrelo en las celdas que desea que se complete el autorellenbo. 
Si es necesario, haga clic en un boton que aparece que es de las: Opciones de autorrelleno y seleccione la opción que quiera.

# Referencia absoluta, referencia relativa y mixta. 
Las referencias de una celda en una hoja de Excel, es la direccion dentro de la hoja y  siempre constará de dos partes: la primera parte indicará la letra (o letras) de la columna a la que pertenece y la segunda parte indicará su número de fila.
Cuando hablamos de los tipos de referencia, estamos hablando de los tipos de comportamiento que tienen las referencias al ser copiadas o trasladadas a otra celda. 

- REFERENCIA RELATIVA: De manera predeterminada, las referencias en Excel son relativas. El término relativo significa que al momento de copiar una fórmula, Excel modificará las referencias en relación a la nueva posición donde se está haciendo la copia de la fórmula. Por ejemplo si en la celda C1, escribimos =A1+B1. Cuando copiemos la formula en la celda C2, no realizara la operacion referenciando las celdas A1 y B1, sino que realizara la operacion con las celdas A2 y B2, porque cambia la referencia segun la posicion con las celdas referenciadas. 

- REFERENCIA ABSOLUTA: Hay ocasiones en las que necesitamos “fijar” la referencia a una celda de manera que permanezca igual aún después de ser copiada. Si queremos impedir que Excel modifique las referencias de una celda al momento de copiar la fórmula, entonces debemos convertir una referencia relativa en absoluta y eso lo podemos hacer anteponiendo el símbolo “$” a la letra de la columna y al número de la fila de la siguiente manera: si seguimos con el ejemplo anterior si en la celda C1 escribimos la formula = $A$1 + $B$1, y copiamos la formula a la celda C2, la operacion que se va a realizar va a seguir siendo con las celdas A1 y B1 ya que las hemos "fijado". 

- REFERENCIA MIXTA: Es similar a la referencia absoluta ya que fijamos una parte de la celda referenciada, por ejemplo su numero de fila o numero de columna. Por ejemplo: Si en la celda C1 escribimos =$A1 + $B1, lo que estamos fijando es el numero de columna a utilizar, por lo que si copiamos la formula a la celda C2, se realizara la formula con $A2 + $B2.

# Herramientas para limpieza de datos

## Quitar filas duplicadas

- Formato condicional
El formato condicional cambia el aspecto de un rango de celdas en función de una condición (o criterios). Puede usar formato condicional para resaltar celdas que contienen valores que cumplen cierta condición. También puede aplicar formato a un rango de celdas y variar el formato exacto cuando varía el valor de cada celda.

- Eliminar duplicados
 Use el formato condicional para buscar y resaltar datos duplicados o filtrar duplicados primero. De esa manera puede revisar duplicados y decidir si desea eliminarlos.
 Cuando use la característica Quitar duplicados, los datos duplicados se eliminarán de manera permanente. Antes de eliminar los duplicados, es una buena idea copiar los datos originales a otra hoja de cálculo para que no pierda ninguna información de forma accidental.

Primero: seleccione el rango de celdas con valores duplicados que desea quitar.
Segundo: Haga clic en Datos > Quitar duplicados y, a continuación, debajo de Columnas, active o desactive las columnas donde desea eliminar los duplicados.
Tercero: Haga click en aceptar.

- IR A ESPECIAL:
El cuadro de ir a especial nos sirve para localizar ciertos tipos de celda dentro de la hoja de cálculo. Podemos acceder al cuadro Ir a especial desde el extremo derecho de la ficha Inicio. Solo tenemos que pulsar en Buscar y seleccionar y luego en Ir a especial. Una vez se abre el cuadro, podemos elegir entre otras opciones, comentarios, constantes, fórmulas, espacios en blanco, etc… 


- UNICOS: La función UNICOS devuelve una lista de valores únicos de una lista o rango. 

Sintaxis: UNICO( rango) 


## Buscar y reemplazar texto 

Utilizar los comandos buscar y buscar y reemplazar. Como tambien las funciones de BUSCARV, BUSCARH, DERECHA, IZQUIERDA, EXTRAE, SUSTITUIR, entre otras. 

## Cambiar a mayuscula, minuscula o nompropio.

A veces el texto es una mezcla, especialmente cuando se refiere a las mayúsculas y minúsculas. Al usar una o varias de las tres funciones de mayúsculas o minúsculas, puedes convertir texto en minúsculas, como direcciones de correo electrónico, mayúsculas, como los códigos, o mayúsculas o minúsculas, como nombres o apellidos.

- MAYUSC(rango a convertir)
- MINUSC( rango a convertir)
- NOMPROPIO(rango a convertir) : pone la primera letra mayuscula y las otras minusculas. 


## Quitar espacios y caracteres no imprimibles del texto

A veces los valores de texto contienen caracteres de espacio incrustado en la primera parte, al final o en varios sitios.

A menudo, estos caracteres pueden producir resultados inesperados al ordenar, filtrar o buscar. 

- Función ESPACIOS

Elimina los espacios del texto, excepto el espacio normal que se deja entre palabras. Utiliza ESPACIOS en texto procedente de otras aplicaciones que pueda contener un espaciado irregular.

Sintaxis

ESPACIOS (texto)

La sintaxis de la función ESPACIOS tiene los siguientes argumentos:

Texto Obligatorio. Es el texto del que deseas quitar espacios.

- Funcion SUSTITUIR

Sustituye texto_original por texto_nuevo dentro de una cadena de texto. Utiliza SUSTITUIR para reemplazar texto específico en una cadena de texto.

Sintaxis

SUSTITUIR (texto, texto_original, texto_nuevo, [núm_de_ocurrencia])

La sintaxis de la función SUSTITUIR tiene los siguientes argumentos:

Texto Obligatorio. Es el texto o la referencia a una celda que contiene el texto en el que deseas sustituir caracteres.

Texto_original Obligatorio. Es el texto que deseas sustituir.

Texto_nuevo Obligatorio. Es el texto por el que deseas reemplazar el texto_original.

Núm_de_ocurrencia Opcional. Especifica la instancia de texto_original que se desea reemplazar por texto_nuevo. Si especifica el argumento núm_de_ocurrencia, solo se remplaza esa instancia de texto_original. De lo contrario, todas las instancias de texto_original en texto se sustituirán por texto_nuevo.


## Convertir fechas almacenadas como texto en fechas

A veces, las fechas pueden adquirir formato de texto y almacenarse como texto en las celdas. Por ejemplo, es posible que hayas escrito una fecha en una celda con formato de texto o que los datos se hayan importado o pegado desde un origen de datos externo como texto.

Las fechas con formato de texto se alinean en una celda a la izquierda (en lugar de a la derecha)

- FECHANUMERO: convierte una fecha almacenada como texto en un número de serie que Excel reconoce como fecha. Para ver un número de serie como una fecha, debe aplicar un formato de fecha a la celda. 
¿Qué es un número de serie de Excel?
Excel almacena las fechas como números de serie secuenciales para que se puedan usar en cálculos. De forma predeterminada, el 1 de enero de 1900 es el número de serie 1 y el 1 de enero de 2008, que es el número de serie 39448 porque es 39.448 días después del 1 de enero 1900. Para copiar la fórmula de conversión en un rango de celdas contiguas, selecciona la celda que contiene la fórmula introducida y, a continuación, arrastra el controlador de relleno  un rango de celdas vacías que coincida en tamaño con el rango de celdas que contiene las fechas de texto. En la pestaña Inicio, haz clic en el selector de la ventana emergente junto a número.
En el cuadro Categoría, haz clic en Fecha y, en la lista Tipo, haz clic en el formato de fecha que desees.

Sintaxis:FECHANUMERO(texto_de_fecha)

La sintaxis de la función FECHANUMERO tiene los siguientes argumentos:

Texto_de_fecha    Obligatorio. Texto que representa una fecha en el formato de fechas de Excel o una referencia a una celda que contiene texto que representa una fecha en un formato de fechas de Excel. 

- REDONDEAR: Cambiar el número de posiciones decimales mostradas sin cambiar el número.
Descripción
La función REDONDEAR redondea un número a un número de decimales especificado.
Sintaxis: REDONDEAR(número; núm_decimales)

La sintaxis de la función REDONDEAR tiene los siguientes argumentos:

número    Obligatorio. Es el número que desea redondear.

núm_decimales    Obligatorio. Es el número de decimales al que desea redondear el argumento número.

- REDONDEAR.MAS: Redondea un número hacia arriba, en dirección contraria a cero.

Sintaxis
REDONDEAR.MAS(número; núm_decimales)

La sintaxis de la función REDONDEAR.MAS tiene los siguientes argumentos:

Número    Obligatorio. Cualquier número real que se desea redondear hacia arriba.

Núm_decimales    Obligatorio. El número de dígitos al que se desea redondear el número.

- REDONDEAR.MENOS: Redondea un número hacia abajo, en dirección hacia cero.

Sintaxis: REDONDEAR.MENOS(número; núm_decimales)

La sintaxis de la función REDONDEAR.MENOS tiene los siguientes argumentos:

Número    Obligatorio. Cualquier número real que se desea redondear hacia abajo.

Núm_decimales    Obligatorio. El número de dígitos al que se desea redondear el número.
## Combinar y dividir columnas

- ¿Cómo separar texto en filas y columnas? 
La funcion se llama "Texto en columnas" nos va a permitir separar texto en filas y columnas los archivos de texto que pueden venir en formatos .csv o para separa texto que pueden venir juntos en una columna como NOMBRE Y APELLIDO.  Primero seleccionamos la columna que contiene el texto, luego nos dirigimos a la pestaña "Datos", se encuentra la seccion "Herramienta de datos" donde esta la opcion "Texto en columnas". Si nuestro texto esta por ejemplo separado por comas, se despliega una ventana donde eberemos seleccion la opcion Delimitados,  en el que podemos elegir que tipo de separados utilizar, en este caso, comas.

- Utilizar la funcion concatenar. 


## Transformar y reorganizar columnas y filas

- Transponer:
Si tiene una hoja de cálculo con datos en columnas que necesita girar para reorganizarla en filas, o viceversa, se usa la funcion TRANSPONER.
Sintaxis: 
Paso 1: Seleccionar celdas en blanco
En primer lugar seleccione varias celdas en blanco. Pero asegúrese de seleccionar el mismo número de celdas que en el conjunto de celdas original, pero en la dirección contraria.
Paso 2: Escribir =TRANSPONER(
Con las mismas celdas en blanco seleccionadas, escriba: =TRANSPONER()
Paso 3: Escribir el rango de las celdas originales.
Ahora, escriba el rango de las celdas que desea transponer.
Paso 4: Para finalizar, presione CTRL+MAYÚS+ENTRAR.
Ahora presione CTRL+MAYÚS+ENTRAR. ¿Por qué? Porque la función TRANSPONER solo se utiliza en fórmulas de matriz y esta es la forma de terminar una fórmula de matriz. En resumen, una fórmula de matriz es una fórmula que se aplica a más de una celda. Como ha seleccionado más de una celda en el paso 1 (lo hizo, ¿verdad?), la fórmula se aplicará a más de una celda.



# Errores comunes y como solucionarlos

- ERROR #¿NOMBRE?
El tipo de error #¿NOMBRE? se genera cuando una celda hace referencia a una función que no existe. Por ejemplo, si introducimos la fórmula =FORMATOFINAL() obtendremos este tipo de error porque es una función inexistente.

Cuando veas desplegado el error #¿NOMBRE? debes asegurarte de que has escrito correctamente el nombre de la función. Y si estás acostumbrado a utilizar el nombre de las funciones en inglés, pero te encuentras utilizando una versión de Excel en español, debes utilizar su equivalente en español o de lo contrario obtendrás este tipo de error.

- ERROR #¡REF!
Cuando una celda intenta hacer referencia a otra celda que no puede ser localizada porque tal vez fue borrada o sobrescrita, entonces obtendremos un error del tipo #¡REF!.

Si obtienes este tipo de error  debes revisar que la función no esté haciendo referencia a alguna  celda que fue eliminada. Este tipo de error es muy común cuando eliminamos filas o columnas que contienen datos que estaban relacionados a una fórmula y al desaparecer se ocasiona que dichas fórmulas muestren el error #¡REF!

- ERROR #¡DIV/0!
Cuando Excel detecta que se ha hecho una división entre cero muestra el error #¡DIV/0! Para resolver este error copia el denominador de la división a otra celda e investiga lo que está causando que sea cero.

- ERROR #¡VALOR!
El error #¡VALOR! sucede cuando proporcionamos un tipo de dato diferente al que espera una función. Por ejemplo, si introducimos la siguiente función =SUMA(1, “a”) obtendremos el error #¡VALOR! porque la función SUMA espera argumentos del tipo número pero hemos proporcionado un carácter.

Para resolver este error debes verificar que has proporcionado los argumentos del tipo adecuado tal como los espera la función ya sean del tipo texto o número. Tal vez tengas que consultar la definición de la función para asegurarte de que estás utilizando el tipo de datos adecuado.

- ERROR #¡NUM!
El error #¡NUM! es el resultado de una operación en Excel que ha sobrepasado sus límites y por lo tanto no puede ser desplegado. Por ejemplo, la fórmula =POTENCIA(1000, 1000) resulta en un número tan grande que Excel muestra el error #¡NUM!

- ERROR #¡NULO!
El error #¡NULO! se genera al especificar incorrectamente un rango en una función. Por ejemplo, si tratamos de hacer una suma =A1 + B1 B5 Excel mostrará este tipo de error. Observa que en lugar de especificar el rango B1:B5 he omitido los dos puntos entre ambas celdas.

Este error se corrige revisando que has especificado correctamente los rangos dentro de la fórmula.

- ERROR #N/A
Este tipo de error indica que el valor que estamos intentando encontrar no existe. Por esta razón el error #N/A es muy común cuando utilizamos funciones de búsqueda como BUSCARV o BUSCARH. Cuando la función BUSCARV no encuentra el valor que estamos buscando, regresa el error de tipo #N/A.

# ¿Cómo saber si un error, es efectivamente un error? 
Existen algunas funciones de información que nos permiten saber si un valor es efectivamente un error. Las funciones que nos ayuda en esta tarea son: la función ESNOD, la función ESERR y la función ESERROR. 



# Formulas utilizadas en negocio

- Variación porcentual
Calcula incrementos y disminuciones porcentuales

Fórmula:

(porcentaje final - porcentaje inicial)/porcentaje inicial

- Aplicar porcentaje de incremento y descuentos
Es la fórmula que utilizamos cuando queremos aplicar IVA y otros impuestos.

Fórmula:

precio inicial + precio inicial * porcentaje

Para aplicar descuentos, cambiamos el signo + por -

- Promedio ponderado
Normalmente, cuando se calcula un promedio, todos los números tienen el mismo peso. Es decir, los números se suman juntos y se dividen por la cantidad de valores. Con un promedio ponderado cada valor se le asigna una ponderación, por lo tanto algunos valores influyen más en el resultado que otros.

Dicho lo anterior, podemos decir que el promedio ponderado se utiliza cuando dentro de una serie de datos, uno de los valores tiene una mayor importancia o hay un dato con mayor peso que el resto. Además, nos ayuda a establecer dicho peso, a través del método conocido como ponderación y utilizar este valor para realizar el cálculo promedio.

La fórmula general es:

(valor1 * ponderación1 + valor2 * ponderación2 + ...)/(ponderación1 + ponderación2…)

La planilla de cálculo nos ofrece la posibilidad de calcular el promedio ponderado combinando dos funciones: la ya conocida SUMA() más SUMAPRODUCTO()

SUMAPRODUCTO() calcula la suma de las multiplicaciones de los datos correspondientes de dos intervalos del mismo tamaño. 

Para obtener el promedio ponderado, debemos hacer la

SUMAPRODUCTO(rango_valores; rango_ponderaciones)/SUMA(rango_ponderaciones)


# Funciones de Google Sheets

- IMPORTRANGE
- QUERY
- FILTER
- ARRAY FORMULA


GOAL SEEK
MACROS
PIVOT 

