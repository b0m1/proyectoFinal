# Proyecto final
## Curso de Excel para principiantes
#####  04/04 - 05/04
- **Trabajar con archivos en excel**: en este apartado se aprenden cosas muy básicas para manejarse con Excel (crear archivos, guardarlos, editarlos, etc). 
- **Formato de celdas**: explica los distintos formatos y estilos que le puedes aplicar a una celda. Existen distintos formatos que puedes aplicarle a números (moneda, contabilidad, fecha, porcentaje, científica, etc) o a textos (negrita, alineación,  color).
- **Realizar cálculos y funciones básicas**: las funciones se insertan con el símbolo _=_ seguido por el nombre de la función y paréntesis _()_ en los que insertamos los parámetros necesarios o, en el caso de que no tenga, los dejamos vacios. Puedes obtener una referencia absoluta de una celda pulsando F4, si pulsas dos veces sólo hace absoluta la referencia de la fila y, si lo haces una tercera vez, se vuelve absoluta la columna.
- **Gráficos**: para crear un gráfico tenemos que seleccionar los datos que queremos que lo compongan y seleccionar _insertar > gráficos recomendados_ y elegimos el que queramos. Si pulsamos en los distintos tipos tenemos una vista previa de cómo quedaría nuestro gráfico.

##### 06/04 - 09/04

- **Hojas de cálculo**: podemos crear más hojas de distintas formas. En la parte inferior pulsando el icono que tiene un + o seleccionando _inicio > insertar > insertar hoja_. Podemos copiar una hoja en el mismo libro si pulsamos CTRL y la arrastramos a la ubicación que deseemos. Otra opción es hacer click derecho en la pestaña y seleccionar _mover o copiar_.
- **Diseño de impresión**: seleccionamos las hojas de cálculo que vayamos a imprimir o guardar en pdf y seleccionamos _archivo > imprimir_.

[Ejercicios](ejercicios/Principiantes.xlsx)

> 09/04
> **Primer intento de exámen**: no superado :(

> 10/04
> **Segundo intento de exámen**: superado :)

![](celebration.gif)

## Curso de Excel: Funciones Básicas
#####  09/04 - 10/04
- **Funciones matemáticas**: explica qué hacen las funciones de este grupo más comunes como:  suma, sumar.si,  abs, redondear, etc. También los distintos tipos de funciones que existen: sin parámetros (pi), con múltiples parámetros (suma) y con parámetros opcionales (promedio) .
- **Funciones Estadísticas**: 
	- **contar**: cuenta las celdas que contienen números.
	- **contara**: cuenta las celdas que no están vacias. 
	- **max y min**: indica el valor numérico más grande/pequeño en un rango de celdas.
	- **mediana**: coge el número que parte a la mitad un conjunto de números que están ordenados de menor a mayor.
	- **moda**: es el número que más se repite en un grupo de números
	- **promedio**: suma un rango de números y lo divide por la cantidad de números que hay en ese rango.

##### 17/04
 >En la hoja **datos** usé la función importHTML para coger datos de lastfm, no es lo más útil pero era por probar. En la hoja **datos formateados** están los datos "importantes" formateados de forma que sean más fácil de leer. 
[Hoja de cálculo pa probar cosas](https://docs.google.com/spreadsheets/d/1hfxe_l0k1tU5pS1mD4TpWGmoE3gJaDTh5dBuOsL6dcE/edit?usp=sharing)

##### 18/04 
- **Funciones de fechas**: las fechas en excel se almacenan como números enteros secuenciales, empezando por el 01/01/1900; por lo tanto, se pueden realizar cálculos con fechas de la misma forma que con números enteros. Funciones útiles: 
	- **hoy(); ahora()** 
	- **dias.lab(fecha_inicial ; fecha_final ; _vacaciones_)**: devuelve el núm de días laborables que existen entre las fechas indicadas. No incluye los fines de semana ni las fechas indicadas en vacaciones.
	- **fecha()**: une tres valores distintos para convertirlos en una fecha.
	- **fecha.mes(fecha_inicial ; meses)**: calcula una fecha con el número de meses que le indiques, si el valor es negativo devolverá  una fecha pasada a fecha_inicial, si es positivo será una fecha futura.
- **Funciones de texto**: Se utiliza el operador & para concatenar cadenas. Podemos compararlas con el símbolo = (no distingue mayúsculas) o con la función igual (si distingue). Funciones útiles: 
	- **espacios()**: elimina espacios que haya al principio o final de la cadena.
	- **limpiar()**:  elimina carácteres no imprimibles.
	- **izquierda(),  derecha()**: devuelve la cantidad de carácteres que indiquemos empezando por la izq o derecha. 
	- **mayus(), minusc()**
	- **nompropio()**: convierte la primera letra de cada cadena que no sea un único carácter a mayúsculas y el resto de las letras a minúsculas.
	- **largo()**: obtiene el número de caracteres de una cadena de texto.

[Ejercicios](ejercicios/Funciones_basicas.xlsx)

 > 18/04
> **Primer intento de exámen**: superado :)

![](celebration.gif)

## Curso de Excel: Funciones Avanzadas
##### 20/04
- **Funciones Lógicas**: if, and, or.  
	- **Formato personalizado**: Para crear un formato de número personalizado tenemos que ir a la ventana de formato de celdas y en la pestaña de números elegir la categoría _personalizado_. Escogemos el tipo que más se parezca a lo que queramos y modificamos como mejor nos convenga.
	- **Formato condicional**: nos permite resaltar un rango de celdas según la condición que le apliquemos. Podemos cambiar el color de la fuente, del fondo de la celda, aplicar escalas de colores, agregar iconos y otras muchas cosas.
- **Funciones de búsqueda**: Funciones para buscar elementos dentro de tablas como Buscar.v y otras como índice, fila, columna para usar en conjunto con la anterior.
- **Funciones varias**:  
	- **Jerarquía.EQV**: devuelve la posición de un número dentro de una lista de números.
	- **Hipervínculo** para crear links, **FormulaTexto** para ver la fórmula que hay en una celda como una cadena.
	- **Pago**: una función financiera que calcula el pago periódico que debemos realizar por un préstamo.
	- **Convertir**: para convertir de un sistema de medida (peso, distancia, temperatura, energía, potencia, etc) a otro. Muy útil.
- **Novedades en Excel 365**: buscarx se puede usar sustituyendo a buscarv ya que es una versión mejorada de ésta. BuscarX puede devolver una matriz con varios elementos.
	- **Ordenar y ordenarpor**: devuelven una matriz ordenada de los elementos de una matriz. La diferencia es que ordenarpor() usamos los datos de otra matriz para ordenar la matriz. 
	- **Si.conjunto**: para no tener que anidar ifs. 
	- **Filtrar**:  permite filtrar un rango de datos según los criterios que definamos.
	- **Unicos**: devuelve los valores que no están repetidos de una lista. 
	- **Secuencia**: crear una lista de números secuenciales pudiendo definir en qué numero empieza y en cuánto se incrementa. 

[Aquí hay buscarv, si, si.error, hipervínculo, formato condicional y otras, he intentado pasarlo a excel pero es un despropósito porque hay funciones propias de sheets y voy a perder mucho tiempo en eso solo. Una pena porque los caracteres japoneses se ven mucho más chulos en excel.](https://docs.google.com/spreadsheets/d/14kz8qrnCIhoKCuyz1kxJKhWkl1hcu4ruL2s1hSQs0H0/edit?usp=sharing)

 > 24/04
> **Primer intento de exámen**: superado :)

![](celebration.gif)


## Curso de Excel: Bases de Datos
##### 25/04
- **Tablas**: Dar formato como tabla nos permite trabajar con los datos de manera más sencilla, por ejemplo: segmentarlos, seleccionarlos, validarlos o filtrarlos. Funciones de bases de  datos como: contar, sumar, max, min pero aplicadas a una base de datos. Fila de totales: crea una fila al final de la tabla en la que puedes aplicar una función para cada columna (suma, promedio...). Vistas de hoja
- **Tablas Dinámicas**: crearlas, ordenarlas, agrupar elementos, segmentarlas, lo visto en las tablas normales y algunas características propias que tienen las dinámicas. Crear campos y elementos calculados. Distintas formas de referenciar celdas de una tabla dinámica (sin import y con import). Gráficos dinámicos (en la ficha análisis de gráfico dinámico). Diseño de la tabla dinámica (pestaña diseño): activar subtotales, formato tabular... Aplicar formato condicional (se hace igual).
- **Importar y relacionar datos**: Hay que tener en cuenta la versión de excel. Datos --> Obtener datos. En excel no se pueden tener relaciones de varios a varios. Se pueden importar de bbdd de páginas web o de archivos de texto.

 > 27/04
> **Exámen**: superado :)

![](celebration.gif)

## Curso de Excel: Gráficos
##### 28/04
- **Minigráficos**, se crean en una celda.
- **Crear un gráfico**: a partir de unos datos, seleccionamos los que queremos que aparezcan y en insertar --> gráficos recomendados o elegimos el que queremos
##### 02/05
- **Gráficos en Excel**
	- **Nuevos tipos de gráficos en excel 365**: mapa (para representar los datos geográficos en un mapa), rectángulos, proyección solar (para cuando existen varias categorias), cajas y bigotes, cascada y embudo.
	- **Modificar un gráfico**. Hay 3 botones en la parte de la derecha: el primero para añadir/quitar elementos del gráfico, el segundo (pincel) para cambiar el estilo del gráfico y el embudo para filtrar información. En la pestaña diseño de gráfico > seleccionar datos podemos cambiar el rango de datos. Para cambiar las series hay que hacer doble click sobre ellas. Podemos modificar el eje y de la misma forma.
	- **Personalizar un gráfico**: en la pestaña diseño de gráfico > agregar elemento de gráfico o el primer botón que hay a la derecha del gráfico. Podemos poner títulos,  lineas de cuadrícula, tabla de datos, etc.
	- **Herramientas de análisis**: en el mismo apartado anterior tenemos los elementos barras de error o línea de tendencia.
	- **Gráficos de sectores**: sólo se puede ver un rango de datos (en el embudo podemos decidir qué datos ver). Tienen una versión con subgráfico de barras que es muy útil para representar datos no proporcionales.
	- **Crear plantillas de gráficos**: haciendo click derecho en un gráfico le damos a guardar como plantilla. Accedemos a ellas desde insertar gráfico en el apartado "plantillas".
- **Imágenes en Excel**.
	- **Capturas de pantalla**: en Insertar > Ilustraciones > Captura.
	- **Insertar imágenes**: en Insertar podemos usar bing desde excel para buscar cualquier imagen. Puedes editar la imagen en la pestaña formato de la imagen. Es importante comprimirlas si vamos a tener muchas.
	- **SmartArts**: otro tipo de gráficos, por ejemplo tenemos la opción de insertar organigramas o diagramas.

[Aquí hay gráficos](https://docs.google.com/spreadsheets/d/1n9YLOcUbrjLEQJuI_hvWsfR8t0vPq5pOpIJQyVew7b0/edit?usp=sharing)

 > 04/05
> **Exámen**: Aprobado

![](celebration.gif)

## Curso de Excel: Macros
##### 04/05
- **Grabar una macro**: en la pestaña de Programador/Desarrollador le damos a "Grabar macro". Hay que ponerle nombre (sin espacios ni puntos), una combinación de teclas para ejecutarla (opcional), especificar dónde la vamos a guardar (para que funcione tiene que estar abierto el libro en el que esté guardada) y ponerle una descripción. Importante: detener la grabación. No se puede usar el deshacer.
- _**Ejercicio**_: una macro que ponga el punto a los miles sin usar el botón por defecto.
- **Formas de ejecutar una macro**: con las teclas asignadas, desde el botón de macros o creando un botón dentro de la hoja y asignándole la macro. También podemos añadir un botón a la barra de herramientas de la macro que queramos y personalizarlo.
- **Seguridad de macros**: hay varias opciones para limitar las macros que puedes utilizar. Puedes dar permiso a una ubicación para que permita usar todas las macros que tengas ahí.
- **Formato de libro con macros**: los archivos de excel normales no permiten guardar macros, hay que guardarlo en formato .xlsm o .xlsb.
- **Crear un complemento**: guardar con formato .xlam y habilitarlo en Programador > Complementos.
- **Referencias absolutas y relativas**:
	- _**Ejercicio**_: con la plantilla que nos da tenemos que crear una macro con referencias absolutas.
- **Hasta donde llegan las macros**: Para poder usar bien las macros hay que saber VBA.

> 05/05
> **Exámen**: Aprobado

![](wiii.gif)

## Curso de Excel: Herramientas de cálculo avanzado
##### 05/05 
- **Búsqueda de objetivos**: Sirve para resolver una ecuación. En la ficha de Datos/Análisis de hipótesis/Buscar objetivo. Ahí indicamos la celda donde queremos coger los datos, el valor objetivo y la celda que queremos cambiar. 
- **Tablas de datos**: Selccionamos los datos y en la ficha de Datos/Análisis de hipótesis/Tabla de datos. Hay que indicar la celda de la que coge los daos la fila o la columna o ambas.
- **Escenarios**: Nos permite analizar nuestros datos creando varios escenarios modificando en cada caso el dato que le indiquemos, de esta forma podemos ver como cambiarían las fórmulas o los otros datos que dependan del que modifiquemos con el escenario. Los podemos crear en Datos/Analisis de hipótesis/Administrador de Escenarios.
- **Consolidar**:
- **Solver**: Hay que instalarla en archivo/opciones/complementos
- **Nuevas herramientas Excel 365**:
- **Protección de datos**:

## Curso de Excel: Powers
#### Power Query
##### 05/05 
- **Importar datos con Power Query**. 
##### 08/05
- **Importar archivos de texto**: hay dos formatos: archivos de texto delimitados (txt) y archivos de texto con valores separados por comas (csv).
- **Importar datos desde una página web**: Obtener datos/Desde otras fuentes/Desde la web. Tiene que estar en formato de tablas.
- **Importar datos desde cualquier base de datos**:Obtener datos/Desde otras fuentes/Desde microsoft query.
- **Pantalla inicial de Power Query**: En datos/consultas y conexiones podemos acceder a power query dándole a editar a cualquiera de nuestras consultas. En el lateral derecho podemos cambiarle el nombre y ver los cambios que hemos realizado. Para deshacer una acción en power query tenemos que utilizar el formulario de pasos aplicados. 
##### 11/05
- **Limpieza de Datos**: Antes de hacerlo es recomendable hacer una copia de seguridad. En la ficha de _inicio > administrar columnas > quitar columnas_ tenemos dos opciones: quitar columnas y quitar otras columnas. La primera quita las que tengamos seleccionadas y la segunda quita todas menos las que tengamos seleccionadas. También en _inicio > reducir filas_ podemos elegir cómo modificar las filas. Tenemos la opción de quitar las filas superiores, inferiores, alternas, duplicados, en blanco y errores. 
 **Combinar consultas**: Si tenemos tablas relacionadas entre sí podemos ir a la columna en la que se encuentra el dato relacionado y en el botón que tiene en el encabezado tenemos la opción de expandir. Lo que hace es insertar en esa tabla todas las columnas que seleccionemos de la tabla relacionada.
- **Agrupar**: Esta opción se encuentra en la ficha de Inicio. Tiene dos opciones: básico (sólo podemos elegir un campo) y avanzado (podemos elegir varios). Tenemos que incidar el campo por el que queremos agrupar (o campos si elegimos el uso avanzado). Además podemos elegir qué operaciones queremos que haga.
- **Unir y separar campos**: Tanto en la ficha _Transformar_ como en la ficha _Agregar Columna_ tenemos la opción _Combinar columnas_. En transformar cambiará la columna en la que estamos mientras que de la otra forma añadirá una columna nueva con los datos combinados. Podemos indicarle qué separador poner. A la hora de separar los datos tenemos muchas más opciones (por delimitador, por número de caracteres, por posiciones).