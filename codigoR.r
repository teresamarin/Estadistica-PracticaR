library(openxlsx)
library(readxl)
library(e1071)  # para curtosis y asimetría
library(ggplot2)
library(dplyr)
library(grid)
library(moments)
library(scales)

getwd()
setwd("") # aquí debe de poner cada uno su ruta a los archivos

# Crear el workbook
wb <- createWorkbook()

# -----------------------------
# -----------------------------
# HOJA 1:
addWorksheet(wb, "Hoja1")
# Leer los datos desde la hoja
datos <- read.xlsx("datos1.xlsx", sheet = 1)
# -----------------------------
# PRIMERA COLUMNA
# Escribir encabezado "X" en A1
writeData(wb, "Hoja1", as.character(names(datos)[1]), startCol = 1, startRow = 1)

# Estilo para línea gruesa arriba de la fila de totales
estilo_linea_gruesa <- createStyle(border = "bottom", borderStyle = "thin")

# Aplicarlo al inicio de la fila de totales
addStyle(wb, "Hoja1",
         style = estilo_linea_gruesa,
         rows = 1,
         cols = 1,
         gridExpand = TRUE)

# Escribir los datos reales en A2 en adelante
writeData(wb, "Hoja1", datos$X, startCol = 1, startRow = 2, colNames = FALSE)

# Texto del comentario (formateado con saltos de línea)
comentario_variable <- paste(
  "X = idioma medido",
  "sobre n = 40 libros",
  "extraídos al azar en una determinada biblioteca",
  "",
  "A = Alemán",
  "E = Español",
  "F = Francés",
  "I = Inglés",
  "R = Ruso",
  sep = "\n"
)

# Crear e insertar comentario
comentario <- createComment(
  comment = comentario_variable,
  author = "Definición",
  visible = FALSE,
  width = 6, height = 5
)

# Añadir el comentario a la primera celda
writeComment(wb, "Hoja1", col = 1, row = 1, comment = comentario)


# -----------------------------
# Preparación para las gráficas:
tabla_frec <- as.data.frame(table(datos$X)) # para las tablas de frecuencias
tabla_frec
#   Var1 Freq
# 1    A    5
# 2    E    9
# 3    F   11
# 4    I   13
# 5    R    2
colnames(tabla_frec) <- c("X", "f.a.")
tabla_frec
#   X f.a.
# 1 A    5
# 2 E    9
# 3 F   11
# 4 I   13
# 5 R    2

# -----------------------------
# DIAGRAMA DE BARRAS
png("barras.png", width = 800, height = 500)

# Creamos un espacio base y calculamos posiciones
pos <- barplot(tabla_frec$`f.a.`,
               plot = FALSE,
               width = 0.3,
               space = 1)

xlim <- range(pos) + c(-0.5, 0.5)
ylim <- c(0, 14)

# Preparamos el "lienzo"
par(mar = c(3, 3, 4, 5.5)) # medidas exactas para encuadrar el gráfico en la imagen
# en esta función par(mar = c(inferior, izquierda, superior, derecha))
plot.new()
plot.window(xlim = xlim, ylim = ylim)

# Línea negra del eje X completa, llegando a los límites de la gráfica
segments(x0 = par("usr")[1], # con esta
         y0 = 0,
         x1 = par("usr")[2], # y esta línea llegamos a los límites del eje
         y1 = 0,
         col = "black")

# Para poner las muescas entre las barras verticales de frecuencias:
pos_muescas <- head(pos, -1) + diff(pos)/2 # calculamos las posiciones en las que estarían
segments(x0 = pos_muescas, y0 = 0, x1 = pos_muescas, y1 = -0.3, col = "black") # y las pintamos

# Pintamos las líneas horizontales antes de las barras de frecuencia, para que queden por debajo:
abline(h = seq(2, 14, by = 2), col = "gray", lty = 1)

# Ahora pintamos las barras de frecuencias por encima:
barplot(tabla_frec$`f.a.`,
        width = 0.3,
        space = 1,
        col = "steelblue",
        border = "black",
        add = TRUE,
        axes = FALSE,
        names.arg = FALSE)

# Ejes y sus etiquetas:
mtext(text = tabla_frec$X, side = 1, at = pos, line = 0.2, cex = 1.2)
axis(2, at = seq(0, 14, by = 2), las = 1, cex.axis = 1.1)

# Título centrado en la imagen (si fuese centrado en el gráfico sería con "title"):
mtext("Diagrama de Barras", side = 3, line = 1, cex = 2, font = 2)

# Leyenda:
legend(x = par("usr")[2] + 0.1,  # a la derecha del área del gráfico
       y = mean(par("usr")[3:4]),  # centrado verticalmente
       legend = "f.a.",
       fill = "steelblue",
       bty = "n", cex = 1.2,
       xpd = TRUE)

dev.off()

# Insertar gráfico de barras
insertImage(wb, "Hoja1", "barras.png",
            startCol = 9,
            startRow = 1,
            width = 6.5,
            height = 3.8,
            units = "in")

# -----------------------------
# TABLA DE FRECUENCIAS
# Calcular frecuencias relativas y añadirlas a la tabla que estamos creando de frecuencias
tabla_frec$`f.r.` <- round(tabla_frec$`f.a.` / sum(tabla_frec$`f.a.`), 3)
tabla_frec$`f.r.%` <- round(tabla_frec$`f.r.` * 100, 1)
nombres_columnas <- colnames(tabla_frec) # copiamos los nombres de las columnas

# Agregar fila de totales
fila_total <- as.data.frame(matrix(NA, nrow = 1, ncol = length(nombres_columnas)))
colnames(fila_total) <- nombres_columnas # para asegurarnos de que la tabla se pueda juntar bien
fila_total <- setNames(data.frame("",
                                  "", # si quisiésemos poner el total de la col: sum(tabla_frec$`f.a.`)
                                  round(sum(tabla_frec$`f.r.`), 3),
                                  round(sum(tabla_frec$`f.r.%`), 1)
                                  ), nombres_columnas)
tabla_completa <- rbind(tabla_frec, fila_total)
# La tabla en su estado final es así:
tabla_completa
#   X  f.a.    f.r.   f.r.%
# 1 A     5   0.125    12.5
# 2 E     9   0.225    22.5
# 3 F    11   0.275    27.5
# 4 I    13   0.325    32.5
# 5 R     2   0.050     5.0
# 6           1.000   100.0


# Moda
(moda_valor <- tabla_frec$X[which.max(tabla_frec$`f.a.`)])
# [1] I



# Estilos para escribir la tabla en el excel
estilo_centrado <- createStyle(halign = "center", valign = "center")
estilo_negrita <- createStyle(textDecoration = "bold")

# Insertar la tabla de frecuencias en C2
start_col <- 3 # C
start_row <- 2 # 2
colnames(tabla_completa)[1] <- "" # Con esto hacemos que ya no aparezca el nombre de la variable "X"
# Ahora la tabla se ve así:
tabla_completa
#     f.a.  f.r. f.r.%
# 1 A    5 0.125  12.5
# 2 E    9 0.225  22.5
# 3 F   11 0.275  27.5
# 4 I   13 0.325  32.5
# 5 R    2 0.050   5.0
# 6        1.000 100.0

# Escribimos la tabla en este último estado en el excel:
writeData(wb, "Hoja1", tabla_completa, startCol = start_col, startRow = start_row, headerStyle = estilo_centrado)

# Estilo para línea gruesa arriba de la fila de totales
estilo_linea_fina_top <- createStyle(border = "top", borderStyle = "thin")
estilo_linea_fina_bottom <- createStyle(border = "bottom", borderStyle = "thin")
estilo_linea_fina_left <- createStyle(border = "left", borderStyle = "thin")
estilo_linea_fina_right <- createStyle(border = "right", borderStyle = "thin")

# Aplicarlo a las zonas necesarias para rellenar la tabla
addStyle(wb, "Hoja1",
         style = estilo_linea_fina_bottom,
         rows = start_row:(start_row + nrow(tabla_completa)-1),
         cols = start_col:(start_col + 3),
         gridExpand = TRUE, stack = TRUE)
addStyle(wb, "Hoja1",
         style = estilo_linea_fina_right,
         rows = start_row:(start_row + nrow(tabla_completa)),
         cols = start_col:(start_col + 2),
         gridExpand = TRUE, stack = TRUE)

# Insertar moda en el excel:
writeData(wb, "Hoja1", "Moda", startCol = start_col, startRow = start_row + 8)
writeData(wb, "Hoja1", moda_valor, startCol = start_col + 2, startRow = start_row + 8)

# Explicaciones:
writeData(wb, "Hoja1", "FRECUENCIA ABSOLUTA", startCol = start_col+1, startRow = start_row + 10)
writeData(wb, "Hoja1", "FRECUENCIA RELATIVA", startCol = start_col+1, startRow = start_row + 11)
writeData(wb, "Hoja1", "fa/tm (tamaño muestral)", startCol = start_col + 4, startRow = start_row + 11)

# -----------------------------
# DIAGRAMA DE SECTORES
# Calcular proporciones
proporciones <- round(tabla_frec$`f.a.` / sum(tabla_frec$`f.a.`), 3)

# Paleta apagada
colores_excel <- c("#4F81BD", "#C0504D", "#9BBB59", "#8064A2", "#4BACC6")  # A, E, F, I, R

# Crear gráfico (no compilar para ver gráfico en ventana "plots")
png("sectores.png", width = 800, height = 500)

par(mar = c(0.5, 1, 4, 1))  # márgenes a nuestro gusto

# Diagrama de sectores invertido, con título:
pie(rev(proporciones),                         # invertir orden
    labels = rev(proporciones),                # invertir etiquetas también 
    col = rev(colores_excel),                  # invertir colores en el mismo orden
    border = NA,                               # sin bordes
    main = "Diagrama de Sectores",             # título
    clockwise = FALSE,                         # antihorario
    radius = .9,                               # radio del círculo
    init.angle = 90,                           # para que empiece a pintar desde arriba
    cex.main = 3)                              # tamaño del título

# Leyenda:
legend(x = 1.3, y = 0.3,                       # posición en coord del gráfico
       legend = tabla_frec$X,                  # texto de la leyenda
       fill = colores_excel,                   # colores a utilizar
       bty = "n",                              # sin borde ("n" = "none")
       cex = 1.2,                              # tmñ del texto en la leyenda
       xpd = TRUE)                             # se podría llegar a pintar la leyenda fuera del gráfico

dev.off()

# Insertamos la leyenda en la hoja del excel:
insertImage(wb, "Hoja1", "sectores.png", 
            startCol = 9, 
            startRow = 20,
            width = 6.5,
            height = 3.8,
            units = "in")

# -----------------------------
# Como ya terminamos la Hoja1, vamos a borrar las variables, dataframes, listas 
# y objetos del entorno (environment) para que se pueda controlar todo mejor 
# (salvo el Workbook, que lo estamos usando en estos momenos):
rm(list = setdiff(ls(), "wb"))
ls()


# -----------------------------
# -----------------------------
# HOJA 2:
addWorksheet(wb, "Hoja2")
# Leemos datos desde "datos1.xlsx", hoja 2
datos <- read.xlsx("datos1.xlsx", sheet = 2, skipEmptyRows = TRUE, skipEmptyCols = TRUE)

# -----------------------------
# PRIMERAS TRES COLUMNAS
columna <- names(datos)[1]  # nombre de la variable: "X"
datos$X <- as.numeric(datos[[1]])

# Calculamos media
media_x <- mean(datos$X)

# Calcular (X - media)^2 y X²
datos$`(X - media)^2` <- round((datos$X - media_x)^2, 6)
datos$`X^2` <- datos$X^2

# Escribir nombre de variable en A1
writeData(wb, "Hoja2", columna, startCol = 1, startRow = 1)

# Añadir comentario manual a A1
comentario <- "X = nº de palabras por línea medido sobre n=48 líneas elegidas al azar en un libro"
comentario_obj <- createComment(comment = comentario, author = "Definición",
                                visible = FALSE, width = 6, height = 4)
writeComment(wb, "Hoja2", col = 1, row = 1, comment = comentario_obj)

# Escribir solo los datos (sin encabezados)
writeData(wb, "Hoja2", datos, startCol = 1, startRow = 2, colNames = FALSE)

# -----------------------------
# TABLAS, hj2
# -----------------------------
# TABLA 1 (la de debajo del todo)
# Nos guardamos los datos en número en la var. x:
x <- as.numeric(datos[[1]])

# Calculamos estadísticos para tabla:
moda <- as.numeric(names(sort(table(x), decreasing = TRUE))[1])
media <- mean(x)
mediana <- median(x)
q1 <- quantile(x, 0.25) # primer cuartil
q3 <- quantile(x, 0.75) # tercer cuartil
rango <- max(x) - min(x)
iqr <- IQR(x)
var_pob <- var(x) * (length(x) - 1) / length(x)
desv_tipica <- sqrt(var_pob)

error_tipico <- sd(x) / sqrt(length(x))
desv_muestral <- sd(x)
var_muestral <- var(x)

# creamos una función para calcular la asimetría como en el excel
asimetria_excel <- function(x) {
  n <- length(x)
  m <- mean(x)
  s <- sd(x)
  sum(((x - m)/s)^3) * n / ((n - 1)*(n - 2))
}

# creamos otra función para calcular la curtosis con en el excel
curtosis_excel <- function(x) {
  n <- length(x)
  m <- mean(x)
  s <- sd(x)
  a <- sum(((x - m)/s)^4)
  b <- n*(n+1)/((n-1)*(n-2)*(n-3))
  c <- 3*(n-1)^2/((n-2)*(n-3))
  return(b*a - c)
}

curtosis <- curtosis_excel(x)
asimetria <- asimetria_excel(x)

minimo <- min(x)
maximo <- max(x)
suma <- sum(x)
cuenta <- length(x)

# Primer bloque (desde fila 54, primera tabla)
nombres1 <- c("Moda", "Media", "Mediana", "Primer cuartil", "Tercer cuartil",
              "Rango", "IQR", "Varianza", "Desviación típica")
valores1 <- c(moda, media, mediana, q1, q3, rango, iqr, var_pob, desv_tipica)
(tabla1 <- data.frame(
  Variable = nombres1,
  Valor = valores1))
# La tabla queda de la siguiente manera:
#            Variable     Valor
# 1              Moda 13.000000
# 2             Media 11.333333
# 3           Mediana 12.000000
# 4     Primer cuarti 10.750000
# 5     Tercer cuarti 13.000000
# 6             Rango 12.000000
# 7               IQR  2.250000
# 8          Varianza  6.597222
# 9 Desviación típica  2.568506
writeData(wb, "Hoja2", tabla1, startCol = 1, startRow = 55, colNames = FALSE)

# Segundo bloque (desde fila 68, segunda tabla)
nombres2 <- c("Media", "Error típico", "Moda", "Mediana", "Desviación estándar",
              "Varianza de la muestra", "Curtosis", "Coeficiente de asimetría",
              "Rango", "Mínimo", "Máximo", "Suma", "Cuenta")
valores2 <- c(media, error_tipico, moda, mediana, desv_muestral,
              var_muestral, curtosis, asimetria, rango, minimo, maximo, suma, cuenta)
(tabla2 <- data.frame(
  Variable = nombres2, 
  Valor = valores2))
# La tabla queda:
#                    Variable       Valor
# 1                     Media  11.3333333
# 2              Error típico   0.3746551
# 3                      Moda  13.0000000
# 4                   Mediana  12.0000000
# 5       Desviación estándar   2.5956865
# 6    Varianza de la muestra   6.7375887
# 7                  Curtosis   2.5253533
# 8  Coeficiente de asimetría  -1.5417234
# 9                     Rango  12.0000000
# 10                   Mínimo   3.0000000
# 11                   Máximo  15.0000000
# 12                     Suma 544.0000000
# 13                   Cuenta  48.0000000
writeData(wb, "Hoja2", tabla2, startCol = 1, startRow = 69, colNames = FALSE)


# Fusionar celdas A68 y A69
mergeCells(wb, sheet = "Hoja2", cols = 1:2, rows = 68)

# Escribir el texto de la celda 68
writeData(wb, sheet = "Hoja2", x = "Columna1", startCol = 1, startRow = 68, colNames = FALSE)

# Crear estilo con bordes gruesos y centrado
estilo_top <- createStyle(
  border = "Top",
  borderStyle = "medium",
  halign = "center",
  valign = "center",
  textDecoration = "italic"
)

# Estilo para borde inferior medio
estilo_bottom <- createStyle(
  border = "bottom",
  borderStyle = "thin"
)

# Aplicar estilos a la celda fusionada
addStyle(wb, sheet = "Hoja2", style = estilo_top,
         rows = 68, cols = 1:2, gridExpand = TRUE, stack = TRUE)
addStyle(wb, sheet = "Hoja2", style = estilo_bottom,
         rows = 68, cols = 1:2, gridExpand = TRUE, stack = TRUE)

# Estilo para borde inferior tabla
estilo_final_tabla <- createStyle(
  border = "bottom",
  borderStyle = "medium"
)
# Aplicar estilo inferior tabla
addStyle(wb, sheet = "Hoja2", style = estilo_final_tabla,
         rows = 81, cols = 1:2, gridExpand = TRUE, stack = TRUE)

# -----------------------------
# TABLA 2 (la que separa en intervalos, la de la mitad de la H2 datos1soluciones.xlsx)
# Definir clases manualmente
clases <- (min(x)-1):max(x) # creamos los datos desde por debajo del mínimo al máximo
intervalos <- cut(x, # creamos los intervalos sobre los datos del excel
                  breaks = clases, # creamos los cortes del intervalo según "clases"
                  right = TRUE) # el num. dcho está incluido (izdo excluido)

# Calculamos frecuencias
(f_a <- as.numeric(table(intervalos))) # calcula las f.a. de x según los intervalos que creamos antes
  # [1]  1  1  1  0  2  0  3  4  7 10 14  3  2
(f_r <- f_a / sum(f_a))
  # [1] 0.02083333 0.02083333 0.02083333 0.00000000 0.04166667
  # [6] 0.00000000 0.06250000 0.08333333 0.14583333 0.20833333
  # [11] 0.29166667 0.06250000 0.04166667
(F_a <- cumsum(f_a))
  #  [1]  1  2  3  3  5  5  8 12 19 29 43 46 48
(F_r <- cumsum(f_r))
  # [1] 0.02083333 0.04166667 0.06250000 0.06250000 0.10416667
  # [6] 0.10416667 0.16666667 0.25000000 0.39583333 0.60416667
  # [11] 0.89583333 0.95833333 1.00000000

# Puntos medios de clase (para normal teórica)
representantes <- min(x):max(x) # tomamos el lado dcho del intervalo como pto referencia


#### Función de densidad normal teórica
media <- mean(x)
varianza <- var(x) * (length(x) - 1) / length(x)  # varianza poblacional
desv_tipica <- sqrt(varianza)

# creamos una función para calcular la normal como en excel
normal_excel <- function(x, media, varianza, desv_tipica) {
  exp(-((x - media)^2) / (2 * varianza)) / (desv_tipica * sqrt(2 * pi))
}

# Calculamos la columna "Normal"
normal_teorica <- normal_excel(representantes, media, varianza, desv_tipica)


# Crear data.frame final (sin totales)
tabla_intervalos_hj2 <- data.frame(
  vacio = paste0("(", head(clases, -1), ",", tail(clases, -1), "]"),
  clases = representantes,
  `f.a.` = f_a,
  `f.r.` = round(f_r, 6),
  `f.a.a.` = F_a,
  `f.r.a.` = round(F_r, 6),
  Normal = round(normal_teorica, 14),
  stringsAsFactors = FALSE
)

colnames(tabla_intervalos_hj2)[1] <- "" # para que la primera celda aparezca vacía

# Crear una fila vacía con los mismos nombres
fila_total <- as.data.frame(matrix(NA, nrow = 1, ncol = ncol(tabla_intervalos_hj2)))
colnames(fila_total) <- colnames(tabla_intervalos_hj2)

# Insertar totales en las columnas deseadas
fila_total[1, "f.a."] <- sum(tabla_intervalos_hj2[["f.a."]])
fila_total[1, "f.r."] <- sum(tabla_intervalos_hj2[["f.r."]])

# Unir tabla y fila total
tabla_final <- rbind(tabla_intervalos_hj2, fila_total)

# Convertir todos los elementos a caracteres para evitar NAs
tabla_final[] <- lapply(tabla_final, function(x) {
  # Si el valor es NA, lo reemplazamos por ""
  x[is.na(x)] <- ""
  return(as.character(x))
})

tabla_final
# que nos queda:
#            clases f.a.     f.r. f.a.a.   f.r.a.       Normal
# 1    (2,3]      3    1 0.020833      1 0.020833 0.0008043945
# 2    (3,4]      4    1 0.020833      2 0.041667 0.0026371609
# 3    (4,5]      5    1 0.020833      3 0.062500 0.0074297514
# 4    (5,6]      6    0 0.000000      3 0.062500 0.0179879651
# 5    (6,7]      7    2 0.041667      5 0.104167 0.0374248345
# 6    (7,8]      8    0 0.000000      5 0.104167 0.0669125900
# 7    (8,9]      9    3 0.062500      8 0.166667 0.1028077587
# 8   (9,10]     10    4 0.083333     12 0.250000 0.1357419370
# 9  (10,11]     11    7 0.145833     19 0.395833 0.1540182883
# 10 (11,12]     12   10 0.208333     29 0.604167 0.1501760380
# 11 (12,13]     13   14 0.291667     43 0.895833 0.1258343158
# 12 (13,14]     14    3 0.062500     46 0.958333 0.0906082288
# 13 (14,15]     15    2 0.041667     48 1.000000 0.0560668672
# 14                  48 0.999999

# Creamos y aplicamos el estilo para la tabla:
estilo <- createStyle(
  border = "TopBottomLeftRight",
  borderStyle = "thin"
)
addStyle(wb, sheet = "Hoja2", style = estilo,
         rows = 34:47, cols = 5:10, gridExpand = TRUE, stack = TRUE)

# Escribir en Hoja2, a partir de fila 35 y columna 1 (A35)
writeData(wb, "Hoja2", tabla_final, startCol = 5, startRow = 34, colNames = TRUE)


# -----------------------------
# TABLA 3 (la segunda de la segunda columna de tablas)
# Todos los valores que queremos que aparezcan
valores_completos <- min(x):max(x)

# Crear tabla incluyendo todos los valores, aunque no aparezcan
tabla_valores <- as.data.frame(table(factor(x, levels = valores_completos)))

# Renombrar columnas
colnames(tabla_valores) <- c("Clase", "Frecuencia")

# Calcular % acumulado
tabla_valores$`% acumulado` <- cumsum(tabla_valores$Frecuencia) / sum(tabla_valores$Frecuencia)

# Convertir columna "Clase" a texto antes de cambiar el último valor
tabla_valores$Clase <- as.character(tabla_valores$Clase)

# Reemplazar el último valor por "y mayor..."
tabla_valores$Clase[nrow(tabla_valores)] <- "y mayor..."

# Escribir en el excel y poner estilos adecuados
writeData(wb, "Hoja2", tabla_valores, startCol = 8, startRow = 18, colNames = TRUE)

estilo_top <- createStyle(
  border = "Top",
  borderStyle = "medium",
  halign = "center",
  valign = "center",
  textDecoration = "italic"
)

estilo_bottom <- createStyle(
  border = "bottom",
  borderStyle = "thin"
)

addStyle(wb, sheet = "Hoja2", style = estilo_top,
         rows = 18, cols = 8:10, gridExpand = TRUE, stack = TRUE)
addStyle(wb, sheet = "Hoja2", style = estilo_bottom,
         rows = 18, cols = 8:10, gridExpand = TRUE, stack = TRUE)

estilo_bottom <- createStyle(
  border = "bottom",
  borderStyle = "medium"
)
addStyle(wb, sheet = "Hoja2", style = estilo_top,
         rows = 32, cols = 8:10, gridExpand = TRUE, stack = TRUE)

# -----------------------------
# TABLA 4 (la que tiene un filtro arriba)
# Crear tabla de frecuencias
tabla_frec <- as.data.frame(table(x))
colnames(tabla_frec) <- c("X", "f.a.")
tabla_frec$X <- as.numeric(as.character(tabla_frec$X))

# Actualmente, tabla_frec tiene:
tabla_frec
#     X f.a.
# 1   3    1
# 2   4    1
# 3   5    1
# 4   7    2
# 5   9    3
# 6  10    4
# 7  11    7
# 8  12   10
# 9  13   14
# 10 14    3
# 11 15    2

# Calcular frecuencia relativa (numérica y de texto)
(f_total <- sum(tabla_frec$`f.a.`))
  # [1] 48
(f_r <- tabla_frec$`f.a.` / f_total)
  # [1] 0.02083333 0.02083333 0.02083333 0.04166667 0.06250000
  # [6] 0.08333333 0.14583333 0.20833333 0.29166667 0.06250000
  # [11] 0.04166667
(f_r_text <- paste0(round(f_r * 100, 2), "%"))
  #  [1] "2.08%"  "2.08%"  "2.08%"  "4.17%"  "6.25%"  "8.33%"  "14.58%"
  # [8] "20.83%" "29.17%" "6.25%"  "4.17%" 

# Calcular Suma - X
suma_x <- tabla_frec$X * tabla_frec$`f.a.`

# Calcular (X - media)^2 y X²
col_num_media <- round((tabla_frec$X - media)^2, 6)
col_cuadrado <- tabla_frec$X^2

# Crear tabla final explícita
tabla_filtro <- data.frame(
  X = c(tabla_frec$X, "Total Resultado"),
  `f.r.` = c(f_r_text, "100.00%"),
  `Suma - X` = c(suma_x, sum(suma_x)),
  `X - media²` = c(col_num_media, ""),
  `X²` = c(col_cuadrado, ""),
  stringsAsFactors = FALSE
)
# ponemos los dos últimos nombres vacíos para coincidir con el excel
colnames(tabla_filtro)[4:5] <- c("", "")

tabla_filtro
# y así nos queda finalmente la tabla:
#                  X    f.r. Suma...X              
# 1                3   2.08%        3 69.444444   9
# 2                4   2.08%        4 53.777778  16
# 3                5   2.08%        5 40.111111  25
# 4                7   4.17%       14 18.777778  49
# 5                9   6.25%       27  5.444444  81
# 6               10   8.33%       40  1.777778 100
# 7               11  14.58%       77  0.111111 121
# 8               12  20.83%      120  0.444444 144
# 9               13  29.17%      182  2.777778 169
# 10              14   6.25%       42  7.111111 196
# 11              15   4.17%       30 13.444444 225
# 12 Total Resultado 100.00%      544  





# Escribir tabla y aplicar filtros
writeData(wb, sheet = "Hoja2", "Datos", startCol = 7, startRow = 1)
writeData(wb, sheet = "Hoja2", x = tabla_filtro, startCol = 6, startRow = 2, withFilter = TRUE)

# Estilo en negrita
estilo_negrita <- createStyle(textDecoration = "bold")
addStyle(wb, "Hoja2", style = estilo_negrita,
         rows = nrow(tabla_filtro) + 2, cols = 6:8, gridExpand = TRUE, stack = TRUE)

# Crea estilo de fondo negro y texto blanco
estilo_negro <- createStyle(
  fgFill = "black",       # Fondo negro
  fontColour = "white",   # Texto blanco para que se lea
  #halign = "CENTER",      # Alineación horizontal opcional
  #valign = "CENTER"       # Alineación vertical opcional
)

# Aplicar estilos
addStyle(wb, sheet = "Hoja2", style = estilo_negro, rows = 1, cols = 7, gridExpand = FALSE)
addStyle(wb, sheet = "Hoja2", style = estilo_negro, rows = 2, cols = 6, gridExpand = FALSE)
addStyle(wb, sheet = "Hoja2", 
         style = createStyle(border = "top", borderStyle = "thick"),
         rows = 1, cols = 6:8, stack = TRUE)
addStyle(wb, sheet = "Hoja2",
         style = createStyle(border = "left", borderStyle = "thick"), 
         rows = 1:(nrow(tabla_filtro) + 2), cols = 6, stack = TRUE)
addStyle(wb, sheet = "Hoja2", 
         style = createStyle(border = "right", borderStyle = "thick"), 
         rows = 1:(nrow(tabla_filtro) + 2), cols = 8, stack = TRUE)
addStyle(wb, sheet = "Hoja2", 
         style = createStyle(border = "bottom", borderStyle = "thick"), 
         rows = nrow(tabla_filtro) + 2, cols = 6:8, stack = TRUE)
addStyle(wb, sheet = "Hoja2", 
         style = createStyle(border = "bottom", borderStyle = "medium"), 
         rows = 1, cols = 6:8, stack = TRUE)
addStyle(wb, sheet = "Hoja2", 
         style = createStyle(border = "bottom", borderStyle = "medium"), 
         rows = 2, cols = 6:8, stack = TRUE)
addStyle(wb, sheet = "Hoja2", 
         style = createStyle(border = "bottom", borderStyle = "medium"), 
         rows = nrow(tabla_filtro) + 1, cols = 6:8, stack = TRUE)
addStyle(wb, sheet = "Hoja2", 
         style = createStyle(border = "right", borderStyle = "medium"), 
         rows = 1:(nrow(tabla_filtro) + 2), cols = 6, stack = TRUE)

# -----------------------------
# GRAFICOS, hj2
# -----------------------------
# PREPARACIÓN PREVIA:
x <- datos[[1]]

# Tabla de frecuencias
valores_completos <- min(x):max(x)
tabla_valores <- as.data.frame(table(factor(x, levels = valores_completos)))
colnames(tabla_valores) <- c("Clase", "Frecuencia")

# Cálculo del porcentaje acumulado
tabla_valores$`pct_acumulado` <- cumsum(tabla_valores$Frecuencia) / sum(tabla_valores$Frecuencia)

# crear una columna numérica para posicionar las barras
tabla_valores$pos_x <- seq_along(tabla_valores$Clase)
# crear las etiquetas personalizadas
etiquetas_x <- as.character(tabla_valores$Clase)
etiquetas_x[length(etiquetas_x)] <- "y mayor..."

# -----------------------------
# Gráfico 1: Barras + línea de acumulado
png("grafico1_frecuencia_acumulada.png", width = 800, height = 600)
ggplot(tabla_valores, aes(x = pos_x)) + # Iniciamos el gráfico con "Clase" como eje X
  # Barras
  geom_bar(aes(y = Frecuencia, fill = "Frecuencia"), # "Frecuencia es el valor fijo para el fill
           stat = "identity", # ya tenemos los datos
           width = 0.5) + # ancho de las barras
  # Línea acumulada
  geom_line(
    aes(y = pct_acumulado * max(Frecuencia), # escalar el % acumulado al mismo eje que las frecuencias
        color = "% acumulado", # así controlamos el color desde scale_color_manual() para generar la leyenda correctamente
        group = 1), # para conectar puntos
    linewidth = 1.5 # ancho de la línea que conecta los puntos
  ) +
  # Rombos acumulados
  geom_point(
    aes(y = pct_acumulado * max(Frecuencia), # calcula la pos vertical del punto 
        color = "% acumulado"),
    shape = 23, # forma del rombo con relleno
    size = 3, fill = "#A52A2A"
  ) +
  # ejes verticales
  scale_y_continuous(
    limits = c(0, 16), # eje izq de 0 a 16
    breaks = seq(0, 16, 2), # marcará de dos en dos
    expand = c(0, 0), # para que no añada espacios ni arriba ni abajo de la gráfica
    sec.axis = sec_axis( # creamos un eje secundario a la dcha
      ~ . / (16 / 1.2),  # para que el izq que va de 0 a 16 permita que el dcho vaya de 0 a 1.2
      name = NULL, # para que no aparezca título en el lado dcho
      breaks = seq(0, 1.2, 0.2), # para que marque de 0.2 en 0.2 en nuestro intervalo
      labels = scales::label_number(accuracy = 0.1) # para que el formato tenga 1 decimal
    )
  ) +
  # Eje X
  scale_x_continuous(
    breaks = tabla_valores$pos_x,
    labels = etiquetas_x,
    name = "Clase"
  ) +
  # Colores asignados a las categorías
  scale_fill_manual(values = c("Frecuencia" = "steelblue")) +
  scale_color_manual(values = c("% acumulado" = "#A52A2A")) +
  # Leyendas
  labs(title = NULL, x = "Clase", y = "Frecuencia", fill = "", color = "") +
  guides(
    fill = guide_legend(order = 1), # primero va "frecuencia"
    color = guide_legend(order = 2, override.aes = list( # luego "%acumulado"
      shape = 23, # con el rombo
      fill = "#A52A2A",
      linewidth = 1.5 # y esta anchura de línea
    ))
  ) +
  # Tema
  theme(
    panel.background = element_rect(fill = "white"), # fondo blanco
    panel.grid.major = element_blank(), # eliminan líneas de cuadrícula
    panel.grid.minor = element_blank(),
    axis.line = element_line(color = "black"), # dibuja los ejes x e y en negro
    axis.text = element_text(size = 12), # tamaño de letra para los ejes
    axis.title = element_text(size = 14, face = "bold"), # para los títulos de los ejes
    legend.position = "right", # la leyenda la pone a la derecha
    # tmñ iconos
    legend.key.size = unit(1, "lines"), # altura
    legend.key.width = unit(2, "cm"), # ancho
    legend.text = element_text(size = 12), # tmñ de texto en la leyenda
    plot.margin = margin(t = 10, r = 10, b = 10, l = 10), # márgenes
    axis.ticks.x = element_line(),              # asegura que las muescas estén visibles
    axis.ticks.length = unit(5, "pt"),          # tamaño de muescas
    axis.text.x = element_text(angle = 45, hjust = 1), # para que las clases en el eje x aparezcan girados
    axis.title.x = element_text(margin = margin(t = 30)) # separamos el título del eje x del propio eje x
  )
dev.off()

insertImage(wb, sheet = "Hoja2", file = "grafico1_frecuencia_acumulada.png",
            startCol = 12, startRow = 9, width = 6, height = 4, units = "in")

# -----------------------------
# Gráfica 2: Diagrama de barras (frecuencias absolutas)
png("grafico2_barras.png", width = 800, height = 600)

# Crear segmentos desde un punto justo encima del eje (efecto de "tick hacia abajo")
tabla_valores$Clase <- as.numeric(as.character(tabla_valores$Clase))

clases_ordenadas <- sort(tabla_valores$Clase)
posiciones_intermedias <- (clases_ordenadas[-1] + clases_ordenadas[-length(clases_ordenadas)]) / 2

segmentos <- data.frame(
  x = posiciones_intermedias,
  xend = posiciones_intermedias,
  y = 0.1,
  yend = 0  # llegan justo al eje
)

# Gráfico
ggplot(tabla_valores, aes(x = Clase, y = Frecuencia, fill = "f.a.")) + # crea el gráf. con tabla_valores, eje x = Clase, y = Frecuencia
  geom_bar(stat = "identity", width = 0.5) + # barras y su anchura
  geom_segment(data = segmentos,
               aes(x = x, xend = xend, y = y, yend = yend),
               inherit.aes = FALSE,
               color = "black", linewidth = 0.4) +
  labs(title = "Diagrama de Barras", x = "", y = "") + # títulos y etiquetas de ejes
  scale_x_continuous(breaks = tabla_valores$Clase) +
  scale_y_continuous(
    limits = c(0, 16),
    breaks = seq(0, 16, 2),
    expand = c(0, 0)
  ) +
  scale_fill_manual(values = c("f.a." = "steelblue"), name = "") + # color de las barras, asignadas en la primera línea
  theme_minimal() +
  theme(
    plot.margin = margin(t = 30, r = 20, b = 20, l = 20), # márgenes con la imagen
    plot.title = element_text(hjust = 0.5, face = "bold", size = 30), # título de la gráfica
    axis.text.x = element_text(size = 14, angle = 45, hjust = 1), # texto en el eje X
    axis.text.y = element_text(size = 14), # texto en el eje Y
    axis.ticks.x = element_blank(), # para que no hayan ticks en el eje X
    axis.ticks.length = unit(0, "pt"),
    legend.text = element_text(size = 14), # texto de la leyenda
    legend.title = element_text(size = 14), # título de la leyenda
    panel.grid.major.x = element_blank(), # eliminar cuadrícula inferior
    panel.grid.minor.x = element_blank(),
    panel.grid.minor.y = element_blank(),
    panel.grid.major.y = element_line(color = "grey30", linewidth = 0.4),
    axis.line.x = element_line(color = "black", linewidth = 0.6), # línea del eje X
    axis.line.y = element_line(color = "black", linewidth = 0.6), # línea del eje Y
    legend.position = "right" # posición de la leyenda
  )
dev.off()

insertImage(wb, sheet = "Hoja2", file = "grafico2_barras.png",
            startCol = 12, startRow = 31, width = 5.5, height = 3.7, units = "in")

# -----------------------------
# Gráfico 3: Polígono de frecuencias + acumuladas
png("grafico3_poligono.png", width = 800, height = 600)
ggplot(tabla_valores, aes(x = Clase)) +
  scale_x_continuous( # config. eje X 
    breaks = tabla_valores$Clase,
    labels = etiquetas_x) +
  scale_y_continuous( # config. eje Y
    limits = c(0, 60),
    breaks = seq(0, 60, 10),
    expand = c(0, 0)
  ) +
  geom_line(aes(y = Frecuencia, group = 1, color = "f.a."), linewidth = 1) +
  geom_point(aes(y = Frecuencia), color = "steelblue") +
  geom_line(aes(y = cumsum(Frecuencia), group = 1, color = "f.a.a."), linewidth = 1) +
  # Rombos acumulados
  geom_point(
    aes(y = cumsum(Frecuencia), color = "f.a.a."),
    shape = 23, # forma de rombo con relleno
    size = 3,
    fill = "#A52A2A"
  ) +
  # Cuadrados acumulados
  geom_point(
    aes(y = Frecuencia, color = "f.a."),
    shape = 22, # forma de cuadrado con relleno
    size = 3,
    fill = "steelblue"
  ) +
  labs(title = "Polígono de frecuencias", x = "", y = "") +
  # Colores asignados a las categorías
  scale_color_manual(values = c("f.a." = "steelblue", "f.a.a." = "#A52A2A")) +
  theme_minimal() +
  theme(
    panel.background = element_rect(fill = "white"),
    panel.grid.major.y = element_line(color = "gray80"),
    panel.grid.major.x = element_blank(),
    panel.grid.minor = element_blank(),
    axis.line = element_line(color = "black"),
    axis.text = element_text(size = 12),
    plot.title = element_text(hjust = 0.5, size = 18, face = "bold"), # centrado y grande
    axis.title = element_text(size = 14, face = "bold"),
    legend.position = "right",
    legend.key.size = unit(1, "lines"), # afecta principalmente a la altura de la leyenda (símbolos)
    legend.key.width = unit(2, "cm"), # ancho de la leyenda (símbolos)
    legend.text = element_text(size = 12),
    plot.margin = margin(t = 10, r = 10, b = 10, l = 10),
    axis.ticks.x = element_line(),
    axis.ticks.length = unit(5, "pt"),
    axis.text.x = element_text(angle = 45, hjust = 1),
    axis.title.x = element_text(margin = margin(t = 30))
  )
dev.off()

insertImage(wb, sheet = "Hoja2", file = "grafico3_poligono.png",
            startCol = 20, startRow = 31, width = 5.5, height = 3.7, units = "in")

# -----------------------------
# Gráfico 4: Normal teórica vs f.r.
# Calcular normal y más datos, por si se perdieron
media <- mean(x)
var_muestral <- var(x)
desv_tipica <- sd(x)
f_r <- as.numeric(table(factor(x, levels = valores_completos))) / length(x)

normal_excel <- function(x, media, var_muestral, desv_tipica) {
  exp(-((x - media)^2) / (2 * var_muestral)) / (desv_tipica * sqrt(2 * pi))
}
normales <- normal_excel(valores_completos, media, var_muestral, desv_tipica)

tabla_comp <- data.frame(
  clase = valores_completos,
  f.r. = f_r,
  Normal = normales
)

png("grafico4_normal_fr.png", width = 600, height = 450)
ggplot(tabla_comp, aes(x = clase)) +
  scale_x_continuous( # config. eje x
    breaks = tabla_valores$Clase) +
  scale_y_continuous( # config. eje y
    limits = c(0, 0.35),
    breaks = seq(0, 0.35, 0.05),
    expand = c(0, 0)
  ) + geom_line(aes(y = Normal, color = "Normal"), linewidth = 1) +
  geom_line(aes(y = f.r., color = "f.r."), linewidth = 1) +
  labs(title = NULL, x = NULL, y = NULL, color = NULL) +  # elimina el título de la leyenda
  scale_color_manual(values = c("Normal" = "darkblue", "f.r." = "magenta")) +
  theme_minimal() +
  theme(
    plot.margin = margin(t = 20, r = 20, b = 20, l = 20),
    panel.background = element_rect(fill = "white"),
    panel.grid.major.y = element_line(color = "gray80"),
    legend.background = element_rect(color = "black", fill = "white", linewidth = 0.2),
    panel.grid.major.x = element_blank(),
    panel.grid.minor = element_blank(),
    axis.line = element_line(color = "black")
  )
dev.off()
insertImage(wb, sheet = "Hoja2", file = "grafico4_normal_fr.png",
            startCol = 12, startRow = 53, width = 5.5, height = 3.7, units = "in")

# -----------------------------
# Gráfico 5: Histograma de frecuencias (por intervalos)
# Asegurar que la columna de intervalos sea factor para mantener orden
tabla_final$intervalo <- factor(tabla_final[[1]], levels = tabla_final[[1]])

# Eliminar fila total (vacía)
tabla_grafico <- tabla_final[!is.na(as.numeric(tabla_final$f.a.a.)), ]

# Convertir f.a.a. a numérica
tabla_grafico$f.a.a. <- as.numeric(tabla_grafico$f.a.a.)

# Graficar
png("grafico5_histograma.png", width = 600, height = 450)
ggplot(tabla_grafico, aes(x = intervalo, y = f.a.a., fill = "f.a.a.")) +
  geom_col(color = "black", width = 1) +
  scale_fill_manual(values = c("f.a.a." = "steelblue"), name = "") + 
  labs( # títulos y etiquetas
    title = "Histograma de frecuencias",
    x = "",
    y = ""
  ) +
  scale_y_continuous(
    limits = c(0, 60), # eje izq de 0 a 16
    breaks = seq(0, 60, 10), # marcará de dos en dos
    expand = c(0, 0)
  ) + 
  theme(
    plot.margin = margin(t = 20, r = 20, b = 20, l = 20),
    panel.background = element_rect(fill = "white"),
    panel.grid.major.y = element_line(color = "gray80"),
    panel.grid.major.x = element_blank(),
    panel.grid.minor = element_blank(),
    plot.title = element_text(hjust = 0.5, size = 20, face = "bold"),
    legend.key.size = unit(0.6, "lines"),
    legend.key.width = unit(0.4, "cm"),
    legend.text = element_text(size = 10),
    axis.line = element_line(color = "gray40"),
    axis.text.x = element_text(angle = 45, hjust = 1, size = 15, color = "gray30"),
    axis.text.y = element_text(hjust = 1, size = 15, color = "gray30")
  )
dev.off()
insertImage(wb, sheet = "Hoja2", file = "grafico5_histograma.png",
            startCol = 20, startRow = 53, width = 5.5, height = 3.7, units = "in")

# -----------------------------
rm(list = setdiff(ls(), "wb"))
ls()

# -----------------------------
# -----------------------------
# HOJA 3:
addWorksheet(wb, "Hoja3")
# Leer los datos desde la hoja
datos <- read.xlsx("datos1.xlsx", sheet = 3)
# -----------------------------
# PRIMERA CELDA
name <- read.xlsx("datos1.xlsx", sheet = 3, rows = 1, cols = 1, colNames = FALSE)

writeData(wb, "Hoja3", name, startCol = 1, startRow = 1, colNames = FALSE)

comentario <- "X = peso en Kg medido sobre n = 25 individuos seleccionados al azar en una determinada población"
comentario_obj <- createComment(comment = comentario, author = "Definición",
                                visible = FALSE, width = 6, height = 4)
writeComment(wb, "Hoja3", col = 1, row = 1, comment = comentario_obj)

# -----------------------------
# PRIMERAS COLUMNAS
# Extraer la columna
x <- sort(datos[[1]])

# Calcular media
media <- mean(x)

# Crear nuevas columnas
cuadrado <- (x - media)^2
cubo <- (x - media)^3
cuarta <- (x - media)^4

# Unir datos
tabla <- data.frame(Original = x, Cuadrado = cuadrado, Cubo = cubo, Cuarta = cuarta)

tabla
#
#    Original   Cuadrado          Cubo       Cuarta
# 1      57.5 308.494096 -5.418390e+03 9.516861e+04
# 2      59.6 239.135296 -3.697988e+03 5.718569e+04
# 3      61.2 192.210496 -2.664806e+03 3.694487e+04
# 4      61.5 183.982096 -2.495533e+03 3.384941e+04
# 5      62.5 157.854096 -1.983279e+03 2.491792e+04
# 6      64.0 122.412096 -1.354367e+03 1.498472e+04
# 7      68.2  47.114496 -3.233939e+02 2.219776e+03
# 8      68.2  47.114496 -3.233939e+02 2.219776e+03
# 9      71.5  12.702096 -4.527027e+01 1.613432e+02
# 10     73.0   4.260096 -8.792838e+00 1.814842e+01
# 11     73.0   4.260096 -8.792838e+00 1.814842e+01
# 12     75.2   0.018496  2.515456e-03 3.421020e-04
# 13     77.5   5.934096  1.445546e+01 3.521350e+01
# 14     77.5   5.934096  1.445546e+01 3.521350e+01
# 15     78.1   9.217296  2.798371e+01 8.495855e+01
# 16     78.3  10.471696  3.388641e+01 1.096564e+02
# 17     78.3  10.471696  3.388641e+01 1.096564e+02
# 18     81.5  41.422096  2.665926e+02 1.715790e+03
# 19     83.6  72.863296  6.219611e+02 5.309060e+03
# 20     85.0  98.724096  9.809226e+02 9.746447e+03
# 21     85.2 102.738496  1.041357e+03 1.055520e+04
# 22     85.9 117.418896  1.272351e+03 1.378720e+04
# 23     87.8 162.205696  2.065852e+03 2.631069e+04
# 24     88.5 180.526096  2.425549e+03 3.258967e+04
# 25     94.0 358.572096  6.789921e+03 1.285739e+05
#
#


# Escribir los datos con cálculos
writeData(wb, "Hoja3", tabla, startRow = 2, startCol = 1, colNames = FALSE)

# Crear un estilo con fondo amarillo
estilo_amarillo <- createStyle(fgFill = "#FFFF00")  # Código HEX para amarillo
addStyle(wb, sheet = "Hoja3", style = estilo_amarillo, rows = 1:length(x)+1, cols = 1, gridExpand = TRUE, stack = TRUE)

# Crear un estilo con fondo naranja
estilo_naranja <- createStyle(fgFill = "#FFA500")  # Código HEX para naranja
addStyle(wb, sheet = "Hoja3", style = estilo_naranja, rows = 1:length(x)+1, cols = 2, gridExpand = TRUE, stack = TRUE)

# Crear un estilo con fondo verde
estilo_verde <- createStyle(fgFill = "#00FF00")  # Código HEX para verde
addStyle(wb, sheet = "Hoja3", style = estilo_verde, rows = 1:length(x)+1, cols = 3, gridExpand = TRUE, stack = TRUE)

# Crear un estilo con fondo azul
estilo_azul <- createStyle(fgFill = "#00E5FF")  # Código HEX para azul
addStyle(wb, sheet = "Hoja3", style = estilo_azul, rows = 1:length(x)+1, cols = 4, gridExpand = TRUE, stack = TRUE)

# -----------------------------
# TABLAS, hj3
# -----------------------------
# TABLA 1
# Definir mínimo como múltiplo de 5 hacia abajo
min_x <- floor(min(x) / 5) * 5

# Definir máximo como múltiplo de 5 hacia arriba
max_x <- ceiling(max(x) / 5) * 5

# Crear cortes de amplitud fija, por ejemplo de 10
cortes <- seq(min_x, max_x, by = 10)

# Crear intervalos
intervalos <- cut(x, breaks = cortes, right = TRUE)

# Tabla de frecuencias
tabla_fa <- as.data.frame(table(intervalos))

# Tabla de frecuencias absolutas
tabla_fa <- as.data.frame(table(intervalos))
colnames(tabla_fa) <- c("Intervalo", "f.a.")

# Calcular frecuencia relativa
tabla_fa$f.r. <- tabla_fa$`f.a.` / sum(tabla_fa$`f.a.`)

# Calcular frecuencia absoluta acumulada
tabla_fa$f.a.a. <- cumsum(tabla_fa$`f.a.`)

# Calcular frecuencia relativa acumulada
tabla_fa$f.r.a. <- cumsum(tabla_fa$f.r.)

# Extraer mínimo y máximo de cada intervalo
minimos <- cortes[-length(cortes)]
maximos <- cortes[-1]
marca_clase <- (minimos + maximos) / 2

# Agregar mínimo, máximo y marca de clase a la tabla
tabla_fa$Minimo <- minimos
tabla_fa$Maximo <- maximos
tabla_fa$m.c. <- marca_clase

# Reordenar columnas
tabla_final <- tabla_fa %>%
  select(Intervalo, Minimo, Maximo, m.c., `f.a.`, f.r., f.a.a., f.r.a.)

# Calcular la media agrupada
media_agrupada <- sum(tabla_final$m.c. * tabla_final$`f.a.`) / sum(tabla_final$`f.a.`)

# Calcular columnas T, U, V del excel (m.c - media_agrupada) al cuadrado, cubo, cuarta
tabla_final$T <- (tabla_final$m.c. - media_agrupada)^2
tabla_final$U <- (tabla_final$m.c. - media_agrupada)^3
tabla_final$V <- (tabla_final$m.c. - media_agrupada)^4

tabla_final$`altura f.r. (área)` <- tabla_final$`f.r.` / 10
tabla_final$`altura f.r.a. (área)` <- tabla_final$`f.r.a.` / 10

# Eliminar los nombres para que se vea como la del excel
colnames(tabla_final)[1] <- ""
colnames(tabla_final)[2] <- ""
colnames(tabla_final)[3] <- ""
colnames(tabla_final)[9] <- ""
colnames(tabla_final)[10] <- ""
colnames(tabla_final)[11] <- ""


# Crear fila de totales con los mismos nombres que tabla_final
fila_total <- as.data.frame(matrix(NA, nrow = 1, ncol = ncol(tabla_final)))
colnames(fila_total) <- colnames(tabla_final)

# Asignar valores a las columnas que sí tienen suma
#fila_total[1, 4] <- NA  # m.c. vacío
fila_total[1, 5] <- sum(tabla_final$`f.a.`, na.rm = TRUE)
fila_total[1, 6] <- sum(tabla_final$`f.r.`, na.rm = TRUE)

# Combinar todo
tabla_final <- rbind(tabla_final, fila_total)


# Mostrar la tabla final
tabla_final
#                 m.c. f.a. f.r. f.a.a. f.r.a.                             altura f.r. (área) altura f.r.a. (área)               
# 1 (55,65] 55 65   60    6 0.24      6   0.24 231.04 -3511.808 53379.4816              0.024                0.024
# 2 (65,75] 65 75   70    5 0.20     11   0.44  27.04  -140.608   731.1616              0.020                0.044
# 3 (75,85] 75 85   80    9 0.36     20   0.80  23.04   110.592   530.8416              0.036                0.080
# 4 (85,95] 85 95   90    5 0.20     25   1.00 219.04  3241.792 47978.5216              0.020                0.100
# 5    <NA> NA NA   NA   25 1.00     NA     NA     NA        NA         NA                 NA                   NA


# Escribir desde L3
writeData(wb, "Hoja3", x = tabla_final, 
          startCol = 12, startRow = 3, colNames = TRUE)

# Añadir borde fino
borde_thin <- createStyle(border = "TopBottomLeftRight", borderStyle = "thin")
addStyle(wb, "Hoja3", style = borde_thin,
         rows = 3:(3 + nrow(tabla_final) - 1), cols = 12:19,
         gridExpand = TRUE, stack = TRUE)

# Poner los estilos creados anteriormente en esta nueva tabla:
addStyle(wb, sheet = "Hoja3", style = estilo_amarillo, rows = 4:(3 + nrow(tabla_final) - 1), cols = 15, gridExpand = TRUE, stack = TRUE)
addStyle(wb, sheet = "Hoja3", style = estilo_naranja, rows = 4:(3 + nrow(tabla_final) - 1), cols = 20, gridExpand = TRUE, stack = TRUE)
addStyle(wb, sheet = "Hoja3", style = estilo_verde, rows = 4:(3 + nrow(tabla_final) - 1), cols = 21, gridExpand = TRUE, stack = TRUE)
addStyle(wb, sheet = "Hoja3", style = estilo_azul, rows = 4:(3 + nrow(tabla_final) - 1), cols = 22, gridExpand = TRUE, stack = TRUE)


# -----------------------------
# TABLA 2
# Crear tabla de frecuencias correctamente formateada
tabla_frecuencias <- data.frame(
  X2 = levels(intervalos),
  `f.a.` = tabla_fa$`f.a.`,
  `f.r.` = sprintf("%.2f%%", tabla_fa$f.r. * 100)  # siempre 2 decimales
)

# Sumar los porcentajes de toda la columna
suma_f_r <- sum(tabla_fa$f.r.) * 100
suma_f_r_formateado <- sprintf("%.2f%%", suma_f_r)

# Crear fila de total usando la suma
fila_total_frecuencias <- data.frame(
  X2 = "Total Resultado",
  `f.a.` = sum(tabla_fa$`f.a.`),
  `f.r.` = suma_f_r_formateado
)

# Unir tabla y fila de totales
tabla_frecuencias <- rbind(tabla_frecuencias, fila_total_frecuencias)

x <- 7 # G, para escribir en posición
# Escribir el título "Datos" arriba de las columnas
writeData(wb, sheet = "Hoja3", x = "Datos", startCol = 1 + x, startRow = 2)

# Escribir la tabla a partir de fila 2, con filtro
writeData(wb, sheet = "Hoja3", x = tabla_frecuencias, startCol = x, startRow = 3, withFilter = TRUE)


# Estilo en negrita
estilo_negrita <- createStyle(textDecoration = "bold")
addStyle(wb, "Hoja3", style = estilo_negrita,
         rows = nrow(tabla_frecuencias) + 3, cols = x:(x + 2), gridExpand = TRUE, stack = TRUE)

# Crea estilo de fondo negro y texto blanco
estilo_negro <- createStyle(
  fgFill = "black",       # Fondo negro
  fontColour = "white",   # Texto blanco para que se lea
)

# Aplicar estilos
addStyle(wb, sheet = "Hoja3", style = estilo_negro, rows = 2, cols = x + 1, gridExpand = FALSE)
addStyle(wb, sheet = "Hoja3", style = estilo_negro, rows = 3, cols = x, gridExpand = FALSE)
addStyle(wb, sheet = "Hoja3", 
         style = createStyle(border = "top", borderStyle = "thick"),
         rows = 2, cols = x:(x+2), stack = TRUE)
addStyle(wb, sheet = "Hoja3",
         style = createStyle(border = "left", borderStyle = "thick"), 
         rows = 2:(nrow(tabla_frecuencias) + 3), cols = x, stack = TRUE)
addStyle(wb, sheet = "Hoja3", 
         style = createStyle(border = "right", borderStyle = "thick"), 
         rows = 2:(nrow(tabla_frecuencias) + 3), cols = x + 2, stack = TRUE)
addStyle(wb, sheet = "Hoja3", 
         style = createStyle(border = "bottom", borderStyle = "thick"), 
         rows = nrow(tabla_frecuencias) + 3, cols = x:(x+2), stack = TRUE)
addStyle(wb, sheet = "Hoja3", 
         style = createStyle(border = "bottom", borderStyle = "medium"), 
         rows = 2, cols = x:(x+2), stack = TRUE)
addStyle(wb, sheet = "Hoja3", 
         style = createStyle(border = "bottom", borderStyle = "medium"), 
         rows = 3, cols = x:(x+2), stack = TRUE)
addStyle(wb, sheet = "Hoja3", 
         style = createStyle(border = "bottom", borderStyle = "medium"), 
         rows = nrow(tabla_frecuencias) + 2, cols = x:(x+2), stack = TRUE)
addStyle(wb, sheet = "Hoja3", 
         style = createStyle(border = "right", borderStyle = "medium"), 
         rows = 2:(nrow(tabla_frecuencias) + 3), cols = x, stack = TRUE)

# -----------------------------
# TABLA 3
# Sacar las clases automáticamente a partir de la columna Maximo
clases <- tabla_fa$Maximo[1:3]  # Tomamos los tres primeros máximos

# Crear la tabla con clases
tabla_clase <- data.frame(
  Clase = clases,
  Frecuencia = c(tabla_fa$`f.a.`[1], tabla_fa$`f.a.`[2], tabla_fa$`f.a.`[3]),
  `% acumulado` = c(
    sprintf("%.2f%%", tabla_fa$f.r.a.[1] * 100),
    sprintf("%.2f%%", tabla_fa$f.r.a.[2] * 100),
    sprintf("%.2f%%", tabla_fa$f.r.a.[3] * 100)
  )
)

# Añadir la fila de "y mayor..."
fila_y_mayor <- data.frame(
  Clase = "y mayor...",
  Frecuencia = tabla_fa$`f.a.`[4],
  `% acumulado` = sprintf("%.2f%%", tabla_fa$f.r.a.[4] * 100)
)

# Unir todo
tabla_clase <- rbind(tabla_clase, fila_y_mayor)

tabla_clase
#        Clase Frecuencia X..acumulado
# 1         65          6       24.00%
# 2         75          5       44.00%
# 3         85          9       80.00%
# 4 y mayor...          5      100.00%

writeData(wb, sheet = "Hoja3", x = tabla_clase, 
          startCol = x, startRow = 11)

addStyle(wb, sheet = "Hoja3", 
         style = createStyle(border = "top", borderStyle = "thick"),
         rows = 11, cols = x:(x+2), stack = TRUE)
addStyle(wb, sheet = "Hoja3", 
         style = createStyle(border = "bottom", borderStyle = "thin"),
         rows = 11, cols = x:(x+2), stack = TRUE)
addStyle(wb, sheet = "Hoja3", 
         style = createStyle(border = "bottom", borderStyle = "thick"),
         rows = 11+nrow(tabla_clase), cols = x:(x+2), stack = TRUE)


# -----------------------------
# TABLA 4
# Cortes extendidos
cortes_ext <- c(45, 55, 65, 75, 85, 95, 105)

# Marcas de clase
(marcas_clase_ext <- (cortes_ext[-length(cortes_ext)] + cortes_ext[-1]) / 2)
# Resultado: 50, 60, 70, 80, 90, 100

# Frecuencia relativa extendida
(f.r.ext <- c(0, tabla_fa$`f.r.`, 0))
# Resultado: 0, 0.24, 0.20, 0.36, 0.20, 0

# Frecuencia relativa acumulada extendida
(f.r.a.ext <- cumsum(f.r.ext))
# Resultado: 0, 0.24, 0.44, 0.80, 1.00, 1.00

# Media agrupada
(n <- sum(tabla_fa$`f.a.`))
# Resultado: 25

x_raw <- read.xlsx("datos1.xlsx", sheet = "Hoja3", startRow = 2, colNames = FALSE)
x <- as.numeric(unlist(x_raw))
x <- x[!is.na(x)]

# Media y varianza poblacional de los datos originales
media <- mean(x)
varianza <- var(x) * (length(x) - 1) / length(x)  # varianza poblacional
desviacion <- sqrt(varianza)

# Normal
Normal <- exp(-((marcas_clase_ext - media)^2) / (2 * varianza)) / (desviacion * sqrt(2 * pi))

Normal
# [1] 0.001717791 0.012814895 0.035113924 0.035339731 0.013063715 0.001773739

# Altura f.r. área
altura_f_r_area <- f.r.ext / 10

# Juntar y crear la tabla completa
tabla_completa <- data.frame(
  Corte = cortes_ext[-1],           # 55, 65, 75, 85, 95, 105
  `f.r.a.` = f.r.a.ext[1:6],         # 0, 0.24, 0.44, 0.80, 1.00, 1.00
  m.c. = marcas_clase_ext,           # 50, 60, 70, 80, 90, 100
  `f.r.` = f.r.ext[1:6],             # 0, 0.24, 0.20, 0.36, 0.20, 0
  Normal = Normal,                   # valores normales que pusimos anteriormente
  `altura f.r. (área)` = altura_f_r_area[1:6]  # 0, 0.024, 0.02, 0.036, 0.02, 0
)
# Eliminar nombres de columnas para que quede como en excel:
colnames(tabla_completa)[1:4] <- c("", "", "", "")

tabla_completa
#                          Normal altura.f.r...área.
# 1  55 0.00  50 0.00 0.001717791              0.000
# 2  65 0.24  60 0.24 0.012814895              0.024
# 3  75 0.44  70 0.20 0.035113924              0.020
# 4  85 0.80  80 0.36 0.035339731              0.036
# 5  95 1.00  90 0.20 0.013063715              0.020
# 6 105 1.00 100 0.00 0.001773739              0.000

# Escribir la tabla_completa a partir de la celda M10 en Hoja3
writeData(wb, 
          sheet = "Hoja3", 
          x = tabla_completa, 
          startCol = 13,   # Columna M es la número 13
          startRow = 10,   # Fila 10
          colNames = TRUE)  # Escribir los nombres de las columnas

# -----------------------------
# TABLAS DE ABAJO
x <- sort(datos[[1]])
moda_deTodosLosDatos <- as.numeric(names(sort(table(x), decreasing = TRUE)[1]))

# Media, Mediana, Cuartiles
media <- mean(x)
mediana <- median(x)
q1 <- quantile(x, 0.25)
q3 <- quantile(x, 0.75)

# Rango, IQR
rango <- max(x) - min(x)
iqr <- IQR(x)

# Varianzas y desviaciones
var_muestral <- var(x)
var_poblacional <- var_muestral * (length(x) - 1) / length(x)
desv_tipica_muestral <- sd(x)
desv_tipica_poblacional <- sqrt(var_poblacional)

# Coeficientes de variación
cv_muestral <- desv_tipica_muestral / media
cv_poblacional <- desv_tipica_poblacional / media

# Limpiar NA
tabla_final_limpia <- tabla_final[!is.na(tabla_final$m.c.), ]

# Moda agrupada
moda <- tabla_final_limpia$m.c.[which.max(tabla_final_limpia$`f.a.`)]

# Media aproximada
media_aprox <- sum(tabla_final_limpia$m.c. * tabla_final_limpia$`f.a.`) / sum(tabla_final_limpia$`f.a.`)

# Recalcular T, U, V si no está hecho
tabla_final_limpia$T <- (tabla_final_limpia$m.c. - media_aprox)^2
tabla_final_limpia$U <- (tabla_final_limpia$m.c. - media_aprox)^3
tabla_final_limpia$V <- (tabla_final_limpia$m.c. - media_aprox)^4

# Varianza agrupada
var_aprox <- sum(tabla_final_limpia$T * tabla_final_limpia$`f.a.`) / sum(tabla_final_limpia$`f.a.`)
desv_tipica_aprox <- sqrt(var_aprox)
cv_aprox <- desv_tipica_aprox / media_aprox





# Tabla 1: Medidas de tendencia central
tabla1 <- data.frame(
  Medida = c("Moda", "Media", "Media aproximada", "Mediana", "Primer cuartil", "Tercer cuartil"),
  Valor = c(moda, media, round(media_aprox, 1), mediana, q1, q3)
)

tabla1
#             Medida  Valor
# 1             Moda 80.000
# 2            Media 75.064
# 3 Media aproximada 75.200
# 4          Mediana 77.500
# 5   Primer cuartil 68.200
# 6   Tercer cuartil 83.600



# Tabla 2: Medidas de dispersión
tabla2 <- data.frame(
  Medida = c("Rango", "IQR",
             "Varianza", "Cuasivarianza", "Varianza aproximada",
             "Desviación típica", "Cuasidesviación típica", "Desviación típica aproximada",
             "Coeficiente de variación", "Coeficiente de variación (cuasivarianza)", "Coeficiente de variación aproximado"),
  Valor = c(rango, iqr,
            var_poblacional, var_muestral, var_aprox,
            desv_tipica_poblacional, desv_tipica_muestral, desv_tipica_aprox,
            cv_poblacional, cv_muestral, cv_aprox)
)

tabla2
#                                      Medida       Valor
# 1                                     Rango  36.5000000
# 2                                       IQR  15.4000000
# 3                                  Varianza  99.8423040
# 4                             Cuasivarianza 104.0024000
# 5                       Varianza aproximada 112.9600000
# 6                         Desviación típica   9.9921121
# 7                    Cuasidesviación típica  10.1981567
# 8              Desviación típica aproximada  10.6282642
# 9                  Coeficiente de variación   0.1331146
# 10 Coeficiente de variación (cuasivarianza)   0.1358595
# 11      Coeficiente de variación aproximado   0.1413333



sd_pob <- sd(x) * sqrt((n-1)/n) # desviación estándar poblacional
sd_muestral <- sd(x) # desviación estándar muestral

# 1. Coef. Pearson
coef_pearson <- (media - moda) / desv_tipica_poblacional

# 2. Coef. Fisher
m3 <- mean((x - media)^3)
coef_fisher <- m3 / sd_pob^3

# 3. Simetría “estandarizada” ("COEFICIENTE.ASIMETRIA" de Excel)
n  <- length(x)
μ  <- mean(x)
s  <- sd(x)       # desviación muestral

coef_skew_excel <- ( n / ((n-1)*(n-2)) ) * 
  sum((x - μ)^3) / (s^3)

# 4. Curtosis Fisher
m4 <- mean((x - media)^4)
coef_kurtosis_fisher <- m4 / sd_pob^4 - 3

# 5. Curtosis “estandarizada” ("CURTOSIS" de Excel)
coef_kurt_excel <- (n*(n+1) * sum((x - mean(x))^4) /
                      ((n-1)*(n-2)*(n-3)*sd_muestral^4)) -
  3 * (n-1)^2 / ((n-2)*(n-3))

# Tabla 3:
tabla3 <- data.frame(
  Medida = c("Coef. Simetría (Pearson)",
             "Coef. Simetría (Fisher)",
             "Coef. Simetría estandarizado",
             "Coef. Curtosis (Fisher)",
             "Coef. Curtosis estandarizado"),
  Valor  = c(coef_pearson,
             coef_fisher,
             coef_skew_excel,
             coef_kurtosis_fisher,
             coef_kurt_excel)
)

tabla3
#                         Medida      Valor
# 1     Coef. Simetría (Pearson) -0.4939897
# 2      Coef. Simetría (Fisher) -0.1096525
# 3 Coef. Simetría estandarizado -0.1167795
# 4      Coef. Curtosis (Fisher) -1.0071151
# 5 Coef. Curtosis estandarizado -0.9573910




# Crear estilos
estilo_borde <- createStyle(border = "TopBottomLeftRight", borderStyle = "thin")
estilo_amarillo <- createStyle(fgFill = "#FFFF00")
estilo_naranja  <- createStyle(fgFill = "#FFA500")
estilo_verde    <- createStyle(fgFill = "#00FF00")

# Escribir tabla 1
writeData(wb, "Hoja3", tabla1, startCol = 1, startRow = 28, colNames = TRUE)
addStyle(wb, "Hoja3", style = estilo_borde, rows = 28:(28 + nrow(tabla1)), cols = 1:2, gridExpand = TRUE, stack = TRUE)
addStyle(wb, "Hoja3", style = estilo_amarillo, rows = 29:(28 + nrow(tabla1)), cols = 2, gridExpand = TRUE, stack = TRUE)

# Escribir tabla 2
startRow_tabla2 <- 28 + nrow(tabla1) + 2
writeData(wb, "Hoja3", tabla2, startCol = 1, startRow = startRow_tabla2, colNames = TRUE)
addStyle(wb, "Hoja3", style = estilo_borde, rows = startRow_tabla2:(startRow_tabla2 + nrow(tabla2)), cols = 1:2, gridExpand = TRUE, stack = TRUE)
addStyle(wb, "Hoja3", style = estilo_naranja, rows = (startRow_tabla2 + 1):(startRow_tabla2 + nrow(tabla2)), cols = 2, gridExpand = TRUE, stack = TRUE)

# Escribir tabla 3
startRow_tabla3 <- startRow_tabla2 + nrow(tabla2) + 2
writeData(wb, "Hoja3", tabla3, startCol = 1, startRow = startRow_tabla3, colNames = TRUE)
addStyle(wb, "Hoja3", style = estilo_borde, rows = startRow_tabla3:(startRow_tabla3 + nrow(tabla3)), cols = 1:2, gridExpand = TRUE, stack = TRUE)
addStyle(wb, "Hoja3", style = estilo_verde, rows = (startRow_tabla3 + 1):(startRow_tabla3 + nrow(tabla3)), cols = 2, gridExpand = TRUE, stack = TRUE)



# Tabla 4: Tabla Resumen
tabla_resumen <- data.frame(
  Estadístico = c("Media", "Error típico", "Mediana", "Moda", 
                  "Desviación estándar", "Varianza de la muestra",
                  "Curtosis", "Coeficiente de asimetría", 
                  "Rango", "Mínimo", "Máximo", "Suma", "Cuenta"),
  Valor = c(
    media,
    desv_tipica_muestral / sqrt(length(x)), # Error típico = sd / sqrt(n)
    mediana,
    moda_deTodosLosDatos,
    desv_tipica_muestral,
    var_muestral,
    coef_kurt_excel, # Curtosis excel
    coef_skew_excel,      # Coeficiente de asimetría estandarizado
    rango,
    min(x),
    max(x),
    sum(x),
    length(x)
  )
)

tabla_resumen
#                 Estadístico        Valor
# 1                     Media   75.0640000
# 2              Error típico    2.0396313
# 3                   Mediana   77.5000000
# 4                      Moda   68.2000000
# 5       Desviación estándar   10.1981567
# 6    Varianza de la muestra  104.0024000
# 7                  Curtosis   -0.9573910
# 8  Coeficiente de asimetría   -0.1167795
# 9                     Rango   36.5000000
# 10                   Mínimo   57.5000000
# 11                   Máximo   94.0000000
# 12                     Suma 1876.6000000
# 13                   Cuenta   25.0000000


# Escribir tabla en Hoja3 (por ejemplo empezando en fila 52 como tu imagen)
writeData(wb, sheet = "Hoja3", x = tabla_resumen, startCol = 1, startRow = 57, colNames = TRUE)

# (Opcional) Agregar estilo de bordes finos
estilo_borde <- createStyle(border = "TopBottomLeftRight", borderStyle = "thin")
addStyle(wb, "Hoja3", style = estilo_borde, rows = 57:(57 + nrow(tabla_resumen)), cols = 1:2, gridExpand = TRUE, stack = TRUE)

# -----------------------------
# GRAFICOS, hj3
# -----------------------------
# Preparaciones previas:
colnames(tabla_final) <- c(
  "Intervalo", "Mínimo", "Máximo", "Marca de clase",
  "f.a.", "f.r.", "f.a.a.", "f.r.a.",
  "T", "U", "V",
  "altura f.r. (área)", "altura f.r.a. (área)"
)

# Preparar tabla de valores para gráficos, modificando tabla_final
tabla_valores <- tabla_final %>%
  filter(!is.na(`f.a.`)) %>%
  mutate(
    pos_x = `Marca de clase`,
    Frecuencia = `f.a.`,
    Frecuencia_acumulada = `f.a.a.`,
    pct_relativo = `f.r.`,
    pct_relativo_acumulado = `f.r.a.`
  ) %>%
  filter(Frecuencia <= 10)  # Para no dibujar barras fuera de escala

# Etiquetas de clase (por ejemplo "(55,65]", "(65,75]", etc)
etiquetas_x <- tabla_valores$Intervalo

# Crear tabla para gráficos a partir de tabla_completa (para marcas clase)
tabla_grafico <- tabla_completa %>%
  rename(
    Corte = 1,
    `f.r.a.` = 2,
    m_c = 3,
    `f.r.` = 4
  ) %>%
  mutate(
    pos_x = m_c,                # eje x basado en la marca de clase
    pct_relativo = `f.r.`,
    pct_relativo_acumulado = `f.r.a.`
  )

# -----------------------------
# Gráfico 1: Histograma de frecuencias (barras frecuencia absoluta + línea frecuencia acumulada)
png("grafico1_histograma_frecuencia.png", width = 960, height = 540)
ggplot(tabla_valores, aes(x = pos_x)) +
  # Barras de frecuencia absoluta
  geom_bar(
    aes(y = Frecuencia, fill = "f.a."),
    stat = "identity",
    width = 10,
    color = "black"  # contorno negro en las barras
  ) +
  scale_y_continuous( # config. eje Y
    limits = c(0, 10),
    breaks = seq(0, 10, 1),
    expand = c(0, 0)  # sin espacio arriba y abajo
  ) +
  scale_x_continuous( # config. eje X
    breaks = tabla_valores$pos_x,
    labels = etiquetas_x,
    expand = c(0, 0),  # sin espacio a los lados
    name = NULL
  ) +
  scale_fill_manual(values = c("f.a." = "steelblue")) +
  labs(
    title = "Histograma de frecuencias",  # título del gráfico completo
    x = NULL, # los títulos de los ejes no los ponemos
    y = NULL,
    fill = ""
  ) +
  theme(
    panel.background = element_rect(fill = "white"),
    panel.grid.major.x = element_blank(),
    panel.grid.major.y = element_line(color = "grey50", linewidth = 0.8),
    panel.grid.minor = element_blank(),
    axis.line = element_line(color = "black"),
    axis.text = element_text(size = 20, color = "black"),
    axis.title = element_text(size = 30, face = "bold", color = "black"),
    axis.title.x = element_text(margin = margin(t = 30)),
    axis.text.x = element_text(angle = 0, hjust = 0.5),
    axis.ticks.x = element_blank(),
    axis.ticks.length = unit(5, "pt"),
    plot.title = element_text(hjust = 0.5, face = "bold", size = 35),  # título del gráfico grande y en negrita
    legend.position = "right",
    legend.key.size = unit(0.5, "cm"),
    legend.key.width = unit(0.5, "cm"),
    legend.text = element_text(size = 15, color = "black"),
    plot.margin = margin(t = 30, r = 15, b = 20, l = 15)  # márgenes del gráfico con los bordes de la imagen
  )
dev.off()
insertImage(wb, sheet = "Hoja3", file = "grafico1_histograma_frecuencia.png", 
            startRow = 18, startCol = 7, width = 6.4, height = 3.6)

# -----------------------------
# Gráfico 2: Histograma de frecuencias acumuladas (barras frecuencia acumulada)
png("grafico2_histograma_frecuencia_acumulada.png", width = 960, height = 540)
ggplot(tabla_valores, aes(x = pos_x)) +
  geom_bar(
    aes(y = Frecuencia_acumulada, fill = "f.a.a."), 
    stat = "identity",
    width = 10,
    color = "black"
  ) +
  scale_y_continuous(
    limits = c(0, 30),
    breaks = seq(0, 30, 5),
    expand = c(0, 0)
  ) +
  scale_x_continuous(
    breaks = tabla_valores$pos_x,
    labels = etiquetas_x,
    expand = c(0, 0),
    name = NULL
  ) +
  scale_fill_manual(values = c("f.a.a." = "steelblue")) +
  labs(
    title = "Histograma de frecuencias",
    x = NULL,
    y = NULL,
    fill = "",
    color = ""
  ) +
  theme(
    panel.background = element_rect(fill = "white"),
    panel.grid.major.x = element_blank(),
    panel.grid.major.y = element_line(color = "grey50", linewidth = 0.8),
    panel.grid.minor = element_blank(),
    axis.line = element_line(color = "black"),
    axis.text = element_text(size = 20, color = "black"),
    axis.title = element_text(size = 30, face = "bold", color = "black"),
    axis.title.x = element_text(margin = margin(t = 30)),
    axis.text.x = element_text(angle = 0, hjust = 0.5),
    axis.ticks.x = element_blank(),
    axis.ticks.length = unit(5, "pt"),
    plot.title = element_text(hjust = 0.5, face = "bold", size = 35),
    legend.position = "right",
    legend.key.size = unit(0.5, "cm"),
    legend.key.width = unit(0.5, "cm"),
    legend.text = element_text(size = 15, color = "black"),
    plot.margin = margin(t = 30, r = 15, b = 20, l = 15)
  )
dev.off()
insertImage(wb, sheet = "Hoja3", file = "grafico2_histograma_frecuencia_acumulada.png", 
            startRow = 18, startCol = 15, width = 6.4, height = 3.6)

# -----------------------------
# Gráfico 3: Polígono de frecuencias (frecuencia relativa)
png("grafico3_poligono_frecuencia.png", width = 800, height = 600)
ggplot(tabla_grafico, aes(x = pos_x, y = pct_relativo, color = "f.r.")) +
  geom_line(
    linewidth = 1.5,
    aes(group = 1)
  ) +
  geom_point(
    shape = 15,  # cuadrado sólido
    size = 3
  ) +
  scale_y_continuous(
    limits = c(0, 0.4),
    breaks = seq(0, 0.4, 0.05),
    expand = expansion(mult = c(0, 0))
  ) +
  scale_x_continuous(
    breaks = tabla_completa$m.c.,
    labels = tabla_completa$m.c.,
    expand = expansion(mult = c(0.05, 0.05)),
    name = NULL
  ) +
  scale_color_manual(values = c("f.r." = "steelblue")) +
  labs(
    title = "Polígono de frecuencias",
    x = NULL,
    y = NULL,
    color = ""
  ) +
  guides(
    color = guide_legend(
      override.aes = list(
        linetype = 1,   # Línea continua
        shape = 15,     # Cuadrado sólido
        size = 2.5,       # Aumentamos el tamaño general (antes 1.5)
        stroke = 2      # Bordes más gruesos en el cuadrado
      )
    )
  ) +
  theme(
    panel.background = element_rect(fill = "white"),
    panel.grid.major.x = element_blank(),
    panel.grid.major.y = element_line(color = "grey70"),
    panel.grid.minor = element_blank(),
    axis.line = element_line(color = "black"),
    axis.text = element_text(size = 12, color = "black"),
    axis.title = element_text(size = 14, face = "bold", color = "black"),
    axis.title.x = element_text(margin = margin(t = 30)),
    axis.text.x = element_text(angle = 0, hjust = 0.5),
    axis.ticks.x = element_blank(),
    axis.ticks.length = unit(5, "pt"),
    plot.title = element_text(hjust = 0.5, face = "bold", size = 22),
    legend.position = "right",
    legend.key.size = unit(0.8, "cm"),
    legend.key.width = unit(1.5, "cm"),
    legend.text = element_text(size = 12, color = "black"),
    plot.margin = margin(t = 15, r = 15, b = 10, l = 15)
  )
dev.off()
insertImage(wb, sheet = "Hoja3", file = "grafico3_poligono_frecuencia.png", 
            startRow = 36, startCol = 7, width = 6, height = 4.5)

# -----------------------------
# Gráfico 4: Polígono de frecuencias acumuladas (frecuencia relativa acumulada)
colnames(tabla_completa)[1:4] <- c("Corte", "f.r.a.", "m.c.", "f.r.")
tabla_completa_extendida <- rbind(
  data.frame(
    Corte = 45,
    `f.r.a.` = 0,
    m.c. = 45,
    `f.r.` = 0,
    Normal = 0,
    `altura f.r. (área)` = 0
  ),
  tabla_completa
)
png("grafico4_poligono_frecuencia_acumulada.png", width = 800, height = 600)
ggplot(tabla_completa_extendida, aes(x = Corte, y = `f.r.a.`, color = "f.r.a.")) +
  geom_line(
    linewidth = 1.5,
    aes(group = 1)
  ) +
  geom_point(
    shape = 15,   # cuadrado sólido
    size = 5
  ) +
  scale_y_continuous(
    limits = c(0, 1.2),
    breaks = seq(0, 1.2, 0.2),
    expand = expansion(mult = c(0, 0))
  ) +
  scale_x_continuous(
    breaks = c(45, tabla_completa$Corte),  # añadimos el 45
    labels = c(45, tabla_completa$Corte),
    expand = expansion(mult = c(0.05, 0.05)),
    name = NULL
  ) +
  scale_color_manual(values = c("f.r.a." = "steelblue")) +
  labs(
    title = "Polígono de frecuencias",
    x = NULL,
    y = NULL,
    color = ""
  ) +
  guides(
    color = guide_legend(
      override.aes = list(
        linetype = 1,
        shape = 15,
        size = 2.5,
        stroke = 2 
      )
    )
  ) +
  theme(
    panel.background = element_rect(fill = "white"),
    panel.grid.major.x = element_blank(),
    panel.grid.major.y = element_line(color = "grey70"),
    panel.grid.minor = element_blank(),
    axis.line = element_line(color = "black"),
    axis.text = element_text(size = 12, color = "black"),
    axis.title = element_text(size = 14, face = "bold", color = "black"),
    axis.title.x = element_text(margin = margin(t = 30)),
    axis.text.x = element_text(angle = 0, hjust = 0.5),
    axis.ticks.x = element_blank(),
    axis.ticks.length = unit(5, "pt"),
    plot.title = element_text(hjust = 0.5, face = "bold", size = 22),
    legend.position = "right",
    legend.key.size = unit(0.5, "cm"),
    legend.key.width = unit(1.5, "cm"),
    legend.text = element_text(size = 12, color = "black"),
    plot.margin = margin(t = 15, r = 15, b = 10, l = 15)
  )
dev.off()
insertImage(wb, sheet = "Hoja3", file = "grafico4_poligono_frecuencia_acumulada.png", 
            startRow = 36, startCol = 15, width = 6, height = 4.5)


# -----------------------------
# --- Gráfico 1: Histograma de frecuencias (área) ---
png("grafico5_histo_frec_area_1.png", width = 960, height = 540)
ggplot(tabla_final, aes(x = `Marca de clase`)) +
  geom_bar(
    aes(y = `altura f.r. (área)`, fill = "f.r."),
    stat = "identity",
    width = 10,
    color = "black"
  ) +
  scale_y_continuous(
    limits = c(0, 0.04),
    breaks = seq(0, 0.04, 0.005),
    expand = c(0, 0)
  ) +
  scale_x_continuous(
    breaks = tabla_final$`Marca de clase`,
    labels = tabla_final$Intervalo,
    expand = c(0, 0),
    name = NULL
  ) +
  scale_fill_manual(values = c("f.r." = "steelblue")) +
  labs(
    title = "Histograma de frecuencias (área)",
    x = NULL,
    y = NULL,
    fill = ""
  ) +
  guides(
    fill = guide_legend(override.aes = list(
      color = "black", fill = "steelblue", shape = 22
    ))
  ) +
  theme(
    panel.background = element_rect(fill = "white"),
    panel.grid.major.x = element_blank(),
    panel.grid.major.y = element_line(color = "grey50", linewidth = 0.8),
    panel.grid.minor = element_blank(),
    axis.line = element_line(color = "black"),
    axis.text = element_text(size = 20, color = "black"),
    axis.title = element_text(size = 30, face = "bold", color = "black"),
    axis.title.x = element_text(margin = margin(t = 30)),
    axis.text.x = element_text(angle = 0, hjust = 0.5),
    axis.ticks.x = element_blank(),
    axis.ticks.length = unit(5, "pt"),
    plot.title = element_text(hjust = 0.5, face = "bold", size = 35),
    legend.position = "right",
    legend.key.size = unit(0.5, "cm"),
    legend.key.width = unit(0.5, "cm"),
    legend.text = element_text(size = 15, color = "black"),
    plot.margin = margin(t = 30, r = 15, b = 20, l = 15)
  )
dev.off()

insertImage(wb, sheet = "Hoja3", file = "grafico5_histo_frec_area_1.png", 
            startRow = 58, startCol = 7, width = 6.4, height = 3.6)

# -----------------------------
# --- Gráfico 2: Histograma de frecuencias acumuladas (área) ---
png("grafico6_histo_frec_area_2.png", width = 960, height = 540)
tabla_final_filtrada <- tabla_final %>% filter(!is.na(`altura f.r.a. (área)`))
ggplot(tabla_final_filtrada, aes(x = `Marca de clase`)) +
  geom_bar(
    aes(y = `altura f.r.a. (área)`, fill = "f.r.a."),
    stat = "identity",
    width = 10,
    color = "black"
  ) +
  scale_y_continuous(
    limits = c(0, 0.12),
    breaks = seq(0, 0.12, 0.02),
    expand = c(0, 0)
  ) +
  scale_x_continuous(
    breaks = tabla_final_filtrada$`Marca de clase`,
    labels = tabla_final_filtrada$Intervalo,
    expand = c(0, 0),
    name = NULL
  ) +
  scale_fill_manual(values = c("f.r.a." = "steelblue")) +
  labs(
    title = "Histograma de frecuencias (área)",
    x = NULL,
    y = NULL,
    fill = ""
  ) +
  guides(
    fill = guide_legend(override.aes = list(
      color = "black", fill = "steelblue", shape = 22
    ))
  ) +
  theme(
    panel.background = element_rect(fill = "white"),
    panel.grid.major.x = element_blank(),
    panel.grid.major.y = element_line(color = "grey50", linewidth = 0.8),
    panel.grid.minor = element_blank(),
    axis.line = element_line(color = "black"),
    axis.text = element_text(size = 20, color = "black"),
    axis.title = element_text(size = 30, face = "bold", color = "black"),
    axis.title.x = element_text(margin = margin(t = 30)),
    axis.text.x = element_text(angle = 0, hjust = 0.5),
    axis.ticks.x = element_blank(),
    axis.ticks.length = unit(5, "pt"),
    plot.title = element_text(hjust = 0.5, face = "bold", size = 35),
    legend.position = "right",
    legend.key.size = unit(0.5, "cm"),
    legend.key.width = unit(0.5, "cm"),
    legend.text = element_text(size = 15, color = "black"),
    plot.margin = margin(t = 30, r = 15, b = 20, l = 15)
  )
dev.off()

insertImage(wb, sheet = "Hoja3", file = "grafico6_histo_frec_area_2.png", 
            startRow = 58, startCol = 15, width = 6.4, height = 3.6)

# -----------------------------
# --- Gráfico 3: Histograma de frecuencias acumuladas ---
png("grafico6_histo_frec_3.png", width = 960, height = 540)
ggplot(tabla_final_filtrada, aes(x = `Marca de clase`)) +
  geom_bar(
    aes(y = `f.r.a.`, fill = "f.r.a."),
    stat = "identity",
    width = 10,
    color = "black"
  ) +
  scale_y_continuous(
    limits = c(0, 1.2),
    breaks = seq(0, 1.2, 0.2),
    expand = c(0, 0)
  ) +
  scale_x_continuous(
    breaks = tabla_final_filtrada$`Marca de clase`,
    labels = tabla_final_filtrada$Intervalo,
    expand = c(0, 0),
    name = NULL
  ) +
  scale_fill_manual(values = c("f.r.a." = "steelblue")) +
  labs(
    title = "Histograma de frecuencias",
    x = NULL,
    y = NULL,
    fill = ""
  ) +
  guides(
    fill = guide_legend(override.aes = list(
      color = "black", fill = "steelblue", shape = 22
    ))
  ) +
  theme(
    panel.background = element_rect(fill = "white"),
    panel.grid.major.x = element_blank(),
    panel.grid.major.y = element_line(color = "grey50", linewidth = 0.8),
    panel.grid.minor = element_blank(),
    axis.line = element_line(color = "black"),
    axis.text = element_text(size = 20, color = "black"),
    axis.title = element_text(size = 30, face = "bold", color = "black"),
    axis.title.x = element_text(margin = margin(t = 30)),
    axis.text.x = element_text(angle = 0, hjust = 0.5),
    axis.ticks.x = element_blank(),
    axis.ticks.length = unit(5, "pt"),
    plot.title = element_text(hjust = 0.5, face = "bold", size = 35),
    legend.position = "right",
    legend.key.size = unit(0.5, "cm"),
    legend.key.width = unit(0.5, "cm"),
    legend.text = element_text(size = 15, color = "black"),
    plot.margin = margin(t = 30, r = 15, b = 20, l = 15)
  )
dev.off()

insertImage(wb, sheet = "Hoja3", file = "grafico6_histo_frec_3.png", 
            startRow = 76, startCol = 15, width = 6.4, height = 3.6)


# -----------------------------
# Preparamos tabla para gráficos
tabla_grafico <- tabla_completa %>%
  rename(
    m_c = 3,
    f_r = 4,
    normal = Normal,
    altura_fr_area = `altura.f.r...área.`
  )

# Crear curvas suavizadas a partir de los datos originales
curva_normal_suave <- as.data.frame(spline(x = tabla_grafico$m_c, y = tabla_grafico$normal, n = 200))
curva_fr_area_suave <- as.data.frame(spline(x = tabla_grafico$m_c, y = tabla_grafico$altura_fr_area, n = 200))

# Graficar
png("grafico5_poligono_y_normal.png", width = 900, height = 600)
ggplot() +
  # Línea suavizada del polígono de f.r. (área)
  geom_line(data = curva_fr_area_suave, aes(x = x, y = y, color = "polígono f.r. (área)"), linewidth = 1.2) +
  # Puntos originales del polígono
  geom_point(data = tabla_grafico, aes(x = m_c, y = altura_fr_area, color = "polígono f.r. (área)"), shape = 18, size = 3) +
  # Línea suavizada de la normal
  geom_line(data = curva_normal_suave, aes(x = x, y = y, color = "Normal"), linewidth = 1.2) +
  # Puntos originales de la normal
  geom_point(data = tabla_grafico, aes(x = m_c, y = normal, color = "Normal"), shape = 15, size = 3) +
  # Colores
  scale_color_manual(
    values = c(
      "polígono f.r. (área)" = "navyblue",
      "Normal" = "magenta"
    )
  ) +
  scale_y_continuous( # config eje Y
    limits = c(0, 0.04),
    breaks = seq(0, 0.04, 0.005),
    expand = expansion(mult = c(0, 0))
  ) +
  scale_x_continuous( # config eje X
    breaks = tabla_grafico$m_c,
    labels = tabla_grafico$m_c,
    expand = expansion(mult = c(0.05, 0.05)),
    name = NULL
  ) +
  labs( # títulos y etiquetas
    title = NULL,
    x = NULL,
    y = NULL,
    color = ""
  ) +
  theme(
    panel.background = element_rect(fill = "white"),
    panel.grid.major.x = element_blank(),
    panel.grid.major.y = element_line(color = "grey50", linewidth = 0.8),
    panel.grid.minor = element_blank(),
    axis.line = element_line(color = "black"),
    axis.text = element_text(size = 14, color = "black"),
    axis.ticks.x = element_blank(),
    plot.title = element_text(hjust = 0.5, face = "bold", size = 24),
    legend.position = "right",
    legend.key.size = unit(0.5, "cm"),
    legend.text = element_text(size = 14, color = "black"),
    plot.margin = margin(t = 20, r = 20, b = 20, l = 20)
  )
dev.off()

insertImage(wb, sheet = "Hoja3",
            file = "grafico5_poligono_y_normal.png",
            startRow = 36, startCol = 23,
            width = 6, height = 4)


# -----------------------------
rm(list = setdiff(ls(), "wb"))
ls()

# -----------------------------
# Guardar el archivo
saveWorkbook(wb, "copiaHechaEnR.xlsx", overwrite = TRUE)
