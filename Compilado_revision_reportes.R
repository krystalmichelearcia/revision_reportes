#----------------------------------------------------------------
#       Código para crear base del estatus de revisión de los
#           reportes modulares para av
#---------------------------------------------------------------
# ---- Cargar las librerías necesarias ----
if(!require(pacman)){install.packages('pacman')}
pacman::p_load(tidyverse, openxlsx, readr, here, stats, writexl, Rcpp, stringr, readxl,janitor, dplyr)

# ---- Funciones y variables necesarias ----
df <- openxlsx::read.xlsx("Distribución.xlsx") #base de la distribución
nombre_archivo <- "Estatus de avance de revisión de reportes en SAMIE 24-07-25.xlsx" #base de estatus

#Leer archivo y compilar hojas 
leer_lista <- function(nombre) {

  hojas <- excel_sheets(nombre_archivo)
  
  lista_hojas <- lapply(hojas, function(hoja) {
    read_excel(nombre_archivo, sheet = hoja)
  })

  # Asigna nombres a los elementos de la lista
  names(lista_hojas) <- hojas

  return(lista_hojas)
}

 # Columnas a resumir
 cols_suma <- c(
   "Grupos",
   "Total de Sugerencias", "Sugerencias Validadas*",
   "Total de Felicitaciones", "Felicitaciones Validadas*"
 )
 
 # Función para aplicar sumas por grupo con sufijo
 agregar_sumas <- function(df, group_var, sufijo) {
   df %>%
     group_by(across(all_of(group_var))) %>%
     mutate(across(
       all_of(cols_suma),
       ~ sum(.x, na.rm = TRUE),
       .names = "{.col} por {sufijo}"
     )) %>%
     ungroup()
 }
 
 directorio <- getwd()
 # ---- Crear columnas de revisores ----
 df <- openxlsx::read.xlsx("Distribución.xlsx")
 df <- df %>%
   fill(everything(), .direction = "down")
 
 # Paso 1: limpiar y separar múltiples revisores
 df_expandido <- df %>%
   mutate(Revisor = str_replace_all(Revisor, "\n", ";")) %>%  # unificar separadores
   separate_rows(Revisor, sep = ";") %>%
   mutate(Revisor = str_trim(Revisor))  # quitar espacios extra
 
 # Paso 2: extraer nombre y rangos
 df_limpio <- df_expandido %>%
   mutate(
     nombre = str_extract(Revisor, "^[^:0-9]+"),
     rango  = str_extract(Revisor, "\\d+\\s*-\\s*\\d+"),
     inicio = as.numeric(str_extract(rango, "^\\d+")),
     fin    = as.numeric(str_extract(rango, "\\d+$")),
     
     # Detectar un solo número aislado
     solo_un_numero = str_detect(Revisor, "\\d+") & is.na(rango),
     
     reportes_asignados = case_when(
       !is.na(inicio) & !is.na(fin) ~ fin - inicio + 1,    # rango normal
       solo_un_numero ~ 1,                                 # un solo número
       TRUE ~ Número.de.reportes                                     # valor total (cuando solo hay un nombre y no hay rango)
     )
   ) %>%
   select(Fecha.de.inicio, Asignación, Número.de.reportes, nombre, reportes_asignados) %>%
   filter(!is.na(nombre)) %>%
   rename(Revisor = nombre)
 
 # Paso 3: pivotear a formato ancho
 df_ancho <- df_limpio %>% # prueba %>% #
   pivot_wider(
     id_cols = c(Fecha.de.inicio, Asignación, Número.de.reportes),
     names_from = Revisor,
     values_from = reportes_asignados
   )

 cols_suma <- unique(df_limpio$Revisor)
 
 resultados_revisores <- df_ancho %>%
   group_by(Fecha.de.inicio) %>%
   mutate(Suma_total_revisores = rowSums(across(all_of(cols_suma)), na.rm = TRUE)) %>%
   mutate(Desajuste = Número.de.reportes != Suma_total_revisores)%>%
   ungroup()

 diferencias <- resultados_revisores %>%
   filter(Número.de.reportes != Suma_total_revisores)
 
 
 revisores <- resultados_revisores %>%
   select(Asignación,which(names(.) == "Vic"):which(names(.) == "Jime"))
  
 # ---- Crear base del estatus de revisión ----
 estatus <- leer_lista(nombre_archivo)
 
 base_distribucion <- lapply(estatus, function(df) { # Filtrar cada data.frame según la primera columna
   df[grepl("^G[0-9]+M|^REC [0-9]+", df[[1]]), ]
 })
 
 compilado_distribucion <- do.call(rbind, base_distribucion) # Unir todos los data.frames filtrados
 
 compilado_distribucion <- compilado_distribucion %>%
   mutate(across(c(
     "Grupos","Reportes generados", "Reasignaciones", "Total de Sugerencias", "Sugerencias Validadas*",
     "Total de Felicitaciones", "Felicitaciones Validadas*"
   ), ~ as.numeric(.)))
 
 # Aplicar sumas por emisión
 compilado_distribucion <- compilado_distribucion %>%
   agregar_sumas("Inicio de asignación", "emisión") %>%
   agregar_sumas("Mes de revisión", "mes") %>%
   mutate(
     `% de Sugerencias Validadas* por mes` = round(
       `Sugerencias Validadas* por mes` / `Total de Sugerencias por mes` * 100, 2
     ),
     `% de Felicitaciones Validadas* por mes` = round(
       `Felicitaciones Validadas* por mes` / `Total de Sugerencias por mes` * 100, 2
     )
   ) %>%
   select(
     which(names(.) == "Mes de revisión"):which(names(.) == "Inicio de asignación"),
     Asignación,
     Grupos,
     matches("mes"),
     matches("emisión"),
     everything()
   ) %>%
   mutate(across(c("Inicio de asignación", "Fecha de envío"), ~ format(.x, "%Y-%m-%d"))) #Para no alterar el formato de las fechas
 
 # ---- Incorporar información en una sola base ----

 Compilado_estatus_revision <- left_join(compilado_distribucion, revisores, by = "Asignación")
 
 # ---- Guardar resultado procesado ---- 

 #write.csv(resultados, here(dir_productos_av,"/",nuevo_nombre), row.names = FALSE, fileEncoding = "Latin1")
 write.xlsx(Compilado_estatus_revision, here(directorio, paste0("Compilado_estatus_revision_24-07-25.xlsx")), rowNames = FALSE)
 

 
 