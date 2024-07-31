#Graficas IMAIEF para Fichas/Boletín
#IMAIEF Datos Abiertos: Actividades Secundarias; Construccion; Industrias Manufactureras
#2024/02/08

#Limpiar variables
rm(list=ls())


########## LIBRERIA ##########
library(openxlsx)
library(inegiR)
library(tidyverse)


########## DIRECTORIO ##########
#MODIFICACIÓN DE DATOS AUTOMÁTICA (la que se usa siempre, a menos que el mes en que se corra no coincida con el mes de publicación de INEGI)
#Determinar el mes y el año de la encuesta, bajo el supuesto que se publica dos meses después
#Hacer adecuaciones en caso de que se ejecute el código en una fecha distinta a la publicación de INEGI (ver "MODIFICACIÓN DE DATOS MANUAL" 4 líneas abajo)
mes = as.numeric(substr(seq(Sys.Date(), length = 2, by = "-4 months")[2],6,7))
año = as.numeric(substr(seq(Sys.Date(), length = 2, by = "-4 months")[2],1,4))

#MODIFICACIÓN DE DATOS MANUAL (la que casi no se usa, solo se corre cuando no coincida con el mes de publicación de INEGI)
#Mantener esta sección comentada. Solo usar en caso de desface de mes con INEGI. Por ejemplo, pruebas para modificar el código; o cuando no se corrió el mes que tocaba.
# mes=mes-1 #a mano

#Ponerle cero al mes2
if (mes < 10) {
  mes2 = paste0("0",mes)
  # Si el mes es menor a 10, le pegamos un 0 delante.
} else {
  mes2 = as.character(mes)
}


########## NÚMEROS CÁRDINALES ##########
#Números cardinales DEL 1 AL 20 en masculino
num_cardinales = c("primer","segundo","tercer","cuarto","quinto","sexto","séptimo","octavo","noveno","décimo","decimoprimer","decimosegundo","decimotercer","decimocuarto","decimoquinto","decimosexto","decimoséptimo","decimoctavo","decimonoveno","vigésimo")


#Directorio
directorio = paste0("C:/Users/arturo.carrillo/Documents/IMAIEF/",año," ",mes2)
setwd(directorio)


########## VARIABLES DE PERIODOS ##########
#Fecha para bases de csv
fcsv=paste(año,mes2,sep="_")

#Mes palabra
meses = c("enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre")
mespal=meses[mes]

#Selector de Mes 
sdm=seq(12,1)[mes]

#Periodos de titulares
periodo1=paste0("enero 2013-",mespal," ",año)
periodo2=paste(mespal,año)
date=paste(año,mes2,"01",sep="-")


########## NOMBRES DE ENTIDADES FEDERATIVAS ##########
#Nombres comunes de Entidades Federativas, los que en verdad se usan
nombre_ef=data.frame(c("Nacional", "Aguascalientes", "Baja California", "Baja California Sur", "Campeche", "Coahuila", "Colima", "Chiapas", "Chihuahua", "Ciudad de México", "Durango", "Guanajuato", "Guerrero", "Hidalgo", "Jalisco", "Estado de México", "Michoacán", "Morelos", "Nayarit", "Nuevo León", "Oaxaca", "Puebla", "Querétaro", "Quintana Roo", "San Luis Potosí", "Sinaloa", "Sonora", "Tabasco", "Tamaulipas", "Tlaxcala", "Veracruz", "Yucatán", "Zacatecas"))
colnames(nombre_ef) = "Nombre_EF"


########## DESCARGAR DATOS ABIERTOS DE INEGI ##########
#descargamos
download.file("https://www.inegi.org.mx/contenidos/programas/aief/2018/datosabiertos/imaief_mensual_csv.zip",
              destfile = paste0(directorio,"/archivo.zip"))
#descomprimimos
unzip("archivo.zip")

##Nombre de la base a buscar en el archivo de INEGI
cdd="conjunto_de_datos_imaief_actividad_"
nombre_base_csv1=paste0(cdd,"industrial",fcsv,".csv")
nombre_base_csv2=paste0(cdd,"23",fcsv,".csv") #Construcción
nombre_base_csv3=paste0(cdd,"31_33",fcsv,".csv") #Manufacturas
nombre_base_csv4=paste0(cdd,"21",fcsv,".csv") #Minería
nombre_base_csv5=paste0(cdd,"22",fcsv,".csv") #Servicios Públicos

#Directorio de las bases
dirBases <- paste0(directorio, '/conjunto_de_datos/')


#Leemos la base
for (i in 1:5){
  assign(paste0("base_csv",i),read.csv(paste0(dirBases,get(paste0("nombre_base_csv",i))),encoding = "UTF-8",header = FALSE))
}

#Borrar archivos que ya no son necesarios
unlink(c("conjunto_de_datos","metadatos","archivo.zip"),recursive=TRUE)


########## PARA DESCARGAR SERIES DESESTACIONALIZADAS DEL BIE DEL INEGI ##########
#Valores necesarios para funciones de descargas
fim=72+(año-2019)*12+mes #fila del índice del mes N
token="539a68cd-649d-087f-c7cf-0ca99f81093b"

#Función para descargar series del BIE y recortarlas del 2013 en adelante a dos columnas: Fecha y Valor
desca_serie_2013 <- function(clave,año,mes){
  is=inegi_series(clave,token)
  serie=is[1:fim,c(1,3)]
  serie$date=format(as.Date(serie$date),'%Y/%m')
  serie=serie[order(serie$date),]
  rownames(serie)=c()
  return(serie)
}


########## ACTIVIDADES SECUNDARIAS ##########
#Leer base
# nombre_base_csv=paste0("conjunto_de_datos_imaief_actividad_industrial",fcsv,".csv")
# base_csv=read.csv(nombre_base_csv,header = FALSE)
base_csv=base_csv1

#Transponer
base_t=t(base_csv)
rownames(base_t)=c()

#Dias y meses
base_t_anio_mes=base_t[15:(nrow(base_t)-sdm),1]
base_t_anio=substr(base_t_anio_mes,1,4)
base_t_mes=substr(base_t_anio_mes,6,15)

#Seleccionar variaciones
base_t=base_t[15:(nrow(base_t)-sdm),35:67]

#Abreviaciones de entidades (columnas de variaciones)
abrent=c("NAC","AGS","BCA","BCS","CAM","COA","COL","CHP","CHI","CDM","DUR","GUA","GUE","HID","JAL","EDM","MIC","MOR","NAY","NLE","OAX","PUE","QRO","QUI","SLP","SIN","SON","TAB","TAM","TLA","VER","YUC","ZAC")

var_ent=as.data.frame(base_t,row.names = NULL)
colnames(var_ent)=abrent

var_anio=as.data.frame(base_t_anio)
colnames(var_anio)="ANIO"
var_mes=as.data.frame(base_t_mes)
colnames(var_mes)="MES"

#Unir bases de anio, mes y variaciones
base_var=cbind.data.frame(var_anio,var_mes,var_ent)
#base_var=base_var[base_var$MES!="Anual",]

#Quitar la <P> de preliminar y <R> de revisado de los meses
base_var$MES=gsub("<P>","",base_var$MES)
base_var$MES=gsub("<R>","",base_var$MES)
base_var$MES=gsub("<","",base_var$MES)

#Cambiar Meses
base_var$MES=gsub("Enero","ENE",base_var$MES)
base_var$MES=gsub("Febrero","FEB",base_var$MES)
base_var$MES=gsub("Marzo","MAR",base_var$MES)
base_var$MES=gsub("Abril","ABR",base_var$MES)
base_var$MES=gsub("Mayo","MAY",base_var$MES)
base_var$MES=gsub("Junio","JUN",base_var$MES)
base_var$MES=gsub("Julio","JUL",base_var$MES)
base_var$MES=gsub("Agosto","AGO",base_var$MES)
base_var$MES=gsub("Septiembre","SEP",base_var$MES)
base_var$MES=gsub("Octubre","OCT",base_var$MES)
base_var$MES=gsub("Noviembre","NOV",base_var$MES)
base_var$MES=gsub("Diciembre","DIC",base_var$MES)

base_var=base_var[base_var$MES!="Anual",]
rownames(base_var)=c()
#unique(base_var$MES)


## Variacion Anual y Promedio de Jalisco ##

#variacion de indice de Jalisco
ind=select(base_var,ANIO,MES,JAL)
ind[,'ANIO']=as.numeric(as.character(ind[,'ANIO']))
ind[,'MES']=as.character(ind[,'MES'])
ind[,'JAL']=as.numeric(as.character(ind[,'JAL']))

#Variacion recortada para que coincida con promedio
var_ind=ind[13:nrow(ind),]

#Promedio de los ultimos 12 meses de las variaciones anuales

#Numero de filas de para base de promedio de variaciones
npv=nrow(ind)-(12)

#Crear base vacia
pro_var=matrix(data=0,npv,ncol(ind))
pro_var=as.data.frame(pro_var)
names(pro_var)=names(ind)

for (i in 1:npv){
  pro_var[i,1]=ind[i+12,1]
  pro_var[i,2]=ind[i+12,2]
  pro_var[i,3]=mean(ind[(i+1):(i+12),3])
}

#Anio para Grafica
#Crear base vacia
anio_graf=matrix(data=0,npv,1)
anio_graf=as.data.frame(anio_graf)
names(anio_graf)="ANIO"

for (i in 1:npv){
  if(pro_var$MES[i]=="ENE"){
    anio_graf[i,1]=pro_var$ANIO[i]
  }else{
    anio_graf[i,1]=""
  }
}

#Base de grafica para variacion y promedio de actividades secundarias
graf_var_as=cbind.data.frame(anio_graf,var_ind$MES,var_ind$JAL,pro_var$JAL)
names(graf_var_as)=c("Anio","Mes","Variacion","Promedio")
graf_var_as=graf_var_as[97:nrow(graf_var_as),]


## Ranking de ultima variacion anual ##

#Variacion Anual
ult_var=base_var[nrow(base_var),]
ult_var=t(ult_var)
ult_var=ult_var[3:nrow(ult_var),]

#Base de grafica para ranking de actividades secundarias
graf_rank_as=cbind.data.frame(nombre_ef,ult_var)
names(graf_rank_as)=c("Entidad","Variacion")
graf_rank_as[,'Variacion']=as.numeric(as.character(graf_rank_as[,'Variacion']))

graf_rank_as=graf_rank_as[order(graf_rank_as$Variacion),]



########## CONSTRUCCION ##########
#Leer base
# nombre_base_csv=paste0("conjunto_de_datos_imaief_actividad_23",fcsv,".csv")
# base_csv=read.csv(nombre_base_csv,header = FALSE)
base_csv=base_csv2

#Transponer
base_t=t(base_csv)
rownames(base_t)=c()

#Dias y meses
base_t_anio_mes=base_t[15:(nrow(base_t)-sdm),1]
base_t_anio=substr(base_t_anio_mes,1,4)
base_t_mes=substr(base_t_anio_mes,6,15)

#Seleccionar variaciones
base_t=base_t[15:(nrow(base_t)-sdm),35:67]

#Abreviaciones de entidades (columnas de variaciones)
abrent=c("NAC","AGS","BCA","BCS","CAM","COA","COL","CHP","CHI","CDM","DUR","GUA","GUE","HID","JAL","EDM","MIC","MOR","NAY","NLE","OAX","PUE","QRO","QUI","SLP","SIN","SON","TAB","TAM","TLA","VER","YUC","ZAC")

var_ent=as.data.frame(base_t,row.names = NULL)
colnames(var_ent)=abrent

var_anio=as.data.frame(base_t_anio)
colnames(var_anio)="ANIO"
var_mes=as.data.frame(base_t_mes)
colnames(var_mes)="MES"

#Unir bases de anio, mes y variaciones
base_var=cbind.data.frame(var_anio,var_mes,var_ent)
#base_var=base_var[base_var$MES!="Anual",]

#Quitar la <P> de preliminar y <R> de revisado de los meses
base_var$MES=gsub("<P>","",base_var$MES)
base_var$MES=gsub("<R>","",base_var$MES)
base_var$MES=gsub("<","",base_var$MES)

#Cambiar Meses
base_var$MES=gsub("Enero","ENE",base_var$MES)
base_var$MES=gsub("Febrero","FEB",base_var$MES)
base_var$MES=gsub("Marzo","MAR",base_var$MES)
base_var$MES=gsub("Abril","ABR",base_var$MES)
base_var$MES=gsub("Mayo","MAY",base_var$MES)
base_var$MES=gsub("Junio","JUN",base_var$MES)
base_var$MES=gsub("Julio","JUL",base_var$MES)
base_var$MES=gsub("Agosto","AGO",base_var$MES)
base_var$MES=gsub("Septiembre","SEP",base_var$MES)
base_var$MES=gsub("Octubre","OCT",base_var$MES)
base_var$MES=gsub("Noviembre","NOV",base_var$MES)
base_var$MES=gsub("Diciembre","DIC",base_var$MES)

base_var=base_var[base_var$MES!="Anual",]
rownames(base_var)=c()
#unique(base_var$MES)


## Variacion Anual y Promedio de Jalisco ##

#variacion de indice de Jalisco
ind=select(base_var,ANIO,MES,JAL)
ind[,'ANIO']=as.numeric(as.character(ind[,'ANIO']))
ind[,'MES']=as.character(ind[,'MES'])
ind[,'JAL']=as.numeric(as.character(ind[,'JAL']))

#Variacion recortada para que coincida con promedio
var_ind=ind[13:nrow(ind),]

#Promedio de los ultimos 12 meses de las variaciones anuales

#Numero de filas de para base de promedio de variaciones
npv=nrow(ind)-(12)

#Crear base vacia
pro_var=matrix(data=0,npv,ncol(ind))
pro_var=as.data.frame(pro_var)
names(pro_var)=names(ind)

for (i in 1:npv){
  pro_var[i,1]=ind[i+12,1]
  pro_var[i,2]=ind[i+12,2]
  pro_var[i,3]=mean(ind[(i+1):(i+12),3])
}

#Anio para Grafica
#Crear base vacia
anio_graf=matrix(data=0,npv,1)
anio_graf=as.data.frame(anio_graf)
names(anio_graf)="ANIO"

for (i in 1:npv){
  if(pro_var$MES[i]=="ENE"){
    anio_graf[i,1]=pro_var$ANIO[i]
  }else{
    anio_graf[i,1]=""
  }
}

#Base de grafica para variacion y promedio de construccion
graf_var_con=cbind.data.frame(anio_graf,var_ind$MES,var_ind$JAL,pro_var$JAL)
names(graf_var_con)=c("Anio","Mes","Variacion","Promedio")
graf_var_con=graf_var_con[97:nrow(graf_var_con),]


## Ranking de ultima variacion anual ##

#Variacion Anual
ult_var=base_var[nrow(base_var),]
ult_var=t(ult_var)
ult_var=ult_var[3:nrow(ult_var),]

#Base de grafica para ranking de construccion
graf_rank_con=cbind.data.frame(nombre_ef,ult_var)
names(graf_rank_con)=c("Entidad","Variacion")
graf_rank_con[,'Variacion']=as.numeric(as.character(graf_rank_con[,'Variacion']))

graf_rank_con=graf_rank_con[order(graf_rank_con$Variacion),]



########## INDUSTRIAS MANUFACTURERAS ##########
#Leer base
# nombre_base_csv=paste0("conjunto_de_datos_imaief_actividad_31_33",fcsv,".csv")
# base_csv=read.csv(nombre_base_csv,header = FALSE)
base_csv=base_csv3

#Transponer
base_t=t(base_csv)
rownames(base_t)=c()

#Dias y meses
base_t_anio_mes=base_t[15:(nrow(base_t)-sdm),1]
base_t_anio=substr(base_t_anio_mes,1,4)
base_t_mes=substr(base_t_anio_mes,6,15)

#Seleccionar variaciones
base_t=base_t[15:(nrow(base_t)-sdm),35:67]

#Abreviaciones de entidades (columnas de variaciones)
abrent=c("NAC","AGS","BCA","BCS","CAM","COA","COL","CHP","CHI","CDM","DUR","GUA","GUE","HID","JAL","EDM","MIC","MOR","NAY","NLE","OAX","PUE","QRO","QUI","SLP","SIN","SON","TAB","TAM","TLA","VER","YUC","ZAC")

var_ent=as.data.frame(base_t,row.names = NULL)
colnames(var_ent)=abrent

var_anio=as.data.frame(base_t_anio)
colnames(var_anio)="ANIO"
var_mes=as.data.frame(base_t_mes)
colnames(var_mes)="MES"

#Unir bases de anio, mes y variaciones
base_var=cbind.data.frame(var_anio,var_mes,var_ent)
#base_var=base_var[base_var$MES!="Anual",]

#Quitar la <P> de preliminar y <R> de revisado de los meses
base_var$MES=gsub("<P>","",base_var$MES)
base_var$MES=gsub("<R>","",base_var$MES)
base_var$MES=gsub("<","",base_var$MES)

#Cambiar Meses
base_var$MES=gsub("Enero","ENE",base_var$MES)
base_var$MES=gsub("Febrero","FEB",base_var$MES)
base_var$MES=gsub("Marzo","MAR",base_var$MES)
base_var$MES=gsub("Abril","ABR",base_var$MES)
base_var$MES=gsub("Mayo","MAY",base_var$MES)
base_var$MES=gsub("Junio","JUN",base_var$MES)
base_var$MES=gsub("Julio","JUL",base_var$MES)
base_var$MES=gsub("Agosto","AGO",base_var$MES)
base_var$MES=gsub("Septiembre","SEP",base_var$MES)
base_var$MES=gsub("Octubre","OCT",base_var$MES)
base_var$MES=gsub("Noviembre","NOV",base_var$MES)
base_var$MES=gsub("Diciembre","DIC",base_var$MES)

base_var=base_var[base_var$MES!="Anual",]
rownames(base_var)=c()
#unique(base_var$MES)


## Variacion Anual y Promedio de Jalisco ##

#variacion de indice de Jalisco
ind=select(base_var,ANIO,MES,JAL)
ind[,'ANIO']=as.numeric(as.character(ind[,'ANIO']))
ind[,'MES']=as.character(ind[,'MES'])
ind[,'JAL']=as.numeric(as.character(ind[,'JAL']))

#Variacion recortada para que coincida con promedio
var_ind=ind[13:nrow(ind),]

#Promedio de los ultimos 12 meses de las variaciones anuales

#Numero de filas de para base de promedio de variaciones
npv=nrow(ind)-(12)

#Crear base vacia
pro_var=matrix(data=0,npv,ncol(ind))
pro_var=as.data.frame(pro_var)
names(pro_var)=names(ind)

for (i in 1:npv){
  pro_var[i,1]=ind[i+12,1]
  pro_var[i,2]=ind[i+12,2]
  pro_var[i,3]=mean(ind[(i+1):(i+12),3])
}

#Anio para Grafica
#Crear base vacia
anio_graf=matrix(data=0,npv,1)
anio_graf=as.data.frame(anio_graf)
names(anio_graf)="ANIO"

for (i in 1:npv){
  if(pro_var$MES[i]=="ENE"){
    anio_graf[i,1]=pro_var$ANIO[i]
  }else{
    anio_graf[i,1]=""
  }
}

#Base de grafica para variacion y promedio de industrias manufactureras
graf_var_man=cbind.data.frame(anio_graf,var_ind$MES,var_ind$JAL,pro_var$JAL)
names(graf_var_man)=c("Anio","Mes","Variacion","Promedio")
graf_var_man=graf_var_man[97:nrow(graf_var_man),]



## Ranking de ultima variacion anual ##

#Variacion Anual
ult_var=base_var[nrow(base_var),]
ult_var=t(ult_var)
ult_var=ult_var[3:nrow(ult_var),]

#Base de grafica para ranking de industrias manufactureras
graf_rank_man=cbind.data.frame(nombre_ef,ult_var)
names(graf_rank_man)=c("Entidad","Variacion")
graf_rank_man[,'Variacion']=as.numeric(as.character(graf_rank_man[,'Variacion']))

graf_rank_man=graf_rank_man[order(graf_rank_man$Variacion),]


########## MINERÍA ##########
#Leer base
base_csv=base_csv4

#Transponer
base_t=t(base_csv)
rownames(base_t)=c()

#Dias y meses
base_t_anio_mes=base_t[15:(nrow(base_t)-sdm),1]
base_t_anio=substr(base_t_anio_mes,1,4)
base_t_mes=substr(base_t_anio_mes,6,15)

#Seleccionar variaciones
base_t=base_t[15:(nrow(base_t)-sdm),35:67]

#Abreviaciones de entidades (columnas de variaciones)
abrent=c("NAC","AGS","BCA","BCS","CAM","COA","COL","CHP","CHI","CDM","DUR","GUA","GUE","HID","JAL","EDM","MIC","MOR","NAY","NLE","OAX","PUE","QRO","QUI","SLP","SIN","SON","TAB","TAM","TLA","VER","YUC","ZAC")

var_ent=as.data.frame(base_t,row.names = NULL)
colnames(var_ent)=abrent

var_anio=as.data.frame(base_t_anio)
colnames(var_anio)="ANIO"
var_mes=as.data.frame(base_t_mes)
colnames(var_mes)="MES"

#Unir bases de anio, mes y variaciones
base_var=cbind.data.frame(var_anio,var_mes,var_ent)
#base_var=base_var[base_var$MES!="Anual",]

#Quitar la <P> de preliminar y <R> de revisado de los meses
base_var$MES=gsub("<P>","",base_var$MES)
base_var$MES=gsub("<R>","",base_var$MES)
base_var$MES=gsub("<","",base_var$MES)

#Cambiar Meses
base_var$MES=gsub("Enero","ENE",base_var$MES)
base_var$MES=gsub("Febrero","FEB",base_var$MES)
base_var$MES=gsub("Marzo","MAR",base_var$MES)
base_var$MES=gsub("Abril","ABR",base_var$MES)
base_var$MES=gsub("Mayo","MAY",base_var$MES)
base_var$MES=gsub("Junio","JUN",base_var$MES)
base_var$MES=gsub("Julio","JUL",base_var$MES)
base_var$MES=gsub("Agosto","AGO",base_var$MES)
base_var$MES=gsub("Septiembre","SEP",base_var$MES)
base_var$MES=gsub("Octubre","OCT",base_var$MES)
base_var$MES=gsub("Noviembre","NOV",base_var$MES)
base_var$MES=gsub("Diciembre","DIC",base_var$MES)

base_var=base_var[base_var$MES!="Anual",]
rownames(base_var)=c()
#unique(base_var$MES)


## Variacion Anual y Promedio de Jalisco ##

#variacion de indice de Jalisco
ind=select(base_var,ANIO,MES,JAL)
ind[,'ANIO']=as.numeric(as.character(ind[,'ANIO']))
ind[,'MES']=as.character(ind[,'MES'])
ind[,'JAL']=as.numeric(as.character(ind[,'JAL']))

#Variacion recortada para que coincida con promedio
var_ind=ind[13:nrow(ind),]

#Promedio de los ultimos 12 meses de las variaciones anuales

#Numero de filas de para base de promedio de variaciones
npv=nrow(ind)-(12)

#Crear base vacia
pro_var=matrix(data=0,npv,ncol(ind))
pro_var=as.data.frame(pro_var)
names(pro_var)=names(ind)

for (i in 1:npv){
  pro_var[i,1]=ind[i+12,1]
  pro_var[i,2]=ind[i+12,2]
  pro_var[i,3]=mean(ind[(i+1):(i+12),3])
}

#Anio para Grafica
#Crear base vacia
anio_graf=matrix(data=0,npv,1)
anio_graf=as.data.frame(anio_graf)
names(anio_graf)="ANIO"

for (i in 1:npv){
  if(pro_var$MES[i]=="ENE"){
    anio_graf[i,1]=pro_var$ANIO[i]
  }else{
    anio_graf[i,1]=""
  }
}

#Base de grafica para variacion y promedio de industrias manufactureras
graf_var_min=cbind.data.frame(anio_graf,var_ind$MES,var_ind$JAL,pro_var$JAL)
names(graf_var_min)=c("Anio","Mes","Variacion","Promedio")
graf_var_min=graf_var_min[97:nrow(graf_var_min),]



## Ranking de ultima variacion anual ##

#Variacion Anual
ult_var=base_var[nrow(base_var),]
ult_var=t(ult_var)
ult_var=ult_var[3:nrow(ult_var),]

#Base de grafica para ranking de industrias manufactureras
graf_rank_min=cbind.data.frame(nombre_ef,ult_var)
names(graf_rank_min)=c("Entidad","Variacion")
graf_rank_min[,'Variacion']=as.numeric(as.character(graf_rank_min[,'Variacion']))

graf_rank_min=graf_rank_min[order(graf_rank_min$Variacion),]


########## SERVICIOS PÚBLICOS ##########
#Leer base
base_csv=base_csv5

#Transponer
base_t=t(base_csv)
rownames(base_t)=c()

#Dias y meses
base_t_anio_mes=base_t[15:(nrow(base_t)-sdm),1]
base_t_anio=substr(base_t_anio_mes,1,4)
base_t_mes=substr(base_t_anio_mes,6,15)

#Seleccionar variaciones
base_t=base_t[15:(nrow(base_t)-sdm),35:67]

#Abreviaciones de entidades (columnas de variaciones)
abrent=c("NAC","AGS","BCA","BCS","CAM","COA","COL","CHP","CHI","CDM","DUR","GUA","GUE","HID","JAL","EDM","MIC","MOR","NAY","NLE","OAX","PUE","QRO","QUI","SLP","SIN","SON","TAB","TAM","TLA","VER","YUC","ZAC")

var_ent=as.data.frame(base_t,row.names = NULL)
colnames(var_ent)=abrent

var_anio=as.data.frame(base_t_anio)
colnames(var_anio)="ANIO"
var_mes=as.data.frame(base_t_mes)
colnames(var_mes)="MES"

#Unir bases de anio, mes y variaciones
base_var=cbind.data.frame(var_anio,var_mes,var_ent)
#base_var=base_var[base_var$MES!="Anual",]

#Quitar la <P> de preliminar y <R> de revisado de los meses
base_var$MES=gsub("<P>","",base_var$MES)
base_var$MES=gsub("<R>","",base_var$MES)
base_var$MES=gsub("<","",base_var$MES)

#Cambiar Meses
base_var$MES=gsub("Enero","ENE",base_var$MES)
base_var$MES=gsub("Febrero","FEB",base_var$MES)
base_var$MES=gsub("Marzo","MAR",base_var$MES)
base_var$MES=gsub("Abril","ABR",base_var$MES)
base_var$MES=gsub("Mayo","MAY",base_var$MES)
base_var$MES=gsub("Junio","JUN",base_var$MES)
base_var$MES=gsub("Julio","JUL",base_var$MES)
base_var$MES=gsub("Agosto","AGO",base_var$MES)
base_var$MES=gsub("Septiembre","SEP",base_var$MES)
base_var$MES=gsub("Octubre","OCT",base_var$MES)
base_var$MES=gsub("Noviembre","NOV",base_var$MES)
base_var$MES=gsub("Diciembre","DIC",base_var$MES)

base_var=base_var[base_var$MES!="Anual",]
rownames(base_var)=c()
#unique(base_var$MES)


## Variacion Anual y Promedio de Jalisco ##

#variacion de indice de Jalisco
ind=select(base_var,ANIO,MES,JAL)
ind[,'ANIO']=as.numeric(as.character(ind[,'ANIO']))
ind[,'MES']=as.character(ind[,'MES'])
ind[,'JAL']=as.numeric(as.character(ind[,'JAL']))

#Variacion recortada para que coincida con promedio
var_ind=ind[13:nrow(ind),]

#Promedio de los ultimos 12 meses de las variaciones anuales

#Numero de filas de para base de promedio de variaciones
npv=nrow(ind)-(12)

#Crear base vacia
pro_var=matrix(data=0,npv,ncol(ind))
pro_var=as.data.frame(pro_var)
names(pro_var)=names(ind)

for (i in 1:npv){
  pro_var[i,1]=ind[i+12,1]
  pro_var[i,2]=ind[i+12,2]
  pro_var[i,3]=mean(ind[(i+1):(i+12),3])
}

#Anio para Grafica
#Crear base vacia
anio_graf=matrix(data=0,npv,1)
anio_graf=as.data.frame(anio_graf)
names(anio_graf)="ANIO"

for (i in 1:npv){
  if(pro_var$MES[i]=="ENE"){
    anio_graf[i,1]=pro_var$ANIO[i]
  }else{
    anio_graf[i,1]=""
  }
}

#Base de grafica para variacion y promedio de industrias manufactureras
graf_var_sp=cbind.data.frame(anio_graf,var_ind$MES,var_ind$JAL,pro_var$JAL)
names(graf_var_sp)=c("Anio","Mes","Variacion","Promedio")
graf_var_sp=graf_var_sp[97:nrow(graf_var_sp),]



## Ranking de ultima variacion anual ##

#Variacion Anual
ult_var=base_var[nrow(base_var),]
ult_var=t(ult_var)
ult_var=ult_var[3:nrow(ult_var),]

#Base de grafica para ranking de industrias manufactureras
graf_rank_sp=cbind.data.frame(nombre_ef,ult_var)
names(graf_rank_sp)=c("Entidad","Variacion")
graf_rank_sp[,'Variacion']=as.numeric(as.character(graf_rank_sp[,'Variacion']))

graf_rank_sp=graf_rank_sp[order(graf_rank_sp$Variacion),]


########## COMPARATIVO ACTSEC, CONSTR, MANUFA, MINERI, SERPUB ##########
graf_comp=rbind.data.frame(graf_var_as$Variacion[nrow(graf_var_as)],graf_var_con$Variacion[nrow(graf_var_con)],
                           graf_var_man$Variacion[nrow(graf_var_man)],graf_var_min$Variacion[nrow(graf_var_min)],
                           graf_var_sp$Variacion[nrow(graf_var_sp)])
nombre_comp=data.frame(c("Actividades secundarias","Industria de la construcción","Industrias manufactureras","Minería","Servicios Públicos"))
graf_comp=cbind.data.frame(nombre_comp,graf_comp)
names(graf_comp)=c("Actividad","Variacion")


########## SERIE DESESTACIONALIZADA ##########
#Series desestacionalizadas y tendencia-ciclo > Total actividad industrial > Jalisco > Serie desestacionalizada > Índice
sdes_jal = desca_serie_2013("740354")

#Series desestacionalizadas y tendencia-ciclo > Total actividad industrial > Jalisco > Tendencia-ciclo > Índice
tencic = desca_serie_2013("740357")

#Columnas de año y mes para graficar
añoymes=select(graf_var_as,Anio,Mes)

desytc=cbind(añoymes,sdes_jal$values,tencic$values)
names(desytc) = c("Año","Mes","Serie desestacionalizada","Tendencia-ciclo")


# Variación anual con cifras desestacionalizadas --------------------------

EntidadClaveVA <- nombre_ef#Base de datos en que se va a escribir la información
Claves <- vector("numeric")

Claves[1] <- 736887 #Clave para el dato Nacional
a <- 740265 #Clave correspondiente a Aguascalientes
for (i in 2:nrow(EntidadClaveVA)){#For para obtener las claves de cada entidad
  Claves[i] <- a
  a <- a + 7 #Las claves por cada entidad ordenadas alfabéticamente crecen en constante de 7
}

EntidadClaveVA <- cbind(EntidadClaveVA, Claves)
EntidadClaveVA$Claves <- as.character(EntidadClaveVA$Claves) #Unimos las claves al dataframe con su entidad correspondiente

Variación <- vector("numeric")
Variación[1] <- inegi_series(EntidadClaveVA[1,2], token)[2,3]#Dato Nacional
for (i in 2:nrow(EntidadClaveVA)) {#For para obtener los datos a partir de las claves
  a <- inegi_series(EntidadClaveVA[i,2], token)
  Variación[i] <- a[1, 3]
}

EntidadClaveVA <- cbind(EntidadClaveVA, Variación)#Se anexa la variación anual a la entidad correspondiente
EntidadClaveVA <- select(EntidadClaveVA, -(Claves))#Quitamos las claves y dejamos el df listo para exportar
EntidadClaveVA <- rename(EntidadClaveVA, Entidad = Nombre_EF)


# Variación mensual con cifras desestacionalizadas --------------------------

EntidadClaveVM <- nombre_ef#Base de datos en que se va a escribir la información
Claves <- vector("numeric")

Claves[1] <- 736886 #Clave para el dato Nacional
a <- 740264 #Clave correspondiente a Aguascalientes
for (i in 2:nrow(EntidadClaveVM)){#For para obtener las claves de cada entidad
  Claves[i] <- a
  a <- a + 7 #Las claves por cada entidad ordenadas alfabéticamente crecen en constante de 7
}

EntidadClaveVM <- cbind(EntidadClaveVM, Claves)
EntidadClaveVM$Claves <- as.character(EntidadClaveVM$Claves) #Unimos las claves al dataframe con su entidad correspondiente

Variación <- vector("numeric")
Variación[1] <- inegi_series(EntidadClaveVM[1,2], token)[2,3]#Dato Nacional
for (i in 2:nrow(EntidadClaveVM)) {#For para obtener los datos a partir de las claves
  a <- inegi_series(EntidadClaveVM[i,2], token)
  Variación[i] <- a[1, 3]
}

EntidadClaveVM <- cbind(EntidadClaveVM, Variación)#Se anexa la variación anual a la entidad correspondiente
EntidadClaveVM <- select(EntidadClaveVM, -(Claves))#Quitamos las claves y dejamos el df listo para exportar
EntidadClaveVM <- rename(EntidadClaveVM, Entidad = Nombre_EF) 





##    TEXTO    ##
# Texto descripción 1: ACT SEC VAR -----------------------------------------------------------------
#Variación porcentual anual y variación promedio anual con cifras originales

#Variables
#mespal#Mes de las cifras reportadas
#año#Año de las cifras reportadas
añoa <- año-1#Año anterior
uv <- round(graf_var_as[nrow(graf_var_as), 3],1)#ultima variacion
uvam <- round(graf_var_as[nrow(graf_var_as)-1, 3],1)#variacion mes anterior
uvaa <- round(graf_var_as[nrow(graf_var_as)-12, 3],1)#variacion año anterior
uvp <- round(graf_var_as[nrow(graf_var_as), 4],1)#ultima variacion promedio
uvpam <- round(graf_var_as[nrow(graf_var_as)-1, 4],1)#variacion promedio mes anterior
vades <-round(((desytc[nrow(desytc),3]-desytc[nrow(desytc)-12,3])/desytc[nrow(desytc)-12,3])*100,1)#Variacion desest anual


descripcion_1 <- "De acuerdo con cifras del Indicador Mensual de la Actividad Industrial por Entidad Federativa (IMAIEF) reportadas por INEGI, la actividad industrial en Jalisco "
if (uv > 0) {
  descripcion_1 <- paste0(descripcion_1, "creció ", format(uv, nsmall = 1), "% a tasa anual en ", mespal, " de ", año, 
                          " con cifras originales, ")
  if(uv > uvam & uvam > 0){
    descripcion_1 <- paste0(descripcion_1, "incremento superior al del mes inmediato anterior, cuando se presentó un aumento de ", format(abs(uvam), nsmall = 1), "% anual. " )
  }else if(uv < uvam & uvam > 0){
    descripcion_1 <- paste0(descripcion_1, "incremento inferior al del mes inmediato anterior, cuando se presentó un aumento de ", format(abs(uvam), nsmall = 1), "% anual. " )
  }else if(uvam == 0){
    descripcion_1 <- paste0(descripcion_1, "cifra superior a la del mes inmediato anterior, cuando el indicador no presentó variación. " )
  }else if(uv > uvam & uvam < 0){
    descripcion_1 <- paste0(descripcion_1, "cifra superior a la del mes inmediato anterior, cuando se presentó una caída de ", format(abs(uvam), nsmall = 1), "% anual. " )
  }
  
} else if (uv < 0){
  descripcion_1 <- paste0(descripcion_1, "disminuyó ", format(abs(uv), nsmall = 1), "%  a tasa anual en ", mespal, " de ", 
                          año, " con cifras originales, ")
  if(uv < uvam & uvam < 0){
    descripcion_1 <- paste0(descripcion_1, "caída mayor a la del mes inmediato anterior, cuando se presentó una disminución de ", format(abs(uvam), nsmall = 1), "% anual. " )
  }else if(uv < uvam & uvam > 0){
    descripcion_1 <- paste0(descripcion_1, "cifra menor a la del mes inmediato anterior, cuando se presentó un incremento de ", format(abs(uvam), nsmall = 1), "% anual. " )
  }else if(uv > uvam & uvam < 0){
    descripcion_1 <- paste0(descripcion_1, "caída menor a la del mes inmediato anterior, cuando se presentó una disminución de ", format(abs(uvam), nsmall = 1), "% anual. " )
  }else if(uvam == 0){
    descripcion_1 <- paste0(descripcion_1, "cifra inferior a la del mes inmediato anterior, cuando el indicador no presentó variación. ")
  }
}else{
  descripcion_1 <- paste0(descripcion_1, "no presentó variación a tasa anual en ", mespal, " de ", 
                          año, " con cifras originales, ", "Mientras que la cifra del mes inmediato anterior fue de ",
                          format(uvam, nsmall = 1), "% anual. ")
}

descripcion_1 <- paste0(descripcion_1,"Cabe señalar que, con cifras desestacionalizadas, la variación anual de Jalisco fue de ",
                        format(vades,nsmall = 1),"%.")


#promedio
descripcion_1_2 <- "El desempeño estatal promedio de los últimos doce meses "

if(uvp > uvpam){#Aumenta promedio
  ac = 0 #aumento consecutivo
  while(graf_var_as[nrow(graf_var_as)-ac,4] > (graf_var_as[nrow(graf_var_as)-1-ac,4])){
    ac = ac + 1
  }
  dcp = 0 #descenso consecutivo previo
  while(graf_var_as[nrow(graf_var_as)-ac-1-dcp,4] < graf_var_as[nrow(graf_var_as)-ac-2-dcp,4]){
    dcp = dcp + 1
  }
  dcp = dcp + 1
  
  if(ac == 1 & dcp <3){
    descripcion_1_2 <-  paste0(descripcion_1_2,"registró un aumento")
  }else if(ac == 1 & dcp >= 3){
    descripcion_1_2 <- paste0(descripcion_1_2,"registró su primer aumento después de ",dcp,
                            " meses consecutivos de descensos")
  }else if(ac > 1 & ac <= 20){
    descripcion_1_2 <- paste0(descripcion_1_2,"aumentó por ",num_cardinales[ac]," mes consecutivo")
  }else if(ac > 1 & ac > 20){
    descripcion_1_2 <- paste0(descripcion_1_2,"mantiene su tendencia creciente")
  }
 descripcion_1_2 <- paste0(descripcion_1_2, " al pasar de ", format(uvpam, nsmall = 1), "% a ", 
                          format(uvp, nsmall = 1), "% respecto al mes inmediato anterior. ")
  
}else if(uvp < uvpam){#Disminuye promedio
  dc = 0 #descenso consecutivo
  while(graf_var_as[nrow(graf_var_as)-dc,4] < (graf_var_as[nrow(graf_var_as)-1-dc,4])){
    dc = dc + 1
  }
  acp = 0 #ascenso consecutivo previo
  while(graf_var_as[nrow(graf_var_as)-dc-1-acp,4] > graf_var_as[nrow(graf_var_as)-dc-2-acp,4]){
    acp = acp + 1
  }
  acp = acp + 1
  
  if(dc == 1 & acp <3){
    descripcion_1_2 <- paste0(descripcion_1_2,"registró un descenso")
  }else if(dc == 1 & acp >= 3){
    descripcion_1_2 <- paste0(descripcion_1_2,"registró su primer descenso después de ",acp,
                            " meses consecutivos de incrementos")
  }else if(dc > 1 & dc <= 20){
    descripcion_1_2 <- paste0(descripcion_1_2,"disminuyó por ",num_cardinales[dc]," mes consecutivo")
  }else if(dc > 1 & dc > 20){
    descripcion_1_2 <- paste0(descripcion_1_2,"mantiene su tendencia decreciente")
  }
 descripcion_1_2 <- paste0(descripcion_1_2, " al pasar de ", format(uvpam, nsmall = 1), "% a ", 
                          format(uvp, nsmall = 1), "% respecto al mes inmediato anterior. ")
}else{
  descripcion_1_2 <- paste0(descripcion_1_2,"se mantuvo sin cambios. ")
}


descripcion_1_3 <- paste0("La actividad industrial considera los sectores: i) Minería; ii) Generación, transmisión y distribución",
" de energía eléctrica, suministro de agua y de gas por ductos al consumidor final; iii) Construcción; y, iv) Industrias manufactureras.")



# Texto descripción 2: ACT SEC RANK ---------------------------------------
#Ranking variación porcentual anual con cifras originales

#df
re1 = filter(graf_rank_as, Entidad != "Nacional")
re1 = arrange(re1, desc(Variacion))
re = filter(mutate(re1, n = 1:nrow(re1), num = num_cardinales[1:nrow(re1)]))#Ranking estados
#Variables
#mespal#Mes de las cifras reportadas
#año#Año de las cifras reportadas
añoa <- año-1#Año anterior
nac <- round(filter(graf_rank_as, Entidad == "Nacional")[,2], 1)#Cifra nacional
jal <- round(filter(graf_rank_as, Entidad == "Jalisco")[,2], 1)#cifra jalisco
lgn <- filter(re, Entidad == "Jalisco")[,3]#Lugar de jalisco numero
lgc <- filter(re, Entidad == "Jalisco")[,4]#Lugar de jalisco cardinal

descripcion_2 <- "En variación anual en cifras originales, el desempeño de la actividad industrial de Jalisco en "
descripcion_2 <- paste0(descripcion_2, mespal, " de ", año, " ")
if (jal < nac){
  descripcion_2 <- paste0(descripcion_2, "fue inferior al nacional, ")
  if(nac > 0){
    descripcion_2 <- paste0(descripcion_2, "que registró un crecimiento anual de ", format(nac, nsmall = 1), "%. ")   
  }else if(nac < 0){
    descripcion_2 <- paste0(descripcion_2, "que registró una disminución anual de ", format(abs(nac), nsmall = 1), "%. ")   
  }else{
    descripcion_2 <- paste0(descripcion_2, "que no registró variación. ")
  }
} else if(jal > nac){
  descripcion_2 <- paste0(descripcion_2, "fue superior al nacional, ")
  if(nac > 0){
    descripcion_2 <- paste0(descripcion_2, "que registró un crecimiento anual de ", format(nac, nsmall = 1), "%. ")   
  }else if(nac < 0){
    descripcion_2 <- paste0(descripcion_2, "que registró una disminución anual de ", format(abs(nac), nsmall = 1), "%. ")   
  }else{
    descripcion_2 <- paste0(descripcion_2, "que no registró variación. ")
  }
}else{
  descripcion_2 <- paste0(descripcion_2, "estuvo en línea con el nacional,")
  if(nac > 0){
    descripcion_2 <- paste0(descripcion_2, " que registró un crecimiento anual de ", format(nac, nsmall = 1), "%. ")   
  }else if(nac < 0){
    descripcion_2 <- paste0(descripcion_2, " que registró una disminución anual de ", format(abs(nac), nsmall = 1), "%. ")   
  }else{
    descripcion_2 <- paste0(descripcion_2, "que no registró variación. ")
  }
}

if (jal > 0){
  descripcion_2 <- paste0(descripcion_2, "El crecimiento de ", format(jal, nsmall = 1), "% ",
                          "de Jalisco ubicó a la entidad en el ")
  if(lgn < 11){
    descripcion_2 <- paste0(descripcion_2, lgc, " lugar a nivel nacional ", "en cuanto a desempeño de la actividad industrial. ")
  }else{
    descripcion_2 <- paste0(descripcion_2, "lugar ", lgn, " a nivel nacional ",
                            "en cuanto a desempeño de la actividad industrial. ")
  }
}else if(jal < 0){
  descripcion_2 <- paste0(descripcion_2, "La variación de ", format(jal, nsmall = 1), "% ",
                          "de Jalisco ubicó a la entidad en el ")
  if(lgn < 11){
    descripcion_2 <- paste0(descripcion_2, lgc, " lugar a nivel nacional ", "en cuanto a desempeño de la 
                            actividad industrial. ")
  }else{
    descripcion_2 <- paste0(descripcion_2, "lugar ", lgn, " a nivel nacional ",
                            "en cuanto a desempeño de la actividad industrial. ")
  }
}else{
  descripcion_2 <- paste0(descripcion_2, "La tasa anual constante de jalisco ubicó a la entidad en el ")
  if(lgn < 11){
    descripcion_2 <- paste0(descripcion_2, lgc, " lugar a nivel nacional ", "en cuanto a desempeño de la actividad industrial. ")
  }else{
    descripcion_2 <- paste0(descripcion_2, "lugar ", lgn, " a nivel nacional ",
                            "en cuanto a desempeño de la actividad industrial. ")
  }
}


# Texto descripción 3: DESYTC -----------------------------------------------------
#Serie desestacionalizada y de tendencia-ciclo

#Variables
#mespal#Mes de las cifras reportadas
#año#Año de las cifras reportadas
añoa <- año-1#Año anterior
uv <- round(desytc[nrow(desytc),3],1)#Ultimo valor indicador
ma <- round(desytc[nrow(desytc)-1,3],1)#valor mes anterior indicador
av <- round(desytc[nrow(desytc)-12,3],1)#valor año anterior indicador
vara <- round(((uv-av)/av)*100,1)#variacion anual
av2 <- round(desytc[nrow(desytc)-24,3],1)#valor dos años anteriores indicador
tc <- round(desytc[nrow(desytc),4],1)#valor indicador tc
tca <- round(desytc[nrow(desytc)-1,4],1)#valor mes anterior indicador tc

descripcion_3 <- "Con respecto a la serie desestacionalizada, el último valor del indicador de la actividad industrial de Jalisco se ubicó en "
descripcion_3 <- paste0(descripcion_3, format(uv, nsmall = 1), ", ")

if (uv > ma){
  descripcion_3 <- paste0(descripcion_3, "cifra superior a la del mes inmediato ",
                          "anterior de ", format(ma, nsmall = 1), ", ")
}else if (uv < ma){
  descripcion_3 <- paste0(descripcion_3, "cifra inferior a la del mes inmediato ",
                          "anterior de ", format(ma, nsmall = 1), ", ")
}else{
  descripcion_3 <- paste0(descripcion_3, "en linea con la del mes inmediato anterior, ")
}


if (uv > ma){
  if(vara > 0){
    descripcion_3 <- paste0(descripcion_3, "además, se observa un crecimiento de ",
                            format(vara, nsmall = 1), "% desde su nivel de ", mespal, " de ", añoa,
                            " el cual se ubicaba en ", format(av, nsmall = 1), ". ")
  }else if(vara < 0){
    descripcion_3 <- paste0(descripcion_3, "sin embargo, se observa una disminución de ",
                            format(abs(vara), nsmall = 1), "% desde su nivel de ", mespal, " de ", añoa,
                            " el cual se ubicaba en ", format(av, nsmall = 1), ". ")
  }else{
    descripcion_3 <- paste0(descripcion_3, "sin embargo, esta cifra se encuentra en línea ",
                            "con la de ", mespal, " de ", añoa, ". ")
  }
}else if(uv < ma){
  if(vara > 0){
    descripcion_3 <- paste0(descripcion_3, "sin embargo, se observa un crecimiento de ",
                            format(vara, nsmall = 1), "% desde su nivel de ", mespal, " de ", añoa,
                            " el cual se ubicaba en ", format(av, nsmall = 1), ". ")
  }else if(vara < 0){
    descripcion_3 <- paste0(descripcion_3, "además, se observa una disminución de ",
                            format(abs(vara), nsmall = 1), "% desde su nivel de ", mespal, " de ", añoa,
                            " el cual se ubicaba en ", format(av, nsmall = 1), ". ")
  }else{
    descripcion_3 <- paste0(descripcion_3, "sin embargo, esta cifra se encuentra en línea ",
                            "con la de ", mespal, " de ", añoa, ". ")
  }
}else{
  if(vara > 0){
    descripcion_3 <- paste0(descripcion_3, "sin embargo, se observa un crecimiento de ",
                            format(vara, nsmall = 1), "% desde su nivel de ", mespal, " de ", añoa,
                            " el cual se ubicaba en ", format(av, nsmall = 1), ". ")
  }else if(vara < 0){
    descripcion_3 <- paste0(descripcion_3, "sin embargo, se observa una disminución de ",
                            format(abs(vara), nsmall = 1), "% desde su nivel de ", mespal, " de ", añoa,
                            " el cual se ubicaba en ", format(av, nsmall = 1), ". ")
  }else{
    descripcion_3 <- paste0(descripcion_3, "además, esta cifra se encuentra en línea ",
                            "con la de ", mespal, " de ", añoa, ". ")
  }
}

if (vara > 0 & (uv < av2)){
  descripcion_3 <- paste0(descripcion_3, "No obstante, aún se encuentra por ",
                          "debajo del nivel de ", mespal, " de ", año-2,
                          " de ", format(av2, nsmall = 1), ". ")
}
if (tc > tca & (vara < 0)){ 
descripcion_3 <- paste0(descripcion_3, "Sin embargo, el indicador de tendencia-ciclo ")
}else if (tc > tca & (vara > 0)){ 
  descripcion_3 <- paste0(descripcion_3, "Asimismo, el indicador de tendencia-ciclo ")
}else if (tc < tca & (vara < 0)){ 
  descripcion_3 <- paste0(descripcion_3, "Asimismo, el indicador de tendencia-ciclo ")
}else if (tc < tca & (vara > 0)){ 
  descripcion_3 <- paste0(descripcion_3, "Sin embargo, el indicador de tendencia-ciclo ")
}
#Tendencia-ciclo
if(tc > tca){#Aumenta promedio
  ac = 0 #aumento consecutivo
  while(desytc[nrow(desytc)-ac,4] > (desytc[nrow(desytc)-1-ac,4])){
    ac = ac + 1
  }
  dcp = 0 #descenso consecutivo previo
  while(desytc[nrow(desytc)-ac-1-dcp,4] < desytc[nrow(desytc)-ac-2-dcp,4]){
    dcp = dcp + 1
  }
  dcp = dcp + 1
  
  if(ac == 1 & dcp <3){
    descripcion_3 <-  paste0(descripcion_3,"registró un aumento.")
  }else if(ac == 1 & dcp >= 3){
    descripcion_3 <- paste0(descripcion_3,"registró su primer aumento después de ",dcp,
                             " meses consecutivos de descensos.")
  }else if(ac > 1 & ac <= 20){
    descripcion_3 <- paste0(descripcion_3,"aumentó por ",num_cardinales[ac]," mes consecutivo.")
  }else if(ac > 1 & ac > 20){
    descripcion_3 <- paste0(descripcion_3,"mantiene su tendencia creciente.")
  }

  
}else if(tc < tca){#Disminuye promedio
  dc = 0 #descenso consecutivo
  while(desytc[nrow(desytc)-dc,4] < (desytc[nrow(desytc)-1-dc,4])){
    dc = dc + 1
  }
  acp = 0 #ascenso consecutivo previo
  while(desytc[nrow(desytc)-dc-1-acp,4] > desytc[nrow(desytc)-dc-2-acp,4]){
    acp = acp + 1
  }
  acp = acp + 1
  
  if(dc == 1 & acp <3){
    descripcion_3 <- paste0(descripcion_3,"registró un descenso.")
  }else if(dc == 1 & acp >= 3){
    descripcion_3 <- paste0(descripcion_3,"registró su primer descenso después de ",acp,
                             " meses consecutivos de incrementos.")
  }else if(dc > 1 & dc <= 20){
    descripcion_3 <- paste0(descripcion_3,"disminuyó por ",num_cardinales[dc]," mes consecutivo.")
  }else if(dc > 1 & dc > 20){
    descripcion_3 <- paste0(descripcion_3,"mantiene su tendencia decreciente.")
  }
}else{
  descripcion_3 <- paste0(descripcion_3,"se mantuvo sin cambios. ")
}


# Texto descripción 4: DESEST RANK ANU ------------------------------------
#Ranking variación porcentual anual con cifras desestacionalizadas

#df
re1 = filter(EntidadClaveVA, Entidad != "Nacional")
re1 = arrange(re1, desc(Variación))
re = filter(mutate(re1, n = 1:nrow(re1), num = num_cardinales[1:nrow(re1)]))#Ranking estados
#Variables
#mespal#Mes de las cifras reportadas
#año#Año de las cifras reportadas
añoa <- año-1#Año anterior
nac <- round(filter(EntidadClaveVA, Entidad == "Nacional")[,2], 1)#Cifra nacional
jal <- round(filter(EntidadClaveVA, Entidad == "Jalisco")[,2], 1)#cifra jalisco
lgn <- filter(re, Entidad == "Jalisco")[,3]#Lugar de jalisco numero
lgc <- filter(re, Entidad == "Jalisco")[,4]#Lugar de jalisco cardinal

descripcion_4 <- "Con cifras desestacionalizadas, el desempeño de la actividad industrial de Jalisco en "

descripcion_4 <- paste0(descripcion_4, mespal, " de ", año, " ")

if (jal < nac){
  descripcion_4 <- paste0(descripcion_4, "fue inferior al nacional, ")
  if(nac > 0){
    descripcion_4 <- paste0(descripcion_4, "que registró un crecimiento anual de ", format(abs(nac), nsmall = 1), "%. ")   
  }else if(nac < 0){
    descripcion_4 <- paste0(descripcion_4, "que registró una disminución anual de ", format(abs(nac), nsmall = 1), "%. ")   
  }else{
    descripcion_4 <- paste0(descripcion_4, "que no registró variación. ")
  }
} else if(jal > nac){
  descripcion_4 <- paste0(descripcion_4, "fue superior al nacional, ")
  if(nac > 0){
    descripcion_4 <- paste0(descripcion_4, "que registró un crecimiento anual de ", format(abs(nac), nsmall = 1), "%. ")   
  }else if(nac < 0){
    descripcion_4 <- paste0(descripcion_4, "que registró una disminución anual de ", format(abs(nac), nsmall = 1), "%. ")   
  }else{
    descripcion_4 <- paste0(descripcion_4, "que no registró variación. ")
  }
}else{
  descripcion_4 <- paste0(descripcion_4, "estuvo en línea con el nacional, ")
  if(nac > 0){
    descripcion_4 <- paste0(descripcion_4, "que registró un crecimiento anual de ", format(abs(nac), nsmall = 1), "%. ")   
  }else if(nac < 0){
    descripcion_4 <- paste0(descripcion_4, "que registró una disminución anual de ", format(abs(nac), nsmall = 1), "%. ")   
  }else{
    descripcion_4 <- paste0(descripcion_4, "que no registró variación. ")
  }
}

if (jal > 0){
  descripcion_4 <- paste0(descripcion_4, "El crecimiento de ", format(abs(jal), nsmall = 1), "% ",
                          "de Jalisco ubicó a la entidad en el ")
  if(lgn < 11){
    descripcion_4 <- paste0(descripcion_4, lgc, " lugar a nivel nacional ", "en cuanto a desempeño de la actividad industrial. ")
  }else{
    descripcion_4 <- paste0(descripcion_4, "lugar ", lgn, " a nivel nacional ",
                            "en cuanto a desempeño de la actividad industrial. ")
  }
}else if(jal < 0){
  descripcion_4 <- paste0(descripcion_4, "La variación de ", format(jal, nsmall = 1), "% ",
                          "de Jalisco ubicó a la entidad en el ")
  if(lgn < 11){
    descripcion_4 <- paste0(descripcion_4, lgc, " lugar a nivel nacional ", "en cuanto a desempeño de la actividad industrial. ")
  }else{
    descripcion_4 <- paste0(descripcion_4, "lugar ", lgn, " a nivel nacional ",
                            "en cuanto a desempeño de la actividad industrial. ")
  }
}else{
  descripcion_4 <- paste0(descripcion_4, "La tasa anual constante de jalisco ubicó a la entidad en el ")
  if(lgn < 11){
    descripcion_4 <- paste0(descripcion_4, lgc, " lugar a nivel nacional ", "en cuanto a desempeño de la actividad industrial. ")
  }else{
    descripcion_4 <- paste0(descripcion_4, "lugar ", lgn, " a nivel nacional ",
                            "en cuanto a desempeño de la actividad industrial. ")
  }
}


# Texto descripción 5: DESEST RANK MEN  ---------------------------------------------------
#Ranking variación porcentual mensual con cifras desestacionalizadas

#df
re1 = filter(EntidadClaveVM, Entidad != "Nacional")
re1 = arrange(re1, desc(Variación))
re = filter(mutate(re1, n = 1:nrow(re1), num = num_cardinales[1:nrow(re1)]))#Ranking estados
#Variables
#mespal#Mes de las cifras reportadas
#año#Año de las cifras reportadas
añoa <- año-1#Año anterior
nac <- round(filter(EntidadClaveVM, Entidad == "Nacional")[,2], 1)#Cifra nacional
jal <- round(filter(EntidadClaveVM, Entidad == "Jalisco")[,2], 1)#cifra jalisco
lgn <- filter(re, Entidad == "Jalisco")[,3]#Lugar de jalisco numero
lgc <- filter(re, Entidad == "Jalisco")[,4]#Lugar de jalisco cardinal


descripcion_5 <- "En variación mensual con cifras desestacionalizadas, el desempeño de la actividad industrial de Jalisco en "
descripcion_5 <- paste0(descripcion_5, mespal, " de ", año, " ")

if (jal < nac){
  descripcion_5 <- paste0(descripcion_5, "fue inferior al nacional, ")
  if(nac > 0){
    descripcion_5 <- paste0(descripcion_5, "que registró un crecimiento mensual de ", format(abs(nac), nsmall = 1), "%. ")   
  }else if(nac < 0){
    descripcion_5 <- paste0(descripcion_5, "que registró una disminución mensual de ", format(abs(nac), nsmall = 1), "%. ")   
  }else{
    descripcion_5 <- paste0(descripcion_5, "que no registró variación. ")
  }
} else if(jal > nac){
  descripcion_5 <- paste0(descripcion_5, "fue superior al nacional, ")
  if(nac > 0){
    descripcion_5 <- paste0(descripcion_5, "que registró un crecimiento mensual de ", format(abs(nac), nsmall = 1), "%. ")   
  }else if(nac < 0){
    descripcion_5 <- paste0(descripcion_5, "que registró una disminución mensual de ", format(abs(nac), nsmall = 1), "%. ")   
  }else{
    descripcion_5 <- paste0(descripcion_5, "que no registró variación. ")
  }
}else{
  descripcion_5 <- paste0(descripcion_5, "estuvo en línea con el nacional, ")
  if(nac > 0){
    descripcion_5 <- paste0(descripcion_5, "que registró un crecimiento mensual de ", format(abs(nac), nsmall = 1), "%. ")   
  }else if(nac < 0){
    descripcion_5 <- paste0(descripcion_5, "que registró una disminución mensual de ", format(abs(nac), nsmall = 1), "%. ")   
  }else{
    descripcion_5 <- paste0(descripcion_5, "que no registró variación. ")
  }
}

if (jal > 0){
  descripcion_5 <- paste0(descripcion_5, "El crecimiento de ", format(abs(jal), nsmall = 1), "% ",
                          "de Jalisco ubicó a la entidad en el ")
  if(lgn < 11){
    descripcion_5 <- paste0(descripcion_5, lgc, " lugar a nivel nacional ", "en cuanto a desempeño de la actividad industrial. ")
  }else{
    descripcion_5 <- paste0(descripcion_5, "lugar ", lgn, " a nivel nacional ",
                            "en cuanto a desempeño de la actividad industrial. ")
  }
}else if(jal < 0){
  descripcion_5 <- paste0(descripcion_5, "La variación de ", format(jal, nsmall = 1), "% ",
                          "de Jalisco ubicó a la entidad en el ")
  if(lgn < 11){
    descripcion_5 <- paste0(descripcion_5, lgc, " lugar a nivel nacional ", "en cuanto a desempeño de la actividad industrial. ")
  }else{
    descripcion_5 <- paste0(descripcion_5, "lugar ", lgn, " a nivel nacional ",
                            "en cuanto a desempeño de la actividad industrial. ")
  }
}else{
  descripcion_5 <- paste0(descripcion_5, "La tasa anual constante de jalisco ubicó a la entidad en el ")
  if(lgn < 11){
    descripcion_5 <- paste0(descripcion_5, lgc, " lugar a nivel nacional ", "en cuanto a desempeño de la actividad industrial. ")
  }else{
    descripcion_5 <- paste0(descripcion_5, "lugar ", lgn, " a nivel nacional ",
                            "en cuanto a desempeño de la actividad industrial. ")
  }
}


# Texto descripcion 6: ACT COM ----------------------------------------------------
#Comparativo de Variación porcentual anual del IMAIEF de actividades secundarias, industria de la construcción, industrias manufactureras, minería y servicios públicos
#Variables
uv <- round(graf_comp[1,2],1) #Última variación del IMAIEF de actividades secundarias 
co <- round(graf_comp[2,2],1) #Industria de la construcción
ma <- round(graf_comp[3,2],1) #Industrias manufactureras
mi <- round(graf_comp[4,2],1) #Minería  
sp <- round(graf_comp[5,2],1) #Servicios Públicos

descripcion_6 <- "Si bien la actividad industrial en Jalisco "
if (uv > 0){
  descripcion_6 <- paste0(descripcion_6,"presentó un crecimiento anual de ", format(uv,nsmall = 1), "% en ", mespal, " de ", año, ", ")
  if(ma > co){
    if(co >= 0){
      descripcion_6 <- paste0(descripcion_6,"las industrias manufactureras tuvieron un incremento de ", format(ma,nsmall = 1),
                              "% anual, siendo este sector el que ocasionó que la variación global de la actividad industrial en la entidad aumentara a una mayor tasa. Por su parte, la industria de la construcción presentó un crecimiento de ",
                              format(co,nsmall = 1), "% anual.")
    }else if(co < 0){
      descripcion_6 <- paste0(descripcion_6,"la industria de la construcción presentó una caída anual de ",format(abs(co),nsmall = 1),
                              "%, mientras que las industrias manufactureras aumentaron ", format(ma,nsmall = 1),
                              "% anual, siendo estas industrias las que ocasionaron que la variación global de la actividad industrial aumentara.")
    }
  }else if(co > ma){
    if(ma >= 0){
      descripcion_6 <- paste0(descripcion_6,"la industria de la construcción tuvo un incremento de ", format(co,nsmall = 1),
                              "% anual, siendo este sector el que ocasionó que la variación global de la actividad industrial en la entidad aumentara a una mayor tasa. Por su parte, las industrias manufactureras presentaron un crecimiento de ",
                              format(ma,nsmall = 1), "% anual.")
    }else if(ma < 0){
      descripcion_6 <- paste0(descripcion_6,"las industrias manufactureras presentaron una caída anual de ",format(abs(ma),nsmall = 1),
                              "%, mientras que la industria de la construcción aumentó ", format(co,nsmall = 1),
                              "% anual, siendo esta industria la que ocasionó que la variación global de la actividad industrial aumentara.")
    }
  }else if(ma == co){
    if(ma >= 0){
      descripcion_6 <- paste0(descripcion_6,"tanto la industria de la construcción como las industrias manufactureras crecieron ",
                              format(ma,nsmall = 1), "% anual.")
    }else if(ma < 0){
      descripcion_6 <- paste0(descripcion_6,"tanto la industria de la construcción como las industrias manufactureras disminuyeron ",
                              format(abs(ma),nsmall = 1), "% anual.")
    }
  }
}else if(uv < 0){
  descripcion_6 <- paste0(descripcion_6,"presentó una caída anual de ", format(uv,nsmall = 1), "% en ", mespal, " de ", año, ", ")
  if(ma > co){
    if(ma > 0){
      descripcion_6 <- paste0(descripcion_6,"las industrias manufactureras tuvieron un incremento de ",format(ma,nsmall = 1),
                              "% anual, siendo este sector el que ocasionó que la variación global de la actividad industrial en la entidad cayera a una menor tasa. Por su parte, la industria de la construcción presentó una caída de ",
                              format(abs(co),nsmall = 1),"% anual.")
    }else if(ma < 0){
      descripcion_6 <- paste0(descripcion_6,"las industrias manufactureras presentaron una caída anual de ", format(abs(ma),nsmall = 1),
                              "%, mientras que la industria de la construcción cayó ", format(abs(co),nsmall = 1),
                              "% anual, siendo esta industria la que ocasionó que la variación global de la actividad industrial cayera a una mayor tasa.")
    }else if(ma == 0){
      descripcion_6 <- paste0(descripcion_6,"las industrias manufactureras no presentaron variación anual, siendo este sector el que ocasionó que la variación global de la actividad industrial en la entidad cayera a una menor tasa. Por su parte, la industria de la construcción tuvo una caída de ",
                              format(abs(co),nsmall = 1),"% anual.")
    }
  }else if (co > ma){
    if(co > 0){
      descripcion_6 <- paste0(descripcion_6,"la industria de la construcción presentó un crecimiento anual de ",format(co,nsmall = 1),
                              "%, mientras que las industrias manufactureras disminuyeron", format(abs(ma),nsmall = 1),
                              "% anual, siendo estas industrias las que ocasionaron que la variación global de la actividad industrial cayera.")
    }else if(co < 0){
      descripcion_6 <- paste0(descripcion_6," la industria de la construcción tuvo una caída anual de ", format(abs(co),nsmall = 1),
                              "%, mientras que las industrias manufactureras cayeron ", format(abs(ma),nsmall = 1),
                              "% anual, siendo estas industrias las que ocasionaron que la variación global de la actividad industrial cayera a una mayor tasa.")
    }else if(con == 0){
      descripcion_6 <- paste0(descripcion_6,"la industria de la construcción no presentó variación anual, siendo este sector el que ocasionó que la variación global de la actividad industrial en la entidad cayera a una menor tasa. Por su parte, las industrias manufactureras tuvieron una caída de ",
                              format(abs(ma),nsmall = 1),"% anual.")
    }
  }else if(ma == co){
    if(ma >= 0){
      descripcion_6 <- paste0(descripcion_6,"tanto la industria de la construcción como las industrias manufactureras crecieron ",
                              format(ma,nsmall = 1), "% anual.")
    }else if(ma < 0){
      descripcion_6 <- paste0(descripcion_6,"tanto la industria de la construcción como las industrias manufactureras disminuyeron ",
                              format(abs(ma),nsmall = 1), "% anual.")
    }
  }
}else if(uv == 0){
  descripcion_6 <- paste0(descripcion_6,"no presentó variación en", mespal, " de ", año, ", ")
  if(ma == co){
    if(ma >= 0){
      descripcion_6 <- paste0(descripcion_6,"tanto la industria de la construcción como las industrias manufactureras crecieron ",
                              format(ma,nsmall = 1), "% anual.")
    }else if(ma < 0){
      descripcion_6 <- paste0(descripcion_6,"tanto la industria de la construcción como las industrias manufactureras disminuyeron ",
                              format(abs(ma),nsmall = 1), "% anual.")
    }
  }else if(ma >= 0 & con >= 0){
    descripcion_6 <- paste0(descripcion_6,"las industrias manufactureras tuvieron un crecimiento de ", format(ma,nsmall = 1),
                            "% anual. Por su parte, la industria de la construcción presentó un incremento de ",
                            format(co,nsmall = 1),"% anual.")
  }else if(ma >= 0 & con < 0){
    descripcion_6 <- paste0(descripcion_6,"las industrias manufactureras tuvieron un crecimiento de ", format(ma,nsmall = 1),
                            "% anual. Por su parte, la industria de la construcción presentó una caída de ",
                            format(abs(co),nsmall = 1),"% anual.")
  }else if(ma < 0 & con >= 0){
    descripcion_6 <- paste0(descripcion_6,"la industria de la construcción tuvo un crecimiento de  ", format(co,nsmall = 1),
                            "% anual. Por su parte, as industrias manufactureras presentaron una caída de ",
                            format(abs(ma),nsmall = 1),"% anual.")
  }else if(ma < 0 & con < 0){
    descripcion_6 <- paste0(descripcion_6,"las industrias manufactureras tuvieron una caída de ", format(abs(ma),nsmall = 1),
                            "% anual. Por su parte, la industria de la construcción presentó una disminución de ",
                            format(abs(co),nsmall = 1),"% anual.")
  }
}else{
  descripcion_6 <- paste0(descripcion_6,"CASO ESPECIAL, REDACCIÓN NO DISPONIBLE")
}

descripcion_6 <- paste0(descripcion_6," Por otro lado, los sectores de minería y servicios públicos (industrias de energía eléctrica, suministro de agua y de gas), los cuales tienen una menor contribución en la variación total del sector industrial de Jalisco, presentaron variaciones anuales de ",
                        format(mi,nsmall = 1), "% y ",format(sp,nsmall = 1),"%, respectivamente.")



# Texto descripcion 7: CON VAR ---------------------------------------------------
#Variación porcentual anual y variación promedio anual con cifras originales de la Industria de la Construcción 

#Variables
#mespal #Mes de las cifras reportadas
#año#Año de las cifras reportadas
añoa <- año-1#Año anterior
uv <- round(graf_var_con[nrow(graf_var_con), 3],1)#ultima variacion
uvam <- round(graf_var_con[nrow(graf_var_con)-1, 3],1)#variacion mes anterior
uvaa <- round(graf_var_con[nrow(graf_var_con)-12, 3],1)#variacion año anterior
uvp <- round(graf_var_con[nrow(graf_var_con), 4],1)#ultima variacion promedio
uvpam <- round(graf_var_con[nrow(graf_var_con)-1, 4],1)#variacion promedio mes anterior

descripcion_7 <- "Con relación a los sectores que componen las actividades secundarias, la industria de la construcción en Jalisco "
#Comparación variación mes anterior:
if (uv > 0) {
  descripcion_7 <- paste0(descripcion_7, "presentó un crecimiento de ", format(uv, nsmall = 1), "% a tasa anual en ", mespal, " de ", año, ", ")
  if(uv > uvam & uvam > 0){
    descripcion_7 <- paste0(descripcion_7, "incremento superior al del mes inmediato anterior, cuando se presentó un aumento de ", format(abs(uvam), nsmall = 1), "% anual. " )
  }else if(uv < uvam & uvam > 0){
    descripcion_7 <- paste0(descripcion_7, "incremento inferior al del mes inmediato anterior, cuando se presentó un aumento de ", format(abs(uvam), nsmall = 1), "% anual. " )
  }else if(uvam == 0){
    descripcion_7 <- paste0(descripcion_7, "cifra superior a la del mes inmediato anterior, cuando el indicador no presentó variación. " )
  }else if(uv > uvam & uvam < 0){
    descripcion_7 <- paste0(descripcion_7, "cifra superior a la del mes inmediato anterior, cuando se presentó una caída de ", format(abs(uvam), nsmall = 1), "% anual. " )
  }
  
} else if (uv < 0){
  descripcion_7 <- paste0(descripcion_7, "presentó una disminución de ", format(abs(uv), nsmall = 1), "% a tasa anual en ", mespal, " de ", año, ", ")
  if(uv < uvam & uvam < 0){
    descripcion_7 <- paste0(descripcion_7, "caída mayor a la del mes inmediato anterior, cuando se presentó un descenso de ", format(abs(uvam), nsmall = 1), "% anual. " )
  }else if(uv < uvam & uvam > 0){
    descripcion_7 <- paste0(descripcion_7, "cifra inferior a la del mes inmediato anterior, cuando se presentó un crecimiento de ", format(abs(uvam), nsmall = 1), "% anual. " )
  }else if(uv > uvam & uvam < 0){
    descripcion_7 <- paste0(descripcion_7, "caída menor a la del mes inmediato anterior, cuando se presentó un descenso de ", format(abs(uvam), nsmall = 1), "% anual. " )
  }else if(uvam == 0){
    descripcion_7 <- paste0(descripcion_7, "cifra inferior a la del mes inmediato anterior, cuando el indicador no presentó variación. ")
  }
}else{
  descripcion_7 <- paste0(descripcion_7, "no presentó variación a tasa anual en ", mespal, " de ", año, ", ")
  if(uvam > 0){
    descripcion_7 <- paste0(descripcion_7, "mientras que el mes inmediato anterior se observó un crecimiento de ",
                            format(uvam, nsmall = 1), "% anual. ")
  }else if(uvam < 0){
    descripcion_7 <- paste0(descripcion_7, "mientras que el mes inmediato anterior se observó una caída de ",
                            format(abs(uvam), nsmall = 1), "% anual. ")
  }else{
    descripcion_7 <- paste0(descripcion_7,"de igual manera que el mes inmediato anterior, cuando tampocó se presentaron cambios en el indicador. ")
  }
}

#Comparación variación año anterior:
#Conector
if((uv > uvam & uv > uvaa) || (uv <= uvam & uv <= uvaa)){
  descripcion_7 <- paste0(descripcion_7,"Además, ")
}else{
  descripcion_7 <- paste0(descripcion_7,"Sin embargo, ")
}

if(uv > 0){
  if(uv > uvaa){
    if(uvaa > 0){
      descripcion_7 <- paste0(descripcion_7,"el crecimiento de ", mespal, " de ", año, " fue mayor al de ",
                              mespal, " de ", añoa, ", cuando se presentó un incremento de ", format(uvaa, nsmall = 1), "% ")
    }else if(uvaa < 0){
      descripcion_7 <- paste0(descripcion_7,"la cifra de ", mespal, " de ", año, " fue superior a la de ",
                              mespal, " de ", añoa, ", cuando se presentó una disminución de ", format(abs(uvaa), nsmall = 1), "% ")
    }else if(uvaa == 0){
      descripcion_7 <- paste0(descripcion_7,"la cifra de ", mespal, " de ", año, " fue superior a la de ",
                              mespal, " de ", añoa, ", cuando no se presentó variación ")
    }
  }else if(uv < uvaa){
    descripcion_7 <- paste0(descripcion_7,"el crecimiento de ", mespal, " de ", año, " fue menor al de ",
                            mespal, " de ", añoa, ", cuando se presentó un incremento de ", format(uvaa, nsmall = 1), "% ")
  }else if(uv == uvaa){
    descripcion_7 <- paste0(descripcion_7,"el crecimiento de ", mespal, " de ", año, " fue igual al de ",
                            mespal, " de ", añoa, ", cuando también se presentó un incremento de ", format(uvaa, nsmall = 1), "% ")
  }
}else if(uv < 0){
  if(uv < uvaa){
    if(uvaa > 0){
      descripcion_7 <- paste0(descripcion_7,"la cifra de ", mespal, " de ", año, " fue inferior a la de ",
                              mespal, " de ", añoa, ", cuando se presentó un crecimiento de ", format(uvaa, nsmall = 1), "% ")
    }else if(uvaa < 0){
      descripcion_7 <- paste0(descripcion_7,"la caída de ", mespal, " de ", año, " fue mayor a la de ",
                              mespal, " de ", añoa, ", cuando se presentó una disminución de ", format(abs(uvaa), nsmall = 1), "% ")
    }else if(uvaa == 0){
      descripcion_7 <- paste0(descripcion_7,"la cifra de ", mespal, " de ", año, " fue inferior a la de ",
                              mespal, " de ", añoa, ", cuando no se presentó variación ")
    }
  }else if(uv > uvaa){
    descripcion_7 <- paste0(descripcion_7,"la caída de ", mespal, " de ", año, " fue menor a la de ",
                            mespal, " de ", añoa, ", cuando se presentó una disminución de ", format(abs(uvaa), nsmall = 1), "% ")
  }else if(uv == uvaa){
    descripcion_7 <- paste0(descripcion_7,"la caída de ", mespal, " de ", año, " fue igual a la de ",
                            mespal, " de ", añoa, ", cuando también se presentó una disminución de ", format(abs(uvaa), nsmall = 1), "% ")
  }
}else if(uv == 0){
  if(uv > uvaa){
    descripcion_7 <- paste0(descripcion_7,"la cifra de ", mespal, " de ", año, " fue superior a la de ",
                            mespal, " de ", añoa, ", cuando se presentó una disminución de ", format(abs(uvaa), nsmall = 1), "% ")
  }else if(uv < uvaa){
    descripcion_7 <- paste0(descripcion_7,"la cifra de ", mespal, " de ", año, " fue inferior a la de ",
                            mespal, " de ", añoa, ", cuando se presentó un crecimiento de ", format(uvaa, nsmall = 1), "% ")
  }else if(uv == uvaa){
    descripcion_7 <- paste0(descripcion_7,"la cifra de ", mespal, " de ", año, " fue igual a la de ",
                            mespal, " de ", añoa, ", cuando tampoco se presentó variación ")
  }
}
descripcion_7 <- paste0(descripcion_7,"en la actividad de la industria de la construcción estatal. ")

#Variación estatal promedio
#Conector
if((uv > uvaa & uvp > uvpam) || (uv <= uvaa & uvp <= uvpam)){
  descripcion_7 <- paste0(descripcion_7,"Asimismo, ")
}else{
  descripcion_7 <- paste0(descripcion_7,"Sin embargo, ")
}

descripcion_7 <- paste0(descripcion_7,"la variación estatal promedio de los últimos doce meses ")
if(uvp > uvpam){
  if(uvpam >= 0){
    descripcion_7 <- paste0(descripcion_7,"aumentó ")
  }else if(uvpam < 0){
    descripcion_7 <- paste0(descripcion_7,"cambió ")
  }
  descripcion_7 <- paste0(descripcion_7,"de ",format(uvpam, nsmall = 1), "% a ",format(uvp, nsmall = 1),"%.")
}else if(uvp < uvpam){
  if(uvp >= 0){
    descripcion_7 <- paste0(descripcion_7,"bajó ")
  }else if(uvp < 0){
    descripcion_7 <- paste0(descripcion_7,"cambió ")
  }
  descripcion_7 <- paste0(descripcion_7,"de ",format(uvpam, nsmall = 1), "% a ",format(uvp, nsmall = 1),"%.")
}else if(uvp == uvpam){
  descripcion_7 <- paste0(descripcion_7,"se mantuvo sin cambios en ",format(uvp, nsmall = 1), "%.")
}


# Texto descripcion 8: CON RANK---------------------------------------------------
#Ranking variación porcentual anual de Industria de la Construcción
#df
re1 = filter(graf_rank_con, Entidad != "Nacional")
re1 = arrange(re1, desc(Variacion))
re = filter(mutate(re1, n = 1:nrow(re1), num = num_cardinales[1:nrow(re1)]))#Ranking estados
#Variables
#mespal#Mes de las cifras reportadas
año#Año de las cifras reportadas
añoa <- año-1#Año anterior
nac <- round(filter(graf_rank_con, Entidad == "Nacional")[,2], 1)#Cifra nacional
jal <- round(filter(graf_rank_con, Entidad == "Jalisco")[,2], 1)#cifra jalisco
lgn <- filter(re, Entidad == "Jalisco")[,3]#Lugar de jalisco numero
lgc <- filter(re, Entidad == "Jalisco")[,4]#Lugar de jalisco cardinal


descripcion_8 <- "El desempeño de la actividad de la industria de la construcción estatal en "

descripcion_8 <- paste0(descripcion_8, mespal, " de ", año, " ")

if (jal < nac){
  descripcion_8 <- paste0(descripcion_8, "fue inferior al nacional, ")
  if(nac > 0){
    descripcion_8 <- paste0(descripcion_8, "que registró un crecimiento anual de ", nac, "%. ")   
  }else if(nac < 0){
    descripcion_8 <- paste0(descripcion_8, "que registró una disminución anual de ", abs(nac), "%. ")   
  }else{
    descripcion_8 <- paste0(descripcion_8, "que no registró variación. ")
  }
} else if(jal > nac){
  descripcion_8 <- paste0(descripcion_8, "fue superior al nacional, ")
  if(nac > 0){
    descripcion_8 <- paste0(descripcion_8, "que registró un crecimiento anual de ", nac, "%. ")   
  }else if(nac < 0){
    descripcion_8 <- paste0(descripcion_8, "que registró una disminución anual de ", abs(nac), "%. ")   
  }else{
    descripcion_8 <- paste0(descripcion_8, "que no registró variación. ")
  }
}else{
  descripcion_8 <- paste0(descripcion_8, "estuvo en línea con el nacional, ")
  if(nac > 0){
    descripcion_8 <- paste0(descripcion_8, "que registró un crecimiento anual de ", nac, "%. ")   
  }else if(nac < 0){
    descripcion_8 <- paste0(descripcion_8, "que registró una disminución anual de ", abs(nac), "%. ")   
  }else{
    descripcion_8 <- paste0(descripcion_8, "que no registró variación. ")
  }
}


descripcion_8 <- paste0(descripcion_8, "Jalisco se ubicó en el ")
  if(lgn < 11){
    descripcion_8 <- paste0(descripcion_8, lgc, " lugar a nivel nacional ", "en cuanto a desempeño de la industria de la construcción. ")
  }else{
    descripcion_8 <- paste0(descripcion_8, "lugar ", lgn, " a nivel nacional ",
                            "en cuanto a desempeño de la industria de la construcción. ")
  }



# Texto descripcion 9: MAN VAR---------------------------------------------------
#Variación porcentual anual y variación promedio anual con cifras originales de las Industria Manufactureras

#Variables
#mespal #Mes de las cifras reportadas
#año#Año de las cifras reportadas
añoa <- año-1#Año anterior
uv <- round(graf_var_man[nrow(graf_var_man), 3],1)#ultima variacion
uvam <- round(graf_var_man[nrow(graf_var_man)-1, 3],1)#variacion mes anterior
uvaa <- round(graf_var_man[nrow(graf_var_man)-12, 3],1)#variacion año anterior
uvp <- round(graf_var_man[nrow(graf_var_man), 4],1)#ultima variacion promedio
uvpam <- round(graf_var_man[nrow(graf_var_man)-1, 4],1)#variacion promedio mes anterior

descripcion_9 <- paste0("Respecto a la actividad de las industrias manufactureras en Jalisco, en ", mespal, " de ", año, " ")
#Comparación variación mes anterior:
if (uv > 0) {
  descripcion_9 <- paste0(descripcion_9, "se presentó un crecimiento de ", format(uv, nsmall = 1), "% a tasa anual en ", mespal, " de ", año, ", ")
  if(uv > uvam & uvam > 0){
    descripcion_9 <- paste0(descripcion_9, "incremento superior al del mes inmediato anterior, cuando se presentó un aumento de ", format(abs(uvam), nsmall = 1), "% anual. " )
  }else if(uv < uvam & uvam > 0){
    descripcion_9 <- paste0(descripcion_9, "incremento inferior al del mes inmediato anterior, cuando se presentó un aumento de ", format(abs(uvam), nsmall = 1), "% anual. " )
  }else if(uvam == 0){
    descripcion_9 <- paste0(descripcion_9, "cifra superior a la del mes inmediato anterior, cuando el indicador no presentó variación. " )
  }else if(uv > uvam & uvam < 0){
    descripcion_9 <- paste0(descripcion_9, "cifra superior a la del mes inmediato anterior, cuando se presentó una caída de ", format(abs(uvam), nsmall = 1), "% anual. " )
  }
  
} else if (uv < 0){
  descripcion_9 <- paste0(descripcion_9, "se presentó una disminución de ", format(abs(uv), nsmall = 1), "% a tasa anual en ", mespal, " de ", año, ", ")
  if(uv < uvam & uvam < 0){
    descripcion_9 <- paste0(descripcion_9, "caída mayor a la del mes inmediato anterior, cuando se presentó un descenso de ", format(abs(uvam), nsmall = 1), "% anual. " )
  }else if(uv < uvam & uvam > 0){
    descripcion_9 <- paste0(descripcion_9, "cifra inferior a la del mes inmediato anterior, cuando se presentó un crecimiento de ", format(abs(uvam), nsmall = 1), "% anual. " )
  }else if(uv > uvam & uvam < 0){
    descripcion_9 <- paste0(descripcion_9, "caída menor a la del mes inmediato anterior, cuando se presentó un descenso de ", format(abs(uvam), nsmall = 1), "% anual. " )
  }else if(uvam == 0){
    descripcion_9 <- paste0(descripcion_9, "cifra inferior a la del mes inmediato anterior, cuando el indicador no presentó variación. ")
  }
}else{
  descripcion_9 <- paste0(descripcion_9, "no se presentó variación a tasa anual en ", mespal, " de ", año, ", ")
  if(uvam > 0){
    descripcion_9 <- paste0(descripcion_9, "mientras que el mes inmediato anterior se observó un crecimiento de ",
                            format(uvam, nsmall = 1), "% anual. ")
  }else if(uvam < 0){
    descripcion_9 <- paste0(descripcion_9, "mientras que el mes inmediato anterior se observó una caída de ",
                            format(abs(uvam), nsmall = 1), "% anual. ")
  }else{
    descripcion_9 <- paste0(descripcion_9,"de igual manera que el mes inmediato anterior cuando tampocó se presentaron cambios en el indicador. ")
  }
}

#Comparación variación año anterior:
#Conector
if((uv > uvam & uv > uvaa) || (uv <= uvam & uv <= uvaa)){
  descripcion_9 <- paste0(descripcion_9,"Además, ")
}else{
  descripcion_9 <- paste0(descripcion_9,"Sin embargo, ")
}

if(uv > 0){
  if(uv > uvaa){
    if(uvaa > 0){
      descripcion_9 <- paste0(descripcion_9,"el crecimiento de ", mespal, " de ", año, " fue mayor al de ",
                              mespal, " de ", añoa, ", cuando se presentó un incremento de ", format(uvaa, nsmall = 1), "% ")
    }else if(uvaa < 0){
      descripcion_9 <- paste0(descripcion_9,"la cifra de ", mespal, " de ", año, " fue superior a la de ",
                              mespal, " de ", añoa, ", cuando se presentó una disminución de ", format(abs(uvaa), nsmall = 1), "% ")
    }else if(uvaa == 0){
      descripcion_9 <- paste0(descripcion_9,"la cifra de ", mespal, " de ", año, " fue superior a la de ",
                              mespal, " de ", añoa, ", cuando no se presentó variación ")
    }
  }else if(uv < uvaa){
    descripcion_9 <- paste0(descripcion_9,"el crecimiento de ", mespal, " de ", año, " fue menor al de ",
                            mespal, " de ", añoa, ", cuando se presentó un incremento de ", format(uvaa, nsmall = 1), "% ")
  }else if(uv == uvaa){
    descripcion_9 <- paste0(descripcion_9,"el crecimiento de ", mespal, " de ", año, " fue igual al de ",
                            mespal, " de ", añoa, ", cuando también se presentó un incremento de ", format(uvaa, nsmall = 1), "% ")
  }
}else if(uv < 0){
  if(uv < uvaa){
    if(uvaa > 0){
      descripcion_9 <- paste0(descripcion_9,"la cifra de ", mespal, " de ", año, " fue inferior a la de ",
                              mespal, " de ", añoa, ", cuando se presentó un crecimiento de ", format(uvaa, nsmall = 1), "% ")
    }else if(uvaa < 0){
      descripcion_9 <- paste0(descripcion_9,"la caída de ", mespal, " de ", año, " fue mayor a la de ",
                              mespal, " de ", añoa, ", cuando se presentó una disminución de ", format(abs(uvaa), nsmall = 1), "% ")
    }else if(uvaa == 0){
      descripcion_9 <- paste0(descripcion_9,"la cifra de ", mespal, " de ", año, " fue inferior a la de ",
                              mespal, " de ", añoa, ", cuando no se presentó variación ")
    }
  }else if(uv > uvaa){
    descripcion_9 <- paste0(descripcion_9,"la caída de ", mespal, " de ", año, " fue menor a la de ",
                            mespal, " de ", añoa, ", cuando se presentó una disminución de ", format(abs(uvaa), nsmall = 1), "% ")
  }else if(uv == uvaa){
    descripcion_9 <- paste0(descripcion_9,"la caída de ", mespal, " de ", año, " fue igual a la de ",
                            mespal, " de ", añoa, ", cuando también se presentó una disminución de ", format(abs(uvaa), nsmall = 1), "% ")
  }
}else if(uv == 0){
  if(uv > uvaa){
    descripcion_9 <- paste0(descripcion_9,"la cifra de ", mespal, " de ", año, " fue superior a la de ",
                            mespal, " de ", añoa, ", cuando se presentó una disminución de ", format(abs(uvaa), nsmall = 1), "% ")
  }else if(uv < uvaa){
    descripcion_9 <- paste0(descripcion_9,"la cifra de ", mespal, " de ", año, " fue inferior a la de ",
                            mespal, " de ", añoa, ", cuando se presentó un crecimiento de ", format(uvaa, nsmall = 1), "% ")
  }else if(uv == uvaa){
    descripcion_9 <- paste0(descripcion_9,"la cifra de ", mespal, " de ", año, " fue igual a la de ",
                            mespal, " de ", añoa, ", cuando tampoco se presentó variación ")
  }
}
descripcion_9 <- paste0(descripcion_9,"en la actividad de las industrias manufactureras. ")

#Variación estatal promedio
#Conector
if((uv > uvaa & uvp > uvpam) || (uv <= uvaa & uvp <= uvpam)){
  descripcion_9 <- paste0(descripcion_9,"Asimismo, ")
}else{
  descripcion_9 <- paste0(descripcion_9,"Sin embargo, ")
}

descripcion_9 <- paste0(descripcion_9,"la variación estatal promedio de los últimos doce meses ")
if(uvp > uvpam){
  if(uvpam >= 0){
    descripcion_9 <- paste0(descripcion_9,"aumentó ")
  }else if(uvpam < 0){
    descripcion_9 <- paste0(descripcion_9,"cambió ")
  }
  descripcion_9 <- paste0(descripcion_9,"de ",format(uvpam, nsmall = 1), "% a ",format(uvp, nsmall = 1),"%.")
}else if(uvp < uvpam){
  if(uvp >= 0){
    descripcion_9 <- paste0(descripcion_9,"bajó ")
  }else if(uvp < 0){
    descripcion_9 <- paste0(descripcion_9,"cambió ")
  }
  descripcion_9 <- paste0(descripcion_9,"de ",format(uvpam, nsmall = 1), "% a ",format(uvp, nsmall = 1),"%.")
}else if(uvp == uvpam){
  descripcion_9 <- paste0(descripcion_9,"se mantuvo sin cambios en ",format(uvp, nsmall = 1), "%.")
}



# Texto descripcion 10: MAN RANK--------------------------------------------------
#Ranking variación porcentual anual de Industrias Manufactureras
#df
re1 = filter(graf_rank_man, Entidad != "Nacional")
re1 = arrange(re1, desc(Variacion))
re = filter(mutate(re1, n = 1:nrow(re1), num = num_cardinales[1:nrow(re1)]))#Ranking estados
#Variables
#mespal#Mes de las cifras reportadas
#año#Año de las cifras reportadas
añoa <- año-1#Año anterior
nac <- round(filter(graf_rank_man, Entidad == "Nacional")[,2], 1)#Cifra nacional
jal <- round(filter(graf_rank_man, Entidad == "Jalisco")[,2], 1)#cifra jalisco
lgn <- filter(re, Entidad == "Jalisco")[,3]#Lugar de jalisco numero
lgc <- filter(re, Entidad == "Jalisco")[,4]#Lugar de jalisco cardinal


descripcion_10 <- "El desempeño estatal de la actividad de las industrias manufactureras, "

if (jal > 0){
  descripcion_10 <- paste0(descripcion_10, "que mostró un crecimiento de ", jal, "% anual en ",
                           mespal, ", ")
}else if (jal < 0){
  descripcion_10 <- paste0(descripcion_10, "que mostró una caída de ", abs(jal), "% anual en ",
                           mespal, ", ")
}else{
  descripcion_10 <- paste0(descripcion_10, "que no registró variación, ")
}

if (jal < nac){
  descripcion_10 <- paste0(descripcion_10, "fue inferior al nacional, ")
  if(nac > 0){
    descripcion_10 <- paste0(descripcion_10, "que registró un crecimiento anual de ", nac, "%. ")   
  }else if(nac < 0){
    descripcion_10 <- paste0(descripcion_10, "que registró una disminución anual de ", abs(nac), "%. ")   
  }else{
    descripcion_10 <- paste0(descripcion_10, "que no registró variación. ")
  }
} else if(jal > nac){
  descripcion_10 <- paste0(descripcion_10, "fue superior al nacional, ")
  if(nac > 0){
    descripcion_10 <- paste0(descripcion_10, "que registró un crecimiento anual de ", nac, "%. ")   
  }else if(nac < 0){
    descripcion_10 <- paste0(descripcion_10, "que registró una disminución anual de ", abs(nac), "%. ")   
  }else{
    descripcion_10 <- paste0(descripcion_10, "que no registró variación. ")
  }
}else{
  descripcion_10 <- paste0(descripcion_10, "estuvo en línea con el nacional, ")
  if(nac > 0){
    descripcion_10 <- paste0(descripcion_10, "que registró un crecimiento anual de ", nac, "%. ")   
  }else if(nac < 0){
    descripcion_10 <- paste0(descripcion_10, "que registró una disminución anual de ", abs(nac), "%. ")   
  }else{
    descripcion_10 <- paste0(descripcion_10, "que no registró variación. ")
  }
}


descripcion_10 <- paste0(descripcion_10, "Jalisco se ubicó en el ")
if(lgn < 11){
  descripcion_10 <- paste0(descripcion_10, lgc, " lugar a nivel nacional ", "en cuanto a desempeño de las industrias manufactureras. ")
}else{
  descripcion_10 <- paste0(descripcion_10, "lugar ", lgn, " a nivel nacional ",
                          "en cuanto a desempeño de las industrias manufactureras. ")
}





########## EXPORTAR A EXCEL ##########
#Fuente y titulos
ft=data.frame(c("Fuente: IIEG, con información de INEGI.",
                "Indicador Mensual de la Actividad Industrial de Jalisco. Variación porcentual anual y variación promedio anual con cifras originales,",
                "Variación porcentual anual del IMAIEF por entidad federativa con cifras originales,",
                "Indicador Mensual de la Actividad Industrial de Jalisco. Serie desestacionalizadas y de tendencia-ciclo,",
                "Variación porcentual anual del IMAIEF por entidad federativa con cifras desestacionalizadas,",
                "Variación porcentual mensual del IMAIEF por entidad federativa con cifras desestacionalizadas,",
                "Variación porcentual anual del IMAIEF de actividades secundarias, industria de la construcción, industrias manufactureras, minería y servicios públicos de Jalisco,",
                "Variación porcentual anual del IMAIEF de construcción de Jalisco y su promedio de últimos doce meses,",
                "Variación porcentual anual del IMAIEF de construcción por entidad federativa,",
                "Variación porcentual anual del IMAIEF de industrias manufactureras de Jalisco y su promedio de últimos doce meses,",
                "Variación porcentual anual del IMAIEF de industrias manufactureras por entidad federativa,"))
colnames(ft)= "FT"


#Nombres de variables
nombre_var=data.frame(c("Año","Mes","Variación","Variación promedio","Entidad"))
colnames(nombre_var)="VARIABLE"
names(graf_var_as)=nombre_var[1:4,1]
names(graf_rank_as)=nombre_var[c(5,3),1]
names(graf_comp)[2]=nombre_var[3,1]
names(graf_var_con)=nombre_var[1:4,1]
names(graf_rank_con)=nombre_var[c(5,3),1]
names(graf_var_man)=nombre_var[1:4,1]
names(graf_rank_man)=nombre_var[c(5,3),1]

#Notas de graficas
notas=data.frame(c("Nota: La variación anual es la variación con respecto al mismo mes del año anterior, mientras que la variación promedio es el promedio de las variaciones los últimos doce meses. Las variaciones son de las cifras originales, es decir, las cifras sin desestacionalizar.",
                 "Nota: La variación anual es la variación con respecto al mismo mes del año anterior. Las variaciones son de las cifras originales, es decir, las cifras sin desestacionalizar.",
                 "Nota: Índice base 2018=100.",
                 "Nota: La variación anual es la variación con respecto al mismo mes del año anterior. Las variaciones son de las cifras desestacionalizadas.",
                 "Nota: La variación mensual es la variación con respecto al mes inmediato anterior. Las variaciones son de las cifras desestacionalizadas.",
                 "Nota: La variación anual es la variación con respecto al mismo mes del año anterior. La variación de servicios públicos se refiere a la generación, transmisión y distribución de energía eléctrica, suministro de agua y de gas por ductos al consumidor final. Las variaciones son de las cifras originales, es decir, las cifras sin desestacionalizar."
                 ))
colnames(notas) = "Notas"


wb=createWorkbook("IIEG DIEEF")
addWorksheet(wb, "ACT SEC VAR")
titulo=paste(ft[2,1],periodo1)
writeData(wb, sheet=1, titulo, startCol=1, startRow=1)
writeData(wb, sheet=1, ft[1,1], startCol=1, startRow=2)
writeData(wb, sheet=1, graf_var_as, startCol=1, startRow=5)

addWorksheet(wb, "ACT SEC RANK")
titulo=paste(ft[3,1],periodo2)
writeData(wb, sheet=2, titulo, startCol=1, startRow=1)
writeData(wb, sheet=2, ft[1,1], startCol=1, startRow=2)
writeData(wb, sheet=2, graf_rank_as, startCol=1, startRow=5)

addWorksheet(wb, "DESYTC")
titulo=paste(ft[4,1],periodo2)
writeData(wb, sheet=3, titulo, startCol=1, startRow=1)
writeData(wb, sheet=3, ft[1,1], startCol=1, startRow=2)
writeData(wb, sheet=3, desytc, startCol=1, startRow=5)

addWorksheet(wb, "DESEST RANK ANU")
titulo=paste(ft[5,1],periodo2)
writeData(wb, sheet=4, titulo, startCol=1, startRow=1)
writeData(wb, sheet=4, ft[1,1], startCol=1, startRow=2)
writeData(wb, sheet=4, EntidadClaveVA, startCol=1, startRow=5)

addWorksheet(wb, "DESEST RANK MEN")
titulo=paste(ft[6,1],periodo2)
writeData(wb, sheet=5, titulo, startCol=1, startRow=1)
writeData(wb, sheet=5, ft[1,1], startCol=1, startRow=2)
writeData(wb, sheet=5, EntidadClaveVM, startCol=1, startRow=5)

addWorksheet(wb, "ACT COM")
titulo=paste(ft[7,1],periodo2)
writeData(wb, sheet=6, titulo, startCol=1, startRow=1)
writeData(wb, sheet=6, ft[1,1], startCol=1, startRow=2)
writeData(wb, sheet=6, graf_comp, startCol=1, startRow=5)

addWorksheet(wb, "CON VAR")
titulo=paste(ft[8,1],periodo1)
writeData(wb, sheet=7, titulo, startCol=1, startRow=1)
writeData(wb, sheet=7, ft[1,1], startCol=1, startRow=2)
writeData(wb, sheet=7, graf_var_con, startCol=1, startRow=5)

addWorksheet(wb, "CON RANK")
titulo=paste(ft[9,1],periodo2)
writeData(wb, sheet=8, titulo, startCol=1, startRow=1)
writeData(wb, sheet=8, ft[1,1], startCol=1, startRow=2)
writeData(wb, sheet=8, graf_rank_con, startCol=1, startRow=5)

addWorksheet(wb, "MAN VAR")
titulo=paste(ft[10,1],periodo1)
writeData(wb, sheet=9, titulo, startCol=1, startRow=1)
writeData(wb, sheet=9, ft[1,1], startCol=1, startRow=2)
writeData(wb, sheet=9, graf_var_man, startCol=1, startRow=5)

addWorksheet(wb, "MAN RANK")
titulo=paste(ft[11,1],periodo2)
writeData(wb, sheet=10, titulo, startCol=1, startRow=1)
writeData(wb, sheet=10, ft[1,1], startCol=1, startRow=2)
writeData(wb, sheet=10, graf_rank_man, startCol=1, startRow=5)



addWorksheet(wb, "Texto")
titulo=paste0("Actividad industrial de Jalisco en ",mespal," de ", año)
writeData(wb, sheet = 11, "Título ficha:", startCol = 1, startRow = 1)
writeData(wb, sheet = 11, titulo, startCol = 1, startRow = 2)
###
writeData(wb, sheet = 11, "ACT SEC VAR:", startCol = 1, startRow = 4)
writeData(wb, sheet = 11, "Texto1_1:", startCol = 1, startRow = 5)
writeData(wb, sheet = 11, "Texto1_2:", startCol = 1, startRow = 6)
writeData(wb, sheet = 11, "Texto1_3:", startCol = 1, startRow = 7)
writeData(wb, sheet = 11, "Gráfica:", startCol = 1, startRow = 8)
writeData(wb, sheet = 11, "Nota:", startCol = 1, startRow = 9)
writeData(wb, sheet = 11, descripcion_1, startCol = 2, startRow = 5)
writeData(wb, sheet = 11, descripcion_1_2, startCol = 2, startRow = 6)
writeData(wb, sheet = 11, descripcion_1_3, startCol = 2, startRow = 7)
writeData(wb, sheet = 11, paste(ft[2,1],periodo1), startCol = 2, startRow = 8)
writeData(wb, sheet = 11, notas[1,1], startCol = 2, startRow = 9)
###
writeData(wb, sheet = 11, "ACT SEC RANK:", startCol = 1, startRow = 11)
writeData(wb, sheet = 11, "Texto_2:", startCol = 1, startRow = 12)
writeData(wb, sheet = 11, "Gráfica:", startCol = 1, startRow = 13)
writeData(wb, sheet = 11, "Nota:", startCol = 1, startRow = 14)
writeData(wb, sheet = 11, descripcion_2, startCol = 2, startRow = 12)
writeData(wb, sheet = 11, paste(ft[3,1],periodo2), startCol = 2, startRow = 13)
writeData(wb, sheet = 11, notas[2,1], startCol = 2, startRow = 14)
###
writeData(wb, sheet = 11, "DESYTC:", startCol = 1, startRow = 16)
writeData(wb, sheet = 11, "Texto_3:", startCol = 1, startRow = 17)
writeData(wb, sheet = 11, "Gráfica:", startCol = 1, startRow = 18)
writeData(wb, sheet = 11, "Nota:", startCol = 1, startRow = 19)
writeData(wb, sheet = 11, descripcion_3, startCol = 2, startRow = 17)
writeData(wb, sheet = 11, paste(ft[4,1],periodo2), startCol = 2, startRow = 18)
writeData(wb, sheet = 11, notas[3,1], startCol = 2, startRow = 19)
###
writeData(wb, sheet = 11, "DESEST RANK ANU:", startCol = 1, startRow = 21)
writeData(wb, sheet = 11, "Texto_4:", startCol = 1, startRow = 22)
writeData(wb, sheet = 11, "Gráfica:", startCol = 1, startRow = 23)
writeData(wb, sheet = 11, "Nota:", startCol = 1, startRow = 24)
writeData(wb, sheet = 11, descripcion_4, startCol = 2, startRow = 22)
writeData(wb, sheet = 11, paste(ft[5,1],periodo2), startCol = 2, startRow = 23)
writeData(wb, sheet = 11, notas[4,1], startCol = 2, startRow = 24)
###
writeData(wb, sheet = 11, "DESEST RANK MEN:", startCol = 1, startRow = 26)
writeData(wb, sheet = 11, "Texto_5:", startCol = 1, startRow = 27)
writeData(wb, sheet = 11, "Gráfica:", startCol = 1, startRow = 28)
writeData(wb, sheet = 11, "Nota:", startCol = 1, startRow = 29)
writeData(wb, sheet = 11, descripcion_5, startCol = 2, startRow = 27)
writeData(wb, sheet = 11, paste(ft[6,1],periodo2), startCol = 2, startRow = 28)
writeData(wb, sheet = 11, notas[5,1], startCol = 2, startRow = 29)
###
writeData(wb, sheet = 11, "ACT COM:", startCol = 1, startRow = 31)
writeData(wb, sheet = 11, "Texto_6:", startCol = 1, startRow = 32)
writeData(wb, sheet = 11, "Gráfica:", startCol = 1, startRow = 33)
writeData(wb, sheet = 11, "Nota:", startCol = 1, startRow = 34)
writeData(wb, sheet = 11, descripcion_6, startCol = 2, startRow = 32)
writeData(wb, sheet = 11, paste(ft[7,1],periodo2), startCol = 2, startRow = 33)
writeData(wb, sheet = 11, notas[6,1], startCol = 2, startRow = 34)
###
writeData(wb, sheet = 11, "CON VAR:", startCol = 1, startRow = 36)
writeData(wb, sheet = 11, "Texto_7:", startCol = 1, startRow = 37)
writeData(wb, sheet = 11, "Gráfica:", startCol = 1, startRow = 38)
writeData(wb, sheet = 11, "Nota:", startCol = 1, startRow = 39)
writeData(wb, sheet = 11, descripcion_7, startCol = 2, startRow = 37)
writeData(wb, sheet = 11, paste(ft[8,1],periodo1), startCol = 2, startRow = 38)
writeData(wb, sheet = 11, notas[1,1], startCol = 2, startRow = 39)
###
writeData(wb, sheet = 11, "CON RANK:", startCol = 1, startRow = 41)
writeData(wb, sheet = 11, "Texto_8:", startCol = 1, startRow = 42)
writeData(wb, sheet = 11, "Gráfica:", startCol = 1, startRow = 43)
writeData(wb, sheet = 11, "Nota:", startCol = 1, startRow = 44)
writeData(wb, sheet = 11, descripcion_8, startCol = 2, startRow = 42)
writeData(wb, sheet = 11, paste(ft[9,1],periodo2), startCol = 2, startRow = 43)
writeData(wb, sheet = 11, notas[2,1], startCol = 2, startRow = 44)
###
writeData(wb, sheet = 11, "MAN VAR:", startCol = 1, startRow = 46)
writeData(wb, sheet = 11, "Texto_9:", startCol = 1, startRow = 47)
writeData(wb, sheet = 11, "Gráfica:", startCol = 1, startRow = 48)
writeData(wb, sheet = 11, "Nota:", startCol = 1, startRow = 49)
writeData(wb, sheet = 11, descripcion_9, startCol = 2, startRow = 47)
writeData(wb, sheet = 11, paste(ft[10,1],periodo1), startCol = 2, startRow = 48)
writeData(wb, sheet = 11, notas[1,1], startCol = 2, startRow = 49)
###
writeData(wb, sheet = 11, "MAN RANK:", startCol = 1, startRow = 51)
writeData(wb, sheet = 11, "Texto_10:", startCol = 1, startRow = 52)
writeData(wb, sheet = 11, "Gráfica:", startCol = 1, startRow = 53)
writeData(wb, sheet = 11, "Nota:", startCol = 1, startRow = 54)
writeData(wb, sheet = 11, descripcion_10, startCol = 2, startRow = 52)
writeData(wb, sheet = 11, paste(ft[11,1],periodo2), startCol = 2, startRow = 53)
writeData(wb, sheet = 11, notas[2,1], startCol = 2, startRow = 54)
###

nombre_wb=paste0("IMAIEF_R-Excel ", fcsv, ".xlsx")
saveWorkbook(wb, nombre_wb, overwrite = TRUE)

