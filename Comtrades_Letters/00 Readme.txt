URL Carta DE03007-25:
https://correspondencia.coordinador.cl/correspondencia/show/envio/682c6ee93563574f3c9fb8f6

======================================================================================================================


***Extractor : 

1) Iniciar sesion manualmente en el sist de correspondencia 
2) Una vez inciada va directamente al url de la carta asociada 
3) Se abre la url con todas las cartas respondidas y me pregunta cuantas cartas quiere que analice
4) Se abre cada enlance, recopila informacion, descarga la carta y los anexos si esque los tiene y los guarda
en las carpeta de anexos y cartas
5) Genera el Excel

***Post Proceso: 

Como el excel generado tiene mucha informacion hago una limpieza:

1) Tomo la carpeta de anexo con las descargas
2) Tomo el excel generado
3) Recorro el excel y busco si existe el anexo asociado a ese correlativo, donde si tiene evaluo si tiene Comtrade
4) Genera un nuevo excel mas limpio con el checkeo de cada anexo, con el resultado de la evulación si tiene Comtrade o no



========================================================================================================================

****Versiones 

Extractor 3: Este tiene incluido los filtro de columnas del excel exportado
Extractor 4: Este tiene incluido las descargas de cartas y anexos
Extractor 5 : se corrige el nombre de la carta descargada que tenia antes un string adicional
EXttractor6 : Se corrige las extensiones de los anexos, los cuales tenian un .pdf alfinal en algunas veces
Exctactor7 : Aqui incluyo las subcarpetas de anexos


PostProcess2 : corrige el error de multiples archivos sueltos y .7z que no los checkeaba
PostProcess3 : Nuevo Formato tipo SQL 
PostProcces4: Cambia el formato para que en hoja Anexos tambien tenga la info de la Subcarpeta