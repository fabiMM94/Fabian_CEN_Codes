>>>>>>>>>>>>>PARA IMPORTAR DATOS:

-Van a la carpeta DATA_import
-Deje 3 archivos importes LFData.xlsx  PVData.xlsx  WPData.xlsx
-En ellos veran dos columnas una con el nombre y otra con la potencia
- Para los nombres les sigiero utilizar los nombres de las centrales Encendididas extraidas del codigo RUN_States.dwj 
- Una vez tengas listos los archivos .xlsx usan el boton.bat y se automaticamente se guardan como csv
- El codigo leerá esos CSV generados    
- finalmente, usan el boton Run_import.dwj en la base de EMTP para ejecutar


>>>>>>>>>>>>PARA Exportar datos: 

- Solo ejecuten el boton Run_export.dwj
- Eso generará csv en la carpeta Data_export 
- alli veran cvs para cada tipo de central


>>>>>>>>>>>>ANTES DE TODO: 

Recuerden cambiar el pathData1 segun la carpeta donde exista este github en su PC
 Y que al mismo tiempo este el archivo EMTP que trabajen