# Lector Excel Conciliación
Aplicación para leer los extractos de los bancos, subirlos al sistema SISGO (a la tabla ARCHIVOSCONCIBANCATMP) y llamar a los procesos correspondientes (PKG_CARGARARCHIVOSAUTO).

Trabaja los archivos en la CARPETAWORK, en su subcarpeta correspondiente (según PID). Finalizado el proceso, lo mueve a la CARPETAOUTPUT, en subcarpetas según fecha de proceso.

Puede trabajar con múltiples instancias.


### Uso

```powershell
.\LectorExcelConciliacion.exe <-nopause> <-hide> <-killall> <-file (rutacompleta) ... -file (rutacompleta)>
  -nopause: Finaliza la aplicación al terminar las operaciones. Por defecto, pausa la aplicación.
  -hide: Oculta la consola
  -killall: Termina todos los procesos en ejecución de nombre LectorExcelConciliacion.exe.
  -file (rutacompleta): Procesa el archivo especifico.
  Sin parametro -file: Procesa todos los archivos en la carpeta INPUT
```

## Requerimientos
 - Microsoft Office

### connString.cs
 - Se incluye un connString.cs dummy, las cadenas ODBC connection para SISGO no están en este repositorio.
Solicitarlo al responsable o editar con los valores correspondientes en el dummy.

## Variable en Base de Datos
```SQL
SELECT TBLDETALLE FROM SYST900 WHERE TBLCODTAB = 50 AND TBLCODARG = 14 --CARPETAINPUT
SELECT TBLDETALLE FROM SYST900 WHERE TBLCODTAB = 50 AND TBLCODARG = 15 --CARPETAOUTPUT
SELECT TBLDETALLE FROM SYST900 WHERE TBLCODTAB = 50 AND TBLCODARG = 29 --CARPETAWORK
```

## Pendiente
 - Log.

## Otros
 - PKG_CARGARARCHIVOSAUTO.P_GEN_CARGABANCOS_CAJA demora.
