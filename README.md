# Lector Excel Conciliación
Aplicación para leer los extractos de los bancos, subirlos al sistema SISGO (a la tabla ARCHIVOSCONCIBANCATMP) y llamar a los procesos correspondientes (PKG_CARGARARCHIVOSAUTO).

Trabaja los archivos en la CARPETAWORK, en su subcarpeta correspondiente (según PID). Finalizado el proceso, lo mueve a la CARPETAOUTPUT, en subcarpetas según fecha de proceso.

Puede trabajar con múltiples instancias.


### Uso

```powershell
.\LectorExcelConciliacion.exe <-nopause> <-killall> <-file (rutacompleta)>
  -nopause: Finaliza la aplicación al terminar las operaciones. Por defecto, pausa la aplicación.
  -killall: Termina todos los procesos en ejecución de nombre LectorExcelConciliacion.exe.
  -file (rutacompleta): Procesa el archivo especifico.
  Sin parámetros: Procesa todos los archivos en la carpeta INPUT.
```

## Requerimientos
 - Microsoft Office

## Antes de copilar
 - Modificar la variable "ambiente" dentro de "Parameters.cs" según corresponda.

### App.config
 - Se incluye un App.config dummy, las configuraciones SISGO no están en este repositorio.
Solicitarlo al responsable o editar con los valores correspondientes dentro del tag "connectionStrings".

## Variable en Base de Datos
```SQL
SELECT TBLDETALLE FROM SYST900 WHERE TBLCODTAB = 50 AND TBLCODARG = 14 --CARPETAINPUT
SELECT TBLDETALLE FROM SYST900 WHERE TBLCODTAB = 50 AND TBLCODARG = 15 --CARPETAOUTPUT
SELECT TBLDETALLE FROM SYST900 WHERE TBLCODTAB = 50 AND TBLCODARG = 29 --CARPETAWORK
```

## Pendiente
 - Log.
 - Validar si hay algún problema al cerrar el Excel tras la lectura y subida a la tabla ARCHIVOSCONCIBANCATMP, potencialmente el usuario puede mover el archivo y la aplicación no lo encontrara tras el proceso (instruir?).

## Otros
 - PKG_CARGARARCHIVOSAUTO.P_GEN_CARGABANCOS_CAJA demora.
