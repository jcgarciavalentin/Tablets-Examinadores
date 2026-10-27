# Gestión de tablets (Microsoft Access)

Base de datos de Access (`GestionTablets.accdb`) para llevar el control de examinadores, tablets y
tarjetas SIM: qué SIM tiene cada tablet, qué tablet tiene cada examinador, y el histórico de todo ello.

## Puesta en marcha (una sola vez)

1. Descarga el proyecto (en GitHub: **Code → Download ZIP**), descomprímelo y copia la carpeta
   completa a un equipo Windows con Microsoft Access. Si Windows bloquea los scripts, antes de
   descomprimir haz clic derecho sobre el ZIP → **Propiedades** → **Desbloquear**.
2. Haz doble clic en **`ConstruirFormularios.vbs`** con la base de datos cerrada. Te pregunta si
   quieres importar los datos de la base antigua (ver más abajo). Después abre Access, crea las
   consultas, los formularios y el menú dentro de `GestionTablets.accdb` y se cierra. Tarda unos
   segundos.
3. Abre `GestionTablets.accdb`. Si aparece la barra amarilla de seguridad, pulsa **Habilitar
   contenido**. También puedes guardar la carpeta en una *ubicación de confianza* (Archivo → Opciones
   → Centro de confianza) para que no vuelva a salir. Sin habilitar el contenido, los botones no
   funcionan.

Se abre el menú principal (`frmMenu`).

> **¿Por qué hace falta el paso 2?** Las tablas y relaciones se generan con Jackcess, una librería
> Java que escribe ficheros de Access sin necesitar Access. Pero los formularios de Access solo los
> puede crear el propio Access, así que el script lo hace con su automatización (`CreateForm`).
>
> Si el `.vbs` está bloqueado en tu equipo, hazlo a mano: abre la base de datos, pulsa `Alt+F11`,
> usa *Archivo → Importar archivo* para importar los dos `.bas` de la carpeta `vba`, abre la ventana
> Inmediato (`Ctrl+G`), escribe `ConstruirInterfaz` y pulsa Intro. Después borra el módulo
> importado de `modConstruirInterfaz.bas`, que ya no hace falta.

## Aplicar los últimos cambios

Cada vez que haya una versión nueva del proyecto, descárgala y ejecuta
**`aplicar_ultimos_cambios.vbs`** sobre la base de datos con la que trabajas, con la base cerrada:

- **Arrastra** tu base de datos sobre el script, o haz **doble clic** y escribe su ruta. La recuerda
  para la próxima vez en `ruta_base_datos.txt`.
- Añade las tablas y campos nuevos que falten y rehace consultas, formularios y código.
- El campo **Foto** de las SIM lo crea siempre este paso, porque Jackcess no sabe crear campos de
  datos adjuntos.

Los dos scripts, `ConstruirFormularios.vbs` (creación inicial, sobre el `GestionTablets.accdb` de su
carpeta) y `aplicar_ultimos_cambios.vbs` (sobre tu base de datos, esté donde esté), hacen lo mismo:

1. **Copia de seguridad** de la base de datos en la carpeta `copias`, junto a ella, con la fecha y
   la hora en el nombre.
2. Si la base tiene datos, pregunta **si conservarlos**:
   - **Sí**: se aplican los cambios sin tocar los datos.
   - **No**: se borran examinadores, tablets, SIM y asignaciones.
   - **Cancelar**: no se hace nada.
3. Si no se conservan los datos, o la base está vacía, pregunta si **importar ahora la base
   antigua**. La eliges en el selector de archivos de Windows, que se abre en la que elegiste la
   última vez (guardada en `ruta_base_antigua.txt`). Si dices que no, la base queda vacía.
4. Aplica los cambios y muestra un resumen.

## Cargar los datos de la base de datos antigua

La importación copia los datos de la base antigua (la de las tablas `T-Examinadores`, `T-Tablets`,
etc.) a la base nueva. Se hace con cualquiera de los dos scripts anteriores, respondiendo **No** a
conservar los datos y **Sí** a importar. Las dos bases de datos deben estar cerradas.

- La base antigua solo se lee y no se modifica.
- Al terminar sale un resumen. Si hay incidencias, abre la tabla **`ImportacionIncidencias`** de la
  base nueva.
- Se puede repetir cuantas veces haga falta, por ejemplo para hacer pruebas con los datos actuales,
  o tras corregir en la base antigua lo que salga en las incidencias.
- **Se pierde lo hecho en la base nueva desde la importación anterior.** La copia de seguridad del
  paso 1 lo conserva por si acaso.

| Base antigua | Base nueva |
|---|---|
| `TSituacionExaminador`, `T-EstadoTablet`, `T-EstadoSIM` | Se añaden a `EstadosExaminador` / `EstadosTablet` / `EstadosSIM` los valores que falten. |
| `T-Examinadores` | `Examinadores`. El DNI se guarda en mayúsculas y sin espacios, guiones ni puntos. Sin situación → *Activo*. |
| `T-Tablets` | `Tablets`. Las fechas, que en la antigua son texto, se convierten. Sin estado → *Operativa*. |
| `T-HistorialAsigTabletExaminador` y `T-HistoricoTablet` | `AsignacionesTablet`. Se enlazan por DNI y nº de serie, y la misma asignación en ambas tablas se carga una sola vez. El antiguo estado de la asignación pasa a las observaciones. |
| `T-HistoriaTabletSim` | Una fila en `TarjetasSIM` por cada número de SIM distinto (campo SIM → IMEI), más su historial en `AsignacionesSIM`. PIN, PUK y teléfono asociado se toman de la primera fila que los tenga. El estado es el de la activación más reciente; además, el estado de cada activación queda en sus observaciones. |
| `T-Accesorios`, `Hoja1`, `Copia de …`, `Errores de pegado`, `Hoja1$_ErroresDeImportación` | No se importan. |

Qué se anota en `ImportacionIncidencias`:

- **Registros no importados**: sin DNI o nº de serie, DNI de más de 9 caracteres, duplicados,
  asignaciones cuya tablet o examinador no existe, y asignaciones sin fecha de alta válida.
- **Datos corregidos**:
  - IMEI repetidos: se pasan a observaciones, porque en la base nueva son únicos.
  - Fechas que no se entienden: el texto original se guarda en observaciones.
  - Nombre o apellidos vacíos: se rellenan con "(sin dato)".
  - Textos demasiado largos: se recortan, y el texto completo queda en la incidencia.
- **Asignaciones vigentes a la vez**: por ejemplo, una tablet con dos examinadores sin fecha de baja.
  Hay que cerrar a mano la que sobre poniendo su fecha de fin.

Formatos de fecha que entiende: 15/03/2023, 15-03-2023, 15.03.2023 y los números de fecha de Excel.

## Tablas

| Tabla | Campos |
|---|---|
| `Examinadores` | IdExaminador, NumeroExaminador, Nombre, Apellidos, DNI, IdEstadoExam (situación), Observaciones |
| `EstadosExaminador` | IdEstadoExam, EstadoExaminador: *Activo, Baja temporal, Baja* |
| `Tablets` | IdTablet, IMEI1, IMEI2, NumeroSerie, Marca, Modelo, FechaRecepcion, FechaDesafectacion, IdEstado, Observaciones, Incidencias |
| `EstadosTablet` | IdEstado, Estado: *Operativa, Averiada, En reparación, Desafectada* |
| `TarjetasSIM` | IdSIM, IMEI, NumeroTelefono (teléfono asociado), IdCompania (compañía, por defecto N/A), PIN, PUK, IdEstadoSIM (estado), Observaciones, Foto (datos adjuntos) |
| `Companias` | IdCompania, Compania: *N/A, Movistar, Orange, Vodafone* |
| `EstadosSIM` | IdEstadoSIM, EstadoSIM: *Activa, Suspendida, Baja* |
| `AsignacionesTablet` | IdAsignacion, IdExaminador, IdTablet, FechaInicio, FechaFin, Observaciones |
| `AsignacionesSIM` | IdAsignacionSIM, IdTablet, IdSIM, FechaInicio, FechaFin, Observaciones |
| `HistorialEstadosTablet` | IdCambio, IdTablet, Fecha, IdEstadoAnterior, IdEstadoNuevo, IdExaminador (quién la devolvió), Observaciones |

- Los valores iniciales de las tablas de estados son orientativos. Se cambian desde el menú
  (*Estados de tablet* / *Situaciones de examinador* / *Estados de SIM* / *Compañías*).
- Son obligatorios: nombre, apellidos y DNI del examinador, nº de serie de la tablet e IMEI de la SIM.
- Son únicos: DNI, nº de examinador, nº de serie, IMEI1, IMEI2 e IMEI de la SIM. Los campos
  opcionales pueden quedar vacíos en varios registros.
- **Asignaciones e histórico.** Cada fila de `AsignacionesTablet` / `AsignacionesSIM` es un periodo.
  Si `FechaFin` está vacía, es la asignación vigente; si no, es histórico. Por eso no se pierde nada al
  reasignar.
- La integridad referencial está activada, sin borrado en cascada. No se puede borrar un
  examinador, una tablet o una SIM que tenga histórico, ni un estado en uso. Para retirarlos, cambia
  su situación o su estado, o pon la fecha de desafectación.

## Formularios (desde el menú)

**Datos**
- **Examinadores**: ficha con el histórico de tablets que ha tenido. Se busca por nº de
  examinador, nombre, apellidos (también "nombre apellidos"), DNI, o por el IMEI 1 o nº de serie de
  la tablet que tiene asignada en ese momento. Bajo el buscador hay una casilla por cada situación
  (Activo, Baja…): desmarcándolas se ocultan los examinadores en esas situaciones. Las casillas se
  crean con las situaciones existentes al ejecutar los scripts; si añades una situación nueva,
  ejecuta `aplicar_ultimos_cambios.vbs` (conservando los datos) para que aparezca su casilla.
- **Tablets**: ficha con el histórico de examinadores, el de tarjetas SIM y el de estados (las
  devoluciones y los cambios de estado hechos a mano en la propia ficha). Se busca por nº de
  serie, IMEI 1, IMEI 2, marca o modelo.
- **Tarjetas SIM**: ficha con la compañía, la **foto de la tarjeta** y el histórico de tablets en que
  ha estado. Se busca por IMEI, teléfono o compañía. Para poner o cambiar la foto, haz doble clic
  en su recuadro y pulsa *Agregar…* (la imagen se guarda dentro de la base de datos).
- **Estados de tablet** / **Situaciones de examinador** / **Estados de SIM** / **Compañías**:
  mantenimiento de las tablas auxiliares (altas, cambios y bajas).

**Asignaciones**
- **Asignar / devolver tablet**: eliges examinador, tablet y fecha.
  - Si la tablet la tenía otro examinador, esa asignación se cierra con esa fecha, previa
    confirmación.
  - Si el examinador ya tenía otra tablet, se abre una ventana para elegir **a qué estado pasa la
    tablet que devuelve** (con una observación opcional). Si cancelas, no se asigna nada.
  - *Registrar devolución* cierra la asignación vigente de la tablet elegida y también pregunta a
    qué estado pasa.
  - El estado elegido se guarda en la tablet, en las observaciones de la asignación que se cierra y
    en `HistorialEstadosTablet`.
  - Si la tablet que vas a asignar no está "Operativa", pide confirmación. Las tablets con fecha de
    desafectación no salen en la lista.
- **Asociar / retirar SIM**: igual, entre tablet y tarjeta SIM. Si la SIM no está "Activa", pide
  confirmación.

**Consultas**
- **Situación actual**: cada tablet con su estado, examinador y SIM actuales.
- **Histórico examinador - tablet**: todas las asignaciones, vigentes y pasadas. Se busca por
  nombre, apellidos, nº de examinador o DNI, o por nº de serie o IMEI 1 de la tablet.
- **Histórico tablet - SIM**: se busca por IMEI o teléfono de la SIM, o por nº de serie o IMEI 1 de
  la tablet.

En los buscadores escribes el texto (vale una parte) y pulsas **Intro**. **Ver todos** quita el filtro
(y vuelve a marcar todas las casillas de situación).

## Contenido del proyecto

```
GestionTablets.accdb           Base de datos (tablas, relaciones y valores iniciales de estados)
ConstruirFormularios.vbs       Creación inicial sobre el GestionTablets.accdb de esta carpeta
aplicar_ultimos_cambios.vbs    Aplica la versión actual a la base de datos con la que trabajas
vba/comun.vbs                  Código común de los dos scripts anteriores
vba/modGestionTablets.bas      Lógica de los botones (asignar, devolver, buscar...), queda en la BD
vba/modConstruirInterfaz.bas   Estructura nueva, consultas y formularios; se borra tras usarlo
vba/modImportarDatosAntiguos.bas  Importación desde la base antigua; se borra tras usarlo
generador/GenerarBaseDatos.java  Genera GestionTablets.accdb desde cero (Jackcess)
```

Los `.bas` y los `.vbs` están en codificación Windows-1252 con CRLF, que es lo que espera Access
(`.gitattributes` evita que git los convierta).

### Regenerar la base de datos vacía

**Esto sobrescribe `GestionTablets.accdb` y se pierden los datos.** Solo tiene sentido si se cambia
la estructura antes de empezar a usarla. Hace falta Java 11+ y la librería Jackcess 4 con sus
dependencias:

```bash
java -cp jackcess-4.0.8.jar:commons-lang3-3.10.jar:commons-logging-1.2.jar \
     generador/GenerarBaseDatos.java GestionTablets.accdb
```

Después, vuelve a ejecutar `ConstruirFormularios.vbs`.

## Qué está probado y qué no

- **Probado** (en Linux, sin Access): la estructura del `.accdb` y que la base de datos rechaza DNI
  o IMEI duplicados, estados inexistentes y asignaciones a registros que no existen. También se ha
  probado todo el SQL de consultas, listas y buscadores con UCanAccess, sobre datos de ejemplo,
  incluida la secuencia de reasignación que cierra la asignación anterior.
- **No probado**: la ejecución en Access de los scripts y del código VBA, porque aquí no hay
  Windows ni Access. Si al ejecutarlos aparece un error, anota el mensaje y la línea resaltada para
  corregirlo. Los scripts se pueden volver a ejecutar, y antes de tocar nada dejan una copia de
  seguridad en `copias`.
