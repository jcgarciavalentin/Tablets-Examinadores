# Gestión de tablets (Microsoft Access)

Base de datos de Access (`GestionTablets.accdb`) para llevar el control de examinadores, tablets y
tarjetas SIM: qué SIM tiene cada tablet, qué tablet tiene cada examinador, y el histórico de todo ello.

## Puesta en marcha (una sola vez)

1. Copia la carpeta `BaseDatosTablets` completa a un equipo Windows con Microsoft Access.
2. Haz doble clic en **`ConstruirFormularios.vbs`** con la base de datos cerrada. Se abre Access,
   crea las consultas, los formularios y el menú dentro de `GestionTablets.accdb` y se cierra. Tarda
   unos segundos.
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

## Tablas

| Tabla | Campos |
|---|---|
| `Examinadores` | IdExaminador, NumeroExaminador, Nombre, Apellidos, DNI, IdEstadoExam (situación), Observaciones |
| `EstadosExaminador` | IdEstadoExam, EstadoExaminador: *Activo, Baja temporal, Baja* |
| `Tablets` | IdTablet, IMEI1, IMEI2, NumeroSerie, Marca, Modelo, FechaRecepcion, FechaDesafectacion, IdEstado, Observaciones, Incidencias |
| `EstadosTablet` | IdEstado, Estado: *Operativa, Averiada, En reparación, Desafectada* |
| `TarjetasSIM` | IdSIM, IMEI, PIN, PUK |
| `AsignacionesTablet` | IdAsignacion, IdExaminador, IdTablet, FechaInicio, FechaFin, Observaciones |
| `AsignacionesSIM` | IdAsignacionSIM, IdTablet, IdSIM, FechaInicio, FechaFin, Observaciones |

- Los valores iniciales de las tablas de estados son orientativos. Se cambian desde el menú
  (*Estados de tablet* / *Situaciones de examinador*).
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
  examinador, nombre o apellidos (también "nombre apellidos").
- **Tablets**: ficha con el histórico de examinadores y el de tarjetas SIM. Se busca por nº de
  serie, IMEI 1, IMEI 2, marca o modelo.
- **Tarjetas SIM**: ficha con el histórico de tablets en que ha estado. Se busca por IMEI.
- **Estados de tablet** / **Situaciones de examinador**: mantenimiento de las tablas auxiliares.

**Asignaciones**
- **Asignar / devolver tablet**: eliges examinador, tablet y fecha. Si la tablet ya tenía
  examinador, o el examinador ya tenía tablet, la asignación anterior se cierra con esa fecha, previa
  confirmación. *Registrar devolución* cierra la asignación vigente de la tablet elegida. Si la
  tablet no está "Operativa", pide confirmación. Las tablets con fecha de desafectación no salen en
  la lista.
- **Asociar / retirar SIM**: igual, entre tablet y tarjeta SIM.

**Consultas**
- **Situación actual**: cada tablet con su estado, examinador y SIM actuales.
- **Histórico examinador - tablet**: todas las asignaciones, vigentes y pasadas. Se busca por
  nombre, apellidos o nº de examinador, o por nº de serie o IMEI 1 de la tablet.
- **Histórico tablet - SIM**: se busca por IMEI de la SIM, o por nº de serie o IMEI 1 de la tablet.

En los buscadores escribes el texto (vale una parte) y pulsas **Intro**. **Ver todos** quita el filtro.

## Contenido de la carpeta

```
GestionTablets.accdb           Base de datos (tablas, relaciones y valores iniciales de estados)
ConstruirFormularios.vbs       Crea consultas, formularios y menú dentro de la base de datos
vba/modGestionTablets.bas      Lógica de los botones (asignar, devolver, buscar...), queda en la BD
vba/modConstruirInterfaz.bas   Construye consultas y formularios; se borra tras usarlo
generador/GenerarBaseDatos.java  Genera GestionTablets.accdb desde cero (Jackcess)
```

Los `.bas` y el `.vbs` están en codificación Windows-1252 con CRLF, que es lo que espera Access
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
- **No probado**: la ejecución en Access del script de formularios y del código VBA, porque aquí no
  hay Windows ni Access. Si al ejecutar `ConstruirFormularios.vbs` aparece un error de VBA, anota el
  mensaje y la línea resaltada para corregirlo. El script se puede volver a ejecutar sin perder datos.
