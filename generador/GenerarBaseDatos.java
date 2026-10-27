import com.healthmarketscience.jackcess.Column;
import com.healthmarketscience.jackcess.ColumnBuilder;
import com.healthmarketscience.jackcess.DataType;
import com.healthmarketscience.jackcess.Database;
import com.healthmarketscience.jackcess.DatabaseBuilder;
import com.healthmarketscience.jackcess.IndexBuilder;
import com.healthmarketscience.jackcess.PropertyMap;
import com.healthmarketscience.jackcess.RelationshipBuilder;
import com.healthmarketscience.jackcess.Table;
import com.healthmarketscience.jackcess.TableBuilder;

import java.io.File;

/**
 * Genera GestionTablets.accdb (formato Access 2010+) con las tablas, índices y relaciones.
 * Los formularios y consultas los crea después el propio Access con vba/ConstruirInterfaz.vbs,
 * porque Jackcess no sabe escribir formularios.
 *
 * Uso: java -cp jackcess.jar:commons-lang3.jar:commons-logging.jar GenerarBaseDatos.java salida.accdb
 */
public class GenerarBaseDatos {

    public static void main(String[] args) throws Exception {
        File salida = new File(args.length > 0 ? args[0] : "GestionTablets.accdb");
        if (salida.exists() && !salida.delete()) {
            throw new IllegalStateException("No se puede sobrescribir " + salida);
        }

        try (Database db = new DatabaseBuilder(salida)
                .setFileFormat(Database.FileFormat.V2010)
                .create()) {

            // Tablas auxiliares con valores iniciales; se pueden editar desde el menú de la aplicación.
            Table estadosExaminador = new TableBuilder("EstadosExaminador")
                    .addColumn(autonumerico("IdEstadoExam"))
                    .addColumn(texto("EstadoExaminador", 50, "Situación del examinador", true))
                    .addIndex(clavePrimaria("IdEstadoExam"))
                    .addIndex(unico("EstadoExaminador"))
                    .toTable(db);
            for (String estado : new String[] {"Activo", "Baja temporal", "Baja"}) {
                estadosExaminador.addRow(Column.AUTO_NUMBER, estado);
            }

            Table estadosTablet = new TableBuilder("EstadosTablet")
                    .addColumn(autonumerico("IdEstado"))
                    .addColumn(texto("Estado", 50, "Estado de la tablet", true))
                    .addIndex(clavePrimaria("IdEstado"))
                    .addIndex(unico("Estado"))
                    .toTable(db);
            for (String estado : new String[] {"Operativa", "Averiada", "En reparación", "Desafectada"}) {
                estadosTablet.addRow(Column.AUTO_NUMBER, estado);
            }

            Table estadosSIM = new TableBuilder("EstadosSIM")
                    .addColumn(autonumerico("IdEstadoSIM"))
                    .addColumn(texto("EstadoSIM", 50, "Estado de la SIM", true))
                    .addIndex(clavePrimaria("IdEstadoSIM"))
                    .addIndex(unico("EstadoSIM"))
                    .toTable(db);
            for (String estado : new String[] {"Activa", "Suspendida", "Baja"}) {
                estadosSIM.addRow(Column.AUTO_NUMBER, estado);
            }

            // N/A debe ser la primera (Id 1): es la compañía por defecto de las tarjetas SIM.
            Table companias = new TableBuilder("Companias")
                    .addColumn(autonumerico("IdCompania"))
                    .addColumn(texto("Compania", 50, "Compañía", true))
                    .addIndex(clavePrimaria("IdCompania"))
                    .addIndex(unico("Compania"))
                    .toTable(db);
            for (String compania : new String[] {"N/A", "Movistar", "Orange", "Vodafone"}) {
                companias.addRow(Column.AUTO_NUMBER, compania);
            }

            new TableBuilder("Examinadores")
                    .addColumn(autonumerico("IdExaminador"))
                    .addColumn(texto("NumeroExaminador", 20, "Nº de examinador", false))
                    .addColumn(texto("Nombre", 50, "Nombre", true))
                    .addColumn(texto("Apellidos", 100, "Apellidos", true))
                    .addColumn(texto("DNI", 9, "DNI/NIE", true))
                    .addColumn(entero("IdEstadoExam", "Situación")
                            .putProperty(PropertyMap.DEFAULT_VALUE_PROP, "1"))   // Activo
                    .addColumn(memo("Observaciones"))
                    .addIndex(clavePrimaria("IdExaminador"))
                    .addIndex(unico("DNI"))
                    .addIndex(unicoOpcional("NumeroExaminador"))
                    .addIndex(new IndexBuilder("IdEstadoExam").addColumns("IdEstadoExam"))
                    .addIndex(new IndexBuilder("Apellidos").addColumns("Apellidos", "Nombre"))
                    .toTable(db);

            new TableBuilder("Tablets")
                    .addColumn(autonumerico("IdTablet"))
                    .addColumn(texto("IMEI1", 20, "IMEI 1", false))
                    .addColumn(texto("IMEI2", 20, "IMEI 2", false))
                    .addColumn(texto("NumeroSerie", 50, "Nº de serie", true))
                    .addColumn(texto("Marca", 50, "Marca", false))
                    .addColumn(texto("Modelo", 50, "Modelo", false))
                    .addColumn(fecha("FechaRecepcion", "Fecha de recepción", false)
                            .putProperty(PropertyMap.DEFAULT_VALUE_PROP, "Date()"))
                    .addColumn(fecha("FechaDesafectacion", "Fecha de desafectación", false))
                    .addColumn(entero("IdEstado", "Estado")
                            .putProperty(PropertyMap.DEFAULT_VALUE_PROP, "1"))   // Operativa
                    .addColumn(memo("Observaciones"))
                    .addColumn(memo("Incidencias"))
                    .addIndex(clavePrimaria("IdTablet"))
                    .addIndex(unico("NumeroSerie"))
                    .addIndex(unicoOpcional("IMEI1"))
                    .addIndex(unicoOpcional("IMEI2"))
                    .addIndex(new IndexBuilder("IdEstado").addColumns("IdEstado"))
                    .toTable(db);

            new TableBuilder("TarjetasSIM")
                    .addColumn(autonumerico("IdSIM"))
                    .addColumn(texto("IMEI", 22, "IMEI", true))
                    .addColumn(texto("NumeroTelefono", 20, "Teléfono asociado", false))
                    .addColumn(entero("IdCompania", "Compañía")
                            .putProperty(PropertyMap.DEFAULT_VALUE_PROP, "1"))   // N/A
                    .addColumn(texto("PIN", 8, "PIN", false))
                    .addColumn(texto("PUK", 10, "PUK", false))
                    .addColumn(entero("IdEstadoSIM", "Estado")
                            .putProperty(PropertyMap.DEFAULT_VALUE_PROP, "1"))   // Activa
                    .addColumn(memo("Observaciones"))
                    .addIndex(clavePrimaria("IdSIM"))
                    .addIndex(unico("IMEI"))
                    .addIndex(new IndexBuilder("NumeroTelefono").addColumns("NumeroTelefono"))
                    .addIndex(new IndexBuilder("IdEstadoSIM").addColumns("IdEstadoSIM"))
                    .addIndex(new IndexBuilder("IdCompania").addColumns("IdCompania"))
                    .toTable(db);

            // Cada fila es un periodo de asignación; FechaFin vacía = asignación vigente.
            // Así la misma tabla sirve de asignación actual y de histórico.
            new TableBuilder("AsignacionesTablet")
                    .addColumn(autonumerico("IdAsignacion"))
                    .addColumn(entero("IdExaminador", "Examinador"))
                    .addColumn(entero("IdTablet", "Tablet"))
                    .addColumn(fecha("FechaInicio", "Fecha de entrega", true))
                    .addColumn(fecha("FechaFin", "Fecha de devolución", false))
                    .addColumn(texto("Observaciones", 255, "Observaciones", false))
                    .addIndex(clavePrimaria("IdAsignacion"))
                    .addIndex(new IndexBuilder("IdExaminador").addColumns("IdExaminador"))
                    .addIndex(new IndexBuilder("IdTablet").addColumns("IdTablet"))
                    .toTable(db);

            new TableBuilder("AsignacionesSIM")
                    .addColumn(autonumerico("IdAsignacionSIM"))
                    .addColumn(entero("IdTablet", "Tablet"))
                    .addColumn(entero("IdSIM", "Tarjeta SIM"))
                    .addColumn(fecha("FechaInicio", "Fecha de instalación", true))
                    .addColumn(fecha("FechaFin", "Fecha de retirada", false))
                    .addColumn(texto("Observaciones", 255, "Observaciones", false))
                    .addIndex(clavePrimaria("IdAsignacionSIM"))
                    .addIndex(new IndexBuilder("IdTablet").addColumns("IdTablet"))
                    .addIndex(new IndexBuilder("IdSIM").addColumns("IdSIM"))
                    .toTable(db);

            // Cambios de estado de las tablets: al devolverlas (IdExaminador = quién la devolvió) o
            // al cambiarlo en su ficha. Igual que la crea ConstruirInterfaz si falta.
            new TableBuilder("HistorialEstadosTablet")
                    .addColumn(autonumerico("IdCambio"))
                    .addColumn(entero("IdTablet", "Tablet"))
                    .addColumn(fecha("Fecha", "Fecha", true))
                    .addColumn(new ColumnBuilder("IdEstadoAnterior", DataType.LONG)
                            .putProperty(PropertyMap.CAPTION_PROP, "Estado anterior"))
                    .addColumn(entero("IdEstadoNuevo", "Estado nuevo"))
                    .addColumn(new ColumnBuilder("IdExaminador", DataType.LONG)
                            .putProperty(PropertyMap.CAPTION_PROP, "Devuelta por"))
                    .addColumn(texto("Observaciones", 255, "Observaciones", false))
                    .addIndex(clavePrimaria("IdCambio"))
                    .addIndex(new IndexBuilder("IdTablet").addColumns("IdTablet"))
                    .toTable(db);

            // Integridad referencial sin borrado en cascada: no se puede eliminar un examinador,
            // tablet o SIM con histórico (para retirarlos se cambia su situación / estado), ni un
            // estado que esté en uso.
            relacion(db, "EstadosExaminador", "Examinadores", "IdEstadoExam");
            relacion(db, "EstadosTablet", "Tablets", "IdEstado");
            relacion(db, "Examinadores", "AsignacionesTablet", "IdExaminador");
            relacion(db, "Tablets", "AsignacionesTablet", "IdTablet");
            relacion(db, "Tablets", "AsignacionesSIM", "IdTablet");
            relacion(db, "TarjetasSIM", "AsignacionesSIM", "IdSIM");
            relacion(db, "EstadosSIM", "TarjetasSIM", "IdEstadoSIM");
            relacion(db, "FK_SIM_Compania", "Companias", "TarjetasSIM", "IdCompania", "IdCompania");
            relacion(db, "FK_HET_Tablet", "Tablets", "HistorialEstadosTablet", "IdTablet", "IdTablet");
            relacion(db, "FK_HET_EstadoAnterior", "EstadosTablet", "HistorialEstadosTablet", "IdEstado", "IdEstadoAnterior");
            relacion(db, "FK_HET_EstadoNuevo", "EstadosTablet", "HistorialEstadosTablet", "IdEstado", "IdEstadoNuevo");
            relacion(db, "FK_HET_Examinador", "Examinadores", "HistorialEstadosTablet", "IdExaminador", "IdExaminador");
        }

        System.out.println("Creada " + salida.getAbsolutePath());
    }

    private static ColumnBuilder autonumerico(String nombre) {
        return new ColumnBuilder(nombre, DataType.LONG).setAutoNumber(true);
    }

    private static ColumnBuilder entero(String nombre, String titulo) {
        return new ColumnBuilder(nombre, DataType.LONG)
                .putProperty(PropertyMap.CAPTION_PROP, titulo)
                .putProperty(PropertyMap.REQUIRED_PROP, true);
    }

    private static ColumnBuilder texto(String nombre, int longitud, String titulo, boolean requerido) {
        return new ColumnBuilder(nombre, DataType.TEXT)
                .setLengthInUnits(longitud)
                .putProperty(PropertyMap.CAPTION_PROP, titulo)
                .putProperty(PropertyMap.REQUIRED_PROP, requerido)
                .putProperty(PropertyMap.ALLOW_ZERO_LEN_PROP, !requerido);
    }

    private static ColumnBuilder memo(String nombre) {
        return new ColumnBuilder(nombre, DataType.MEMO)
                .putProperty(PropertyMap.CAPTION_PROP, nombre);
    }

    private static ColumnBuilder fecha(String nombre, String titulo, boolean requerido) {
        return new ColumnBuilder(nombre, DataType.SHORT_DATE_TIME)
                .putProperty(PropertyMap.CAPTION_PROP, titulo)
                .putProperty(PropertyMap.REQUIRED_PROP, requerido)
                .putProperty(PropertyMap.FORMAT_PROP, "Short Date");
    }

    private static IndexBuilder clavePrimaria(String columna) {
        return new IndexBuilder(IndexBuilder.PRIMARY_KEY_NAME).addColumns(columna).setPrimaryKey();
    }

    private static IndexBuilder unico(String columna) {
        return new IndexBuilder(columna).addColumns(columna).setUnique();
    }

    /** Único cuando tiene valor; admite varias filas vacías. */
    private static IndexBuilder unicoOpcional(String columna) {
        return unico(columna).setIgnoreNulls();
    }

    private static void relacion(Database db, String padre, String hija, String columna) throws Exception {
        new RelationshipBuilder(padre, hija)
                .addColumns(columna, columna)
                .setReferentialIntegrity()
                .toRelationship(db);
    }

    /** Relación con nombre propio: hace falta cuando una tabla se relaciona dos veces con la misma. */
    private static void relacion(Database db, String nombre, String padre, String hija,
                                 String columnaPadre, String columnaHija) throws Exception {
        new RelationshipBuilder(padre, hija)
                .addColumns(columnaPadre, columnaHija)
                .setReferentialIntegrity()
                .setName(nombre)
                .toRelationship(db);
    }
}
