package encuestaSociopruebas;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Procesador {

    public static void main(String[] args) {

        // =========================
        // 1. RUTAS DE LOS ARCHIVOS
        // =========================
        String rutaEncuesta = "src/fichero/EncuestaSocio.xlsx";
        String rutaTablas   = "src/fichero/EvolucionEncuestaSocio.xlsx";
        String rutaSalida   = "src/fichero/Evolucion_2025.xlsx";

        // =========================
        // 2. TRADUCTOR DE NOMBRES
        // =========================
        Map<String, String> traductor = new HashMap<>();
        traductor.put("F.B. Básica", "FPB");
        traductor.put("G.M. Administración", "GMA");
        traductor.put("G.M.Informática", "SMR");
        traductor.put("G.S. Administración y finanzas", "GSA");
        traductor.put("G.S. Marketing y publicidad", "MPU");
        traductor.put("G.S. Desarrollo de apliciones multiplataforma", "DAM");

        System.out.println("=== Iniciando actualización anual de la encuesta ===");

        try (
            FileInputStream fisOrigen = new FileInputStream(new File(rutaEncuesta));
            Workbook wbOrigen = new XSSFWorkbook(fisOrigen);

            FileInputStream fisDestino = new FileInputStream(new File(rutaTablas));
            Workbook wbDestino = new XSSFWorkbook(fisDestino);
        ) {

            Sheet hojaOrigen  = wbOrigen.getSheetAt(0);
            Sheet hojaDestino = wbDestino.getSheetAt(0);

            // =========================
            // 3. EXTRAER DATOS NUEVOS
            // =========================
            Map<String, Double> datosNuevos = new HashMap<>();

            for (int i = 1; i <= 6; i++) {
                Row fila = hojaOrigen.getRow(i);
                if (fila == null) continue;

                String nombreLargo = fila.getCell(0).getStringCellValue();
                double valor       = fila.getCell(1).getNumericCellValue();

                if (traductor.containsKey(nombreLargo)) {
                    String sigla = traductor.get(nombreLargo);
                    datosNuevos.put(sigla, valor);
                }
            }

            System.out.println("Datos nuevos: " + datosNuevos);

            // =========================
            // 4. ACTUALIZAR TODAS LAS TABLAS
            // =========================
            procesarTablaMediaEdad(hojaDestino, datosNuevos);

            // =========================
            // 5. GUARDAR RESULTADO
            // =========================
            try (FileOutputStream fos = new FileOutputStream(rutaSalida)) {
                wbDestino.write(fos);
                System.out.println("Archivo generado correctamente: " + rutaSalida);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Procesa TODAS las tablas de la hoja:
     * - Detecta cada cabecera de años
     * - Desplaza los valores a la izquierda
     * - Inserta el valor nuevo del año actual
     * - Limpia columnas sobrantes para evitar duplicados
     */
    private static void procesarTablaMediaEdad(Sheet hoja, Map<String, Double> datosNuevos) {

        int colInicio = 1; // Columna B
        int colFin    = 6; // Columna G

        for (Row filaAnios : hoja) {

            Cell posibleAnio = filaAnios.getCell(colInicio);

            // Detectamos fila de encabezados (años)
            if (posibleAnio != null &&
                posibleAnio.getCellType() == CellType.NUMERIC &&
                posibleAnio.getNumericCellValue() > 2000) {

                // 1️⃣ Actualizar encabezados
                actualizarEncabezadosAnios(filaAnios, colInicio, colFin);

                int filaDatosActual = filaAnios.getRowNum() + 1;

                // 2️⃣ Recorrer filas de la tabla
                while (true) {
                    Row filaDatos = hoja.getRow(filaDatosActual);
                    if (filaDatos == null) break;

                    Cell celdaSigla = filaDatos.getCell(0);
                    if (celdaSigla == null || celdaSigla.toString().isBlank()) break;

                    String sigla = celdaSigla.toString().trim();

                    if (datosNuevos.containsKey(sigla)) {

                        // Mover datos a la izquierda
                        for (int c = colInicio; c < colFin; c++) {
                            Cell origen  = filaDatos.getCell(c + 1);
                            Cell destino = filaDatos.getCell(c);

                            if (destino == null) destino = filaDatos.createCell(c);

                            if (origen != null && origen.getCellType() == CellType.NUMERIC) {
                                destino.setCellValue(origen.getNumericCellValue());
                            } else {
                                destino.setBlank();
                            }
                        }

                        // Insertar nuevo dato
                        Cell celdaNueva = filaDatos.getCell(colFin);
                        if (celdaNueva == null) celdaNueva = filaDatos.createCell(colFin);

                        celdaNueva.setCellValue(datosNuevos.get(sigla));

                        // Limpiar columnas sobrantes (el problema que tenías)
                        for (int c = colFin + 1; c <= colFin + 5; c++) {
                            Cell sobrante = filaDatos.getCell(c);
                            if (sobrante != null) filaDatos.removeCell(sobrante);
                        }
                    }

                    filaDatosActual++;
                }
            }
        }
    }

    /**
     * Actualiza los años del encabezado
     */
    private static void actualizarEncabezadosAnios(Row row, int colInicio, int colFin) {

        int anio = 2020; // ajusta si quieres que empiece en otro año

        for (int c = colInicio; c <= colFin; c++) {
            Cell celda = row.getCell(c);
            if (celda == null) celda = row.createCell(c);
            celda.setCellValue(anio++);
        }
    }
}