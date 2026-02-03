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
        String rutaEncuesta    = "src/fichero/EncuestaSocio.xlsx";           // Excel con datos nuevos (2025)
        String rutaTablas = "src/fichero/EvolucionEncuestaSocio.xlsx";  // Excel histórico
        String rutaSalida    = "src/fichero/Evolucion_2025.xlsx";          // Nuevo archivo generado

        // =========================
        // 2. TRADUCTOR DE NOMBRES
        // =========================
        // Relaciona el nombre largo del ciclo con su sigla corta
        Map<String, String> traductor = new HashMap<>();
        traductor.put("F.B. Básica", "FPB");
        traductor.put("G.M. Administración", "GMA");
        traductor.put("G.M.Informática", "SMR");
        traductor.put("G.S. Administración y finanzas", "GSA");
        traductor.put("G.S. Marketing y publicidad", "MPU");
        traductor.put("G.S. Desarrollo de apliciones multiplataforma", "DAM");

        System.out.println("=== Iniciando actualización anual de la encuesta ===");

        // =========================
        // 3. APERTURA DE EXCEL
        // =========================
        try (
            FileInputStream fisOrigen = new FileInputStream(new File(rutaEncuesta));
            Workbook wbOrigen = new XSSFWorkbook(fisOrigen);

            FileInputStream fisDestino = new FileInputStream(new File(rutaTablas));
            Workbook wbDestino = new XSSFWorkbook(fisDestino);
        ) {

            Sheet hojaOrigen  = wbOrigen.getSheetAt(0);
            Sheet hojaDestino = wbDestino.getSheetAt(0);

            // =========================
            // 4. EXTRAER DATOS NUEVOS
            // =========================
            // Guardamos los datos en un Map: "DAM" -> 21.5
            Map<String, Double> datosNuevos = new HashMap<>();

            // Recorremos las filas donde está la media de edad (filas 1 a 6)
            for (int i = 1; i <= 6; i++) {
                Row fila = hojaOrigen.getRow(i);
                if (fila == null) continue;

                String nombreLargo = fila.getCell(0).getStringCellValue();
                double valor       = fila.getCell(1).getNumericCellValue();

                // Convertimos nombre largo a sigla
                if (traductor.containsKey(nombreLargo)) {
                    String sigla = traductor.get(nombreLargo);
                    datosNuevos.put(sigla, valor);
                }
            }

            System.out.println("Datos nuevos extraídos: " + datosNuevos);

            // =========================
            // 5. ACTUALIZAR HISTÓRICO
            // =========================
            procesarTablaMediaEdad(hojaDestino, datosNuevos);

            // =========================
            // 6. GUARDAR ARCHIVO NUEVO
            // =========================
            try (FileOutputStream fos = new FileOutputStream(rutaSalida)) {
                wbDestino.write(fos);
                System.out.println("Archivo generado correctamente: " + rutaSalida);
            }

        } catch (Exception e) {
            System.err.println("Error durante el proceso: " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * Actualiza la tabla de "Media de edad":
     * - Mueve los valores antiguos a la izquierda
     * - Inserta los datos nuevos del año actual
     */
    private static void procesarTablaMediaEdad(Sheet hoja, Map<String, Double> datosNuevos) {

        int colInicio = 1; // Columna B (primer año visible)
        int colFin    = 6; // Última columna (año más reciente)

        // =========================
        // 1. BUSCAR FILA DE AÑOS
        // =========================
        int filaAnios = -1;

        for (Row row : hoja) {
            Cell celda = row.getCell(colInicio);
            if (celda != null 
                && celda.getCellType() == CellType.NUMERIC 
                && celda.getNumericCellValue() > 2000) {

                filaAnios = row.getRowNum();
                actualizarEncabezadosAnios(row, colInicio, colFin);
                break;
            }
        }

        if (filaAnios == -1) {
            System.out.println("⚠️ No se encontró la fila de encabezados de años.");
            return;
        }

        // =========================
        // 2. RECORRER CICLOS
        // =========================
        for (Row row : hoja) {

            Cell celdaSigla = row.getCell(0);
            if (celdaSigla == null) continue;

            String sigla = celdaSigla.toString();

            if (datosNuevos.containsKey(sigla)) {

                // ---- A) DESPLAZAR VALORES A LA IZQUIERDA ----
                for (int c = colInicio; c < colFin; c++) {
                    Cell origen  = row.getCell(c + 1);
                    Cell destino = row.getCell(c);

                    if (origen != null) {
                        if (destino == null) destino = row.createCell(c);

                        if (origen.getCellType() == CellType.NUMERIC) {
                            destino.setCellValue(origen.getNumericCellValue());
                        } else {
                            destino.setCellValue(origen.toString());
                        }
                    }
                }

                // ---- B) INSERTAR EL DATO NUEVO EN LA ÚLTIMA COLUMNA ----
                Cell celdaNueva = row.getCell(colFin);
                if (celdaNueva == null) celdaNueva = row.createCell(colFin);

                double valorNuevo = datosNuevos.get(sigla);
                celdaNueva.setCellValue(valorNuevo);

                System.out.println("→ " + sigla + " actualizado con el valor " + valorNuevo);
            }
        }
    }

    /**
     * Actualiza los años del encabezado (por ejemplo: 2020 a 2025)
     */
    private static void actualizarEncabezadosAnios(Row row, int colInicio, int colFin) {

        int anio = 2020; // Nuevo primer año visible

        for (int c = colInicio; c <= colFin; c++) {
            Cell celda = row.getCell(c);
            if (celda == null) celda = row.createCell(c);

            celda.setCellValue(anio);
            anio++;
        }

        System.out.println("Encabezados de años actualizados.");
    }
}