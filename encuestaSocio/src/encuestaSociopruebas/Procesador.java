package encuestaSociopruebas;

import org.apache.poi.ss.usermodel.*;
import java.io.*;
import java.time.LocalDate;
import java.util.*;

public class Procesador {

    // Bloque fijo según tu Excel
    private static final int ALTURA_BLOQUE = 7; // año + 6 filas
    private static final int ANCHO_BLOQUE  = 6; // 2019–2024

    private static final Map<String, String> TRADUCTOR = new HashMap<>();
    static {
        TRADUCTOR.put("F.B. Básica", "FPB");
        TRADUCTOR.put("G.M. Administración", "GMA");
        TRADUCTOR.put("G.M.Informática", "SMR");
        TRADUCTOR.put("G.S. Administración y finanzas", "GSA");
        TRADUCTOR.put("G.S. Marketing y publicidad", "MPU");
        TRADUCTOR.put("G.S. Desarrollo de apliciones multiplataforma", "DAM");
    }

    public static void main(String[] args) {

        int añoViejo = LocalDate.now().getYear() - 7;
        int añoNuevo = LocalDate.now().getYear() - 1;

        File fOrigen    = new File("src/fichero/EncuestaSocio.xlsx");
        File fPlantilla = new File("src/fichero/EvolucionEncuestaSocio.xlsx");
        File fResultado = new File("src/fichero/Evolucion_Actualizada.xlsx");

        try (
                Workbook wbOrigen = WorkbookFactory.create(new FileInputStream(fOrigen));
                Workbook wbDestino = WorkbookFactory.create(new FileInputStream(fPlantilla))
        ) {

            Sheet hojaDestino = wbDestino.getSheetAt(0);
            Sheet hojaOrigen  = wbOrigen.getSheetAt(0);

            // 1️⃣ cargar datos del Excel origen
            Map<String, Map<String, List<Double>>> datosOrigen =
                    cargarDatosOrigen(hojaOrigen);

            // 2️⃣ mover bloque + escribir año nuevo (YA FUNCIONABA)
            moverBloqueIzquierdaDesdeAño(
                    hojaDestino, añoViejo, añoNuevo, datosOrigen
            );

            try (FileOutputStream fos = new FileOutputStream(fResultado)) {
                wbDestino.write(fos);
            }

            System.out.println("✅ Archivo nuevo generado correctamente");
            System.out.println(fResultado.getAbsolutePath());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // =====================================================
    // MOVER BLOQUE + AÑO NUEVO + DATOS
    // =====================================================

    private static void moverBloqueIzquierdaDesdeAño(
            Sheet hoja,
            int añoViejo,
            int añoNuevo,
            Map<String, Map<String, List<Double>>> datosOrigen
    ) {

        String seccionActual = "";

        for (Row row : hoja) {
            if (row == null) continue;

            String col0 = normalizar(getCellString(row.getCell(0)));
            if (!col0.isEmpty() && !TRADUCTOR.containsKey(col0)) {
                seccionActual = col0.replace(":", "").trim();
            }

            for (Cell cell : row) {

                if (esCeldaValor(cell, añoViejo)) {

                    int colInicio  = cell.getColumnIndex();
                    int filaInicio = row.getRowNum();

                    // --- mover bloque (NO TOCADO) ---
                    for (int f = filaInicio; f < filaInicio + ALTURA_BLOQUE; f++) {

                        Row r = hoja.getRow(f);
                        if (r == null) continue;

                        for (int c = colInicio + 1; c < colInicio + ANCHO_BLOQUE; c++) {

                            Cell origen  = r.getCell(c);
                            Cell destino = r.getCell(c - 1);

                            if (destino == null) {
                                destino = r.createCell(c - 1);
                            }

                            if (origen != null) {
                                copiarCelda(origen, destino);
                            } else {
                                destino.setBlank();
                            }
                        }

                        Cell ultima = r.getCell(colInicio + ANCHO_BLOQUE - 1);
                        if (ultima != null) ultima.setBlank();
                    }

                    // --- escribir año nuevo ---
                    Row filaAño = hoja.getRow(filaInicio);
                    Cell celdaAñoNueva = filaAño.getCell(colInicio + ANCHO_BLOQUE - 1);
                    if (celdaAñoNueva == null) {
                        celdaAñoNueva = filaAño.createCell(colInicio + ANCHO_BLOQUE - 1);
                    }
                    celdaAñoNueva.setCellValue(añoNuevo);

                    // --- insertar datos debajo (CORREGIDO) ---
                    String ciclo = normalizar(getCellString(filaAño.getCell(0)));

                    insertarDatosNuevoAño(
                            hoja,
                            filaInicio,
                            colInicio + ANCHO_BLOQUE - 1,
                            seccionActual,
                            ciclo,
                            datosOrigen
                    );
                }
            }
        }
    }

    // =====================================================
    // INSERTAR DATOS DEL NUEVO AÑO (VERSIÓN CORRECTA)
    // =====================================================

    private static void insertarDatosNuevoAño(
            Sheet hoja,
            int filaInicio,
            int colNueva,
            String seccion,
            String cicloIgnorado, // ya no se usa
            Map<String, Map<String, List<Double>>> datos
    ) {

        if (!datos.containsKey(seccion)) return;

        // recorrer las 6 filas de datos
        for (int i = 0; i < 6; i++) {

            Row filaDato = hoja.getRow(filaInicio + 1 + i);
            if (filaDato == null) continue;

            // 🔑 el ciclo se obtiene AQUÍ
            String sigla = normalizar(getCellString(filaDato.getCell(0)));

            if (!datos.get(seccion).containsKey(sigla)) continue;

            List<Double> valores = datos.get(seccion).get(sigla);

            // 👉 SIEMPRE el último valor
            double valorNuevo = valores.get(valores.size() - 1);

            Cell destino = filaDato.getCell(colNueva);
            if (destino == null) destino = filaDato.createCell(colNueva);

            destino.setCellValue(valorNuevo);
        }
    }


    // =====================================================
    // CARGAR DATOS ORIGEN
    // =====================================================

    private static Map<String, Map<String, List<Double>>> cargarDatosOrigen(Sheet hoja) {

        Map<String, Map<String, List<Double>>> datos = new HashMap<>();
        String seccion = "";

        for (Row row : hoja) {
            String col0 = normalizar(getCellString(row.getCell(0)));
            if (col0.isEmpty()) continue;

            if (TRADUCTOR.containsKey(col0)) {

                String sigla = TRADUCTOR.get(col0);
                List<Double> valores = new ArrayList<>();

                for (int c = 1; c <= 6; c++) {
                    Cell cell = row.getCell(c);
                    if (cell != null && cell.getCellType() == CellType.NUMERIC) {
                        valores.add(cell.getNumericCellValue());
                    }
                }

                datos.putIfAbsent(seccion, new HashMap<>());
                datos.get(seccion).put(sigla, valores);

            } else {
                seccion = col0.replace(":", "").trim();
            }
        }
        return datos;
    }

    // =====================================================
    // NORMALIZACIÓN
    // =====================================================

    private static String normalizar(String texto) {
        if (texto == null) return "";
        return texto
                .replace("Ã¡", "á")
                .replace("Ã©", "é")
                .replace("Ã­", "í")
                .replace("Ã³", "ó")
                .replace("Ãº", "ú")
                .replace("Ã±", "ñ")
                .replace("Ã", "Á")
                .trim();
    }

    // =====================================================
    // UTILIDADES
    // =====================================================

    private static boolean esCeldaValor(Cell c, int val) {
        return c != null && c.getCellType() == CellType.NUMERIC
                && (int) c.getNumericCellValue() == val;
    }

    private static String getCellString(Cell c) {
        return (c == null) ? "" : c.toString().trim();
    }

    private static void copiarCelda(Cell origen, Cell destino) {
        switch (origen.getCellType()) {
            case NUMERIC -> destino.setCellValue(origen.getNumericCellValue());
            case STRING  -> destino.setCellValue(origen.getStringCellValue());
            case FORMULA -> destino.setCellFormula(origen.getCellFormula());
            case BOOLEAN -> destino.setCellValue(origen.getBooleanCellValue());
            default      -> destino.setBlank();
        }
    }
}
