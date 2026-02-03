package encuestaSociopruebas;

import org.apache.poi.ss.usermodel.*;
import java.io.*;
import java.time.LocalDate;
import java.util.*;

public class Procesador {

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
        // DINÁMICO: El año nuevo es siempre el anterior al actual
        int añoNuevo = LocalDate.now().getYear() - 1; 

        File fOrigen = new File("src/fichero/EncuestaSocio.xlsx");
        File fPlantilla = new File("src/fichero/EvolucionEncuestaSocio.xlsx");
        File fResultado = new File("src/fichero/Evolucion_Actualizada.xlsx");

        try (FileInputStream fisO = new FileInputStream(fOrigen);
             Workbook wbO = WorkbookFactory.create(fisO);
             FileInputStream fisP = new FileInputStream(fPlantilla);
             Workbook wbD = WorkbookFactory.create(fisP)) {

            Sheet hojaDestino = wbD.getSheetAt(0);
            
            // 1. Buscamos el año más antiguo presente en la plantilla para saber qué "aplastar"
            int añoViejo = buscarAñoMasAntiguo(hojaDestino);
            System.out.println("🔍 Detectado año más antiguo en plantilla: " + añoViejo);
            System.out.println("📅 Preparando inserción de datos para: " + añoNuevo);

            // 2. Cargamos datos nuevos
            Map<String, Map<String, List<Double>>> datosNuevos = cargarDatosOrigen(wbO.getSheetAt(0));

            // 3. Ejecutamos el desplazamiento dinámico
            evolucionarDinamico(hojaDestino, datosNuevos, añoViejo, añoNuevo);

            try (FileOutputStream fos = new FileOutputStream(fResultado)) {
                wbD.write(fos);
                System.out.println("✅ ¡Archivo generado con éxito!");
            }

        } catch (Exception e) {
            System.err.println("❌ Error: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private static int buscarAñoMasAntiguo(Sheet hoja) {
        int minAño = Integer.MAX_VALUE;
        for (Row row : hoja) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.NUMERIC) {
                    double val = cell.getNumericCellValue();
                    if (val > 1900 && val < 2100) { // Rango razonable de años
                        minAño = Math.min(minAño, (int)val);
                    }
                }
            }
        }
        return (minAño == Integer.MAX_VALUE) ? 2019 : minAño;
    }

    private static void evolucionarDinamico(Sheet hoja, Map<String, Map<String, List<Double>>> datos, int añoViejo, int añoNuevo) {
        List<Integer> columnasAncla = new ArrayList<>();
        
        // Localizar todas las columnas donde empieza un bloque (donde esté el año más viejo)
        for (Row row : hoja) {
            for (Cell cell : row) {
                if (esCeldaValor(cell, añoViejo)) {
                    if (!columnasAncla.contains(cell.getColumnIndex())) columnasAncla.add(cell.getColumnIndex());
                }
            }
        }

        String seccionActual = "";

        for (int f = 0; f <= hoja.getLastRowNum(); f++) {
            Row row = hoja.getRow(f);
            if (row == null) continue;

            String label = getCellString(row.getCell(0));
            if (!label.isEmpty() && !TRADUCTOR.containsValue(label)) {
                seccionActual = label.replace(":", "").trim();
            }

            for (int i = 0; i < columnasAncla.size(); i++) {
                int colInicio = columnasAncla.get(i);
                
                // --- MOVIMIENTO DE VENTANA (Desplazar a la izquierda) ---
                // Asumimos que el bloque termina cuando hay una celda vacía o después de 5-6 años
                int anchoBloque = detectarAnchoBloque(hoja, colInicio);
                
                for (int c = colInicio; c < colInicio + anchoBloque - 1; c++) {
                    Cell destino = getOrCreateCell(row, c);
                    Cell origen = row.getCell(c + 1);
                    if (origen != null) copiarCelda(origen, destino);
                }

                // --- INSERTAR DATO NUEVO EN LA ÚLTIMA COLUMNA ---
                int colFinal = colInicio + anchoBloque - 1;
                Cell celdaNueva = getOrCreateCell(row, colFinal);

                if (esCeldaValor(row.getCell(colInicio), añoViejo + 1)) {
                    celdaNueva.setCellValue(añoNuevo);
                } else if (TRADUCTOR.containsValue(label)) {
                    insertarDatoEncuesta(celdaNueva, datos, seccionActual, label, i);
                }
            }
        }
    }

    private static int detectarAnchoBloque(Sheet hoja, int colInicio) {
        // Busca en la fila 2 (donde suelen estar los años) cuántos años seguidos hay
        Row row = hoja.getRow(2); 
        int ancho = 0;
        while (row != null && esCeldaNumerica(row.getCell(colInicio + ancho))) {
            ancho++;
            if (ancho > 10) break; // Seguridad
        }
        return (ancho == 0) ? 5 : ancho; // Por defecto 5 si no detecta
    }

    private static void insertarDatoEncuesta(Cell celda, Map<String, Map<String, List<Double>>> datos, String sec, String ciclo, int index) {
        if (datos.containsKey(sec) && datos.get(sec).containsKey(ciclo)) {
            List<Double> valores = datos.get(sec).get(ciclo);
            if (index < valores.size()) {
                celda.setCellValue(valores.get(index));
            }
        }
    }

    private static Map<String, Map<String, List<Double>>> cargarDatosOrigen(Sheet hoja) {
        Map<String, Map<String, List<Double>>> datos = new HashMap<>();
        String seccion = "";
        for (Row row : hoja) {
            String col0 = getCellString(row.getCell(0));
            if (col0.isEmpty()) continue;
            if (TRADUCTOR.containsKey(col0)) {
                String sigla = TRADUCTOR.get(col0);
                List<Double> valores = new ArrayList<>();
                for (int c = 1; c < 15; c++) {
                    Cell cell = row.getCell(c);
                    if (cell != null && cell.getCellType() == CellType.NUMERIC) valores.add(cell.getNumericCellValue());
                }
                datos.putIfAbsent(seccion, new HashMap<>());
                datos.get(seccion).put(sigla, valores);
            } else {
                seccion = col0.replace(":", "").trim();
            }
        }
        return datos;
    }

    // --- UTILS ---
    private static boolean esCeldaValor(Cell c, int val) {
        return c != null && c.getCellType() == CellType.NUMERIC && (int)c.getNumericCellValue() == val;
    }

    private static boolean esCeldaNumerica(Cell c) {
        return c != null && c.getCellType() == CellType.NUMERIC;
    }

    private static String getCellString(Cell c) {
        return (c == null) ? "" : c.toString().trim();
    }

    private static Cell getOrCreateCell(Row r, int c) {
        Cell cell = r.getCell(c);
        return (cell == null) ? r.createCell(c) : cell;
    }

    private static void copiarCelda(Cell origen, Cell destino) {
        switch (origen.getCellType()) {
            case NUMERIC: destino.setCellValue(origen.getNumericCellValue()); break;
            case STRING: destino.setCellValue(origen.getStringCellValue()); break;
            case FORMULA: destino.setCellFormula(origen.getCellFormula()); break;
            default: destino.setBlank();
        }
    }
}