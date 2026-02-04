package encuestaSociopruebas;

import org.apache.poi.ss.usermodel.*;
import java.io.*;
import java.time.LocalDate;

public class Procesador {

    private static final int ALTURA_BLOQUE = 7; 
    private static final int ANCHO_BLOQUE  = 6; 

    public static void main(String[] args) {
        int añoViejo = LocalDate.now().getYear() - 7; 
        int añoNuevo = LocalDate.now().getYear() - 1; 

        File fOrigen    = new File("src/fichero/EncuestaSocio.xlsx");
        File fPlantilla = new File("src/fichero/EvolucionEncuestaSocio.xlsx");
        File fResultado = new File("src/fichero/Evolucion_Actualizada.xlsx");

        try (
            Workbook wbOrigen  = WorkbookFactory.create(new FileInputStream(fOrigen));
            Workbook wbDestino = WorkbookFactory.create(new FileInputStream(fPlantilla))
        ) {
            Sheet hojaOrigen  = wbOrigen.getSheetAt(0);
            Sheet hojaDestino = wbDestino.getSheetAt(0);

            moverBloquesYCopiarDatos(hojaDestino, hojaOrigen, añoViejo, añoNuevo);

            try (FileOutputStream fos = new FileOutputStream(fResultado)) {
                wbDestino.write(fos);
            }
            System.out.println("✅ Archivo generado correctamente: " + fResultado.getName());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void moverBloquesYCopiarDatos(
            Sheet hojaDestino,
            Sheet hojaOrigen,
            int añoViejo,
            int añoNuevo
    ) {
        int cursorGlobalOrigen = 0;

        for (Row row : hojaDestino) {
            if (row == null) continue;

            int bloquesEnEstaFila = 0;
            // Variable para recordar la fila de origen dentro de una misma fila de destino
            int filaOrigenFijadaParaEstaFila = -1; 

            for (Cell cell : row) {
                if (esCeldaValor(cell, añoViejo)
                    && hojaDestino.getRow(row.getRowNum() + 1) != null
                    && !esCeldaValor(hojaDestino.getRow(row.getRowNum() + 1).getCell(cell.getColumnIndex()), añoViejo)
                ) {

                    int colInicio  = cell.getColumnIndex();
                    int filaInicio = row.getRowNum(); 

                    // 1️⃣ MOVER BLOQUE (Igual que antes)
                    for (int f = filaInicio; f < filaInicio + ALTURA_BLOQUE; f++) {
                        Row r = hojaDestino.getRow(f);
                        if (r == null) continue;
                        for (int c = colInicio + 1; c < colInicio + ANCHO_BLOQUE; c++) {
                            Cell origen  = r.getCell(c);
                            Cell destino = r.getCell(c - 1);
                            if (destino == null) destino = r.createCell(c - 1);
                            if (origen != null) copiarCelda(origen, destino);
                            else destino.setBlank();
                        }
                        Cell ultima = r.getCell(colInicio + ANCHO_BLOQUE - 1);
                        if (ultima != null) ultima.setBlank();
                    }

                    // 2️⃣ ESCRIBIR AÑO NUEVO (Igual que antes)
                    Row filaAño = hojaDestino.getRow(filaInicio);
                    Cell celdaAñoNueva = filaAño.getCell(colInicio + ANCHO_BLOQUE - 1);
                    if (celdaAñoNueva == null) celdaAñoNueva = filaAño.createCell(colInicio + ANCHO_BLOQUE - 1);
                    celdaAñoNueva.setCellValue(añoNuevo);

                    // 3️⃣ BUSCAR DATOS (Lógica Optimizada)
                    int colOrigenFinal = -1;
                    int filaBaseOrigenParaEsteBloque = -1;

                    // --- DETECTOR PARA LA COLUMNA M (Índice 12) ---
                    // Si estamos en la fila 54 (index 53) y la columna de destino es M (index 12)
                    if (filaInicio == 53 && colInicio == 12) { 
                        filaBaseOrigenParaEsteBloque = 10; // Fila de "Estudios Previos" en Origen
                        colOrigenFinal = 4;               // Columna E (Última)
                        System.out.println("⭐ SALTO ACTIVADO: Pegando en M54 desde E42 (Origen)");
                    } 
                    else {
                        // Lógica normal para el resto de tablas
                        if (bloquesEnEstaFila == 0) {
                            filaOrigenFijadaParaEstaFila = buscarSiguienteFilaDatosOrigen(hojaOrigen, cursorGlobalOrigen);
                            if (filaOrigenFijadaParaEstaFila != -1) {
                                cursorGlobalOrigen = filaOrigenFijadaParaEstaFila + 6; 
                            }
                        }
                        filaBaseOrigenParaEsteBloque = filaOrigenFijadaParaEstaFila;
                        colOrigenFinal = 1 + bloquesEnEstaFila;
                    }

                    // 4️⃣ COPIAR DATOS
                    if (filaBaseOrigenParaEsteBloque != -1) {
                        int filaBaseDestino = filaInicio + 1;

                        for (int i = 0; i < 6; i++) { 
                            Row filaOrigen  = hojaOrigen.getRow(filaBaseOrigenParaEsteBloque + i);
                            Row filaDestino = hojaDestino.getRow(filaBaseDestino + i);

                            if (filaOrigen == null || filaDestino == null) continue;

                            Cell celdaOrigen = filaOrigen.getCell(colOrigenFinal);
                            // Pegamos en la última columna del bloque (colInicio + 5)
                            Cell celdaDestino = filaDestino.getCell(colInicio + ANCHO_BLOQUE - 1);

                            if (celdaDestino == null) celdaDestino = filaDestino.createCell(colInicio + ANCHO_BLOQUE - 1);

                            if (celdaOrigen != null && celdaOrigen.getCellType() == CellType.NUMERIC) {
                                celdaDestino.setCellValue(celdaOrigen.getNumericCellValue());
                            } else {
                                celdaDestino.setBlank();
                            }
                        }
                    }
                    bloquesEnEstaFila++;
                }
            }
        }
    }

    // =====================================================
    // UTILIDADES
    // =====================================================
    private static int buscarSiguienteFilaDatosOrigen(Sheet hoja, int filaDesde) {
        for (int f = filaDesde; f < filaDesde + 100; f++) { 
            Row r = hoja.getRow(f);
            if (r == null) continue;
            // Buscamos una celda numérica en la columna B (índice 1) para saber que hay datos
            Cell c = r.getCell(1); 
            if (c != null && c.getCellType() == CellType.NUMERIC) {
                return f; 
            }
        }
        return -1;
    }

    private static boolean esCeldaValor(Cell c, int val) {
        return c != null && c.getCellType() == CellType.NUMERIC && (int) c.getNumericCellValue() == val;
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