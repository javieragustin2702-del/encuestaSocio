package encuestaSociopruebas;

import org.apache.poi.ss.usermodel.*;
import java.io.*;
import java.time.LocalDate;

public class Procesador {

    private static final int ALTURA_BLOQUE = 7; // año + 6 filas de datos
    private static final int ANCHO_BLOQUE  = 6; // 2019–2024

    public static void main(String[] args) {

        int añoViejo = LocalDate.now().getYear() - 7; // 2019
        int añoNuevo = LocalDate.now().getYear() - 1; // 2025

        File fOrigen    = new File("src/fichero/EncuestaSocio.xlsx");
        File fPlantilla = new File("src/fichero/EvolucionEncuestaSocio.xlsx");
        File fResultado = new File("src/fichero/Evolucion_Actualizada.xlsx");

        try (
            Workbook wbOrigen  = WorkbookFactory.create(new FileInputStream(fOrigen));
            Workbook wbDestino = WorkbookFactory.create(new FileInputStream(fPlantilla))
        ) {

            Sheet hojaOrigen  = wbOrigen.getSheetAt(0);
            Sheet hojaDestino = wbDestino.getSheetAt(0);

            moverBloquesYCopiarDatos(
                hojaDestino,
                hojaOrigen,
                añoViejo,
                añoNuevo
            );

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

        // Cursor global que recuerda por dónde vamos en el archivo vertical (Origen)
        int cursorGlobalOrigen = 0;

        for (Row row : hojaDestino) {
            if (row == null) continue;

            // Variables para controlar tablas que están una al lado de otra
            int bloquesEnEstaFila = 0;
            int filaBaseOrigenParaEstaFila = -1;

            for (Cell cell : row) {

                // Detectamos la cabecera del año viejo (ej. 2019)
                if (
                    esCeldaValor(cell, añoViejo)
                    && hojaDestino.getRow(row.getRowNum() + 1) != null
                    && !esCeldaValor(
                        hojaDestino
                            .getRow(row.getRowNum() + 1)
                            .getCell(cell.getColumnIndex()),
                        añoViejo
                    )
                ) {

                    int colInicio  = cell.getColumnIndex();
                    int filaInicio = row.getRowNum();

                    // ===============================
                    // 1️⃣ MOVER BLOQUE (DESPLAZAR AÑOS)
                    // ===============================
                    for (int f = filaInicio; f < filaInicio + ALTURA_BLOQUE; f++) {
                        Row r = hojaDestino.getRow(f);
                        if (r == null) continue;

                        // Copiamos de derecha a izquierda: Columna 2->1, 3->2...
                        for (int c = colInicio + 1; c < colInicio + ANCHO_BLOQUE; c++) {
                            Cell origen  = r.getCell(c);
                            Cell destino = r.getCell(c - 1);

                            if (destino == null) destino = r.createCell(c - 1);

                            if (origen != null) copiarCelda(origen, destino);
                            else destino.setBlank();
                        }

                        // Limpiar la última columna (donde irá el 2025)
                        Cell ultima = r.getCell(colInicio + ANCHO_BLOQUE - 1);
                        if (ultima != null) ultima.setBlank();
                    }

                    // ===============================
                    // 2️⃣ ESCRIBIR AÑO NUEVO (2025)
                    // ===============================
                    Row filaAño = hojaDestino.getRow(filaInicio);
                    Cell celdaAñoNueva =
                            filaAño.getCell(colInicio + ANCHO_BLOQUE - 1);

                    if (celdaAñoNueva == null)
                        celdaAñoNueva =
                                filaAño.createCell(colInicio + ANCHO_BLOQUE - 1);

                    celdaAñoNueva.setCellValue(añoNuevo);


                    // ===============================
                    // 3️⃣ BUSCAR DATOS EN ORIGEN (INTELIGENTE)
                    // ===============================
                    
                    // Si es el primer bloque que encontramos en esta fila, buscamos nueva posición vertical
                    if (bloquesEnEstaFila == 0) {
                        filaBaseOrigenParaEstaFila = buscarSiguienteFilaDatosOrigen(hojaOrigen, cursorGlobalOrigen);
                        // Avanzamos el cursor global para que la siguiente FILA del excel destino no repita estos datos
                        if (filaBaseOrigenParaEstaFila != -1) {
                            cursorGlobalOrigen = filaBaseOrigenParaEstaFila + 6; 
                        }
                    }

                    // ===============================
                    // 4️⃣ COPIAR DATOS
                    // ===============================
                    if (filaBaseOrigenParaEstaFila != -1) {
                        
                        // Lógica clave:
                        // Bloque 1 (ej ESO) -> lee Columna B (index 1)
                        // Bloque 2 (ej FPB) -> lee Columna C (index 2)
                        // Bloque 3 (ej GM)  -> lee Columna D (index 3)...
                        int colOrigen = 1 + bloquesEnEstaFila; 
                        
                        int filaBaseDestino = filaInicio + 1;

                        for (int i = 0; i < 6; i++) { // 6 filas de datos
                            Row filaOrigen  = hojaOrigen.getRow(filaBaseOrigenParaEstaFila + i);
                            Row filaDestino = hojaDestino.getRow(filaBaseDestino + i);

                            if (filaOrigen == null || filaDestino == null) continue;

                            Cell celdaOrigen = filaOrigen.getCell(colOrigen);
                            Cell celdaDestino =
                                filaDestino.getCell(colInicio + ANCHO_BLOQUE - 1);

                            if (celdaDestino == null)
                                celdaDestino =
                                    filaDestino.createCell(colInicio + ANCHO_BLOQUE - 1);

                            if (celdaOrigen != null && celdaOrigen.getCellType() == CellType.NUMERIC) {
                                celdaDestino.setCellValue(celdaOrigen.getNumericCellValue());
                            }
                        }
                    }

                    // Incrementamos contador de bloques en esta fila horizontal
                    bloquesEnEstaFila++;
                    
                    // 🛑 IMPORTANTE: Quitamos el 'break' para que siga buscando
                    // más tablas a la derecha en la misma fila.
                }
            }
        }
    }

    // =====================================================
    // BUSCADOR DE DATOS EN EL ORIGEN (Vertical)
    // =====================================================
    private static int buscarSiguienteFilaDatosOrigen(Sheet hoja, int filaDesde) {
        // Buscamos hacia abajo una celda numérica en la columna B
        for (int f = filaDesde; f < filaDesde + 100; f++) { // Límite seguridad aumentado
            Row r = hoja.getRow(f);
            if (r == null) continue;

            Cell c = r.getCell(1); // Columna B
            if (c != null && c.getCellType() == CellType.NUMERIC) {
                return f; 
            }
        }
        return -1;
    }

    // =====================================================
    // UTILIDADES
    // =====================================================
    private static boolean esCeldaValor(Cell c, int val) {
        return c != null
            && c.getCellType() == CellType.NUMERIC
            && (int) c.getNumericCellValue() == val;
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