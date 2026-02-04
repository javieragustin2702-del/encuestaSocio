package encuestaSociopruebas;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDate;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Prueba {

    // ================= BORRAR COLUMNA =================
    public static void borrarColumna(Sheet hoja, Cell cell) {

        int columna = cell.getColumnIndex();
        int filaInicial = cell.getRowIndex();

        for (int i = filaInicial; i <= filaInicial + 6; i++) {

            Row fila = hoja.getRow(i);
            if (fila == null) continue;

            Cell celdaABorrar = fila.getCell(columna);
            if (celdaABorrar != null) {
                celdaABorrar.setBlank();
            }
        }
    }

    // ================= MOVER BLOQUE =================
    public static void moverBloqueIzquierda(Sheet hoja, Cell cell) {

        int colBase = cell.getColumnIndex();
        int filaBase = cell.getRowIndex();

        for (int fila = filaBase; fila <= filaBase + 6; fila++) {

            Row row = hoja.getRow(fila);
            if (row == null) continue;

            for (int col = colBase + 1; col <= colBase + 5; col++) {

                Cell origen = row.getCell(col);
                if (origen == null) continue;

                Cell destino = row.getCell(col - 1);
                if (destino == null)
                    destino = row.createCell(col - 1);

                copiarValorCelda(origen, destino);
                origen.setBlank();
            }
        }
    }

    // ================= COPIAR CELDA =================
    private static void copiarValorCelda(Cell origen, Cell destino) {

        switch (origen.getCellType()) {

        case STRING:
            destino.setCellValue(origen.getStringCellValue());
            break;

        case NUMERIC:
            if (DateUtil.isCellDateFormatted(origen)) {
                destino.setCellValue(origen.getDateCellValue());
            } else {
                destino.setCellValue(origen.getNumericCellValue());
            }
            break;

        case BOOLEAN:
            destino.setCellValue(origen.getBooleanCellValue());
            break;

        case FORMULA:
            destino.setCellFormula(origen.getCellFormula());
            break;

        default:
            destino.setBlank();
        }
    }

    // ================= RELLENAR 2025 + DATOS =================
    private static void rellenar2025YDatos(
            Sheet destino,
            Cell celdaBase,
            Sheet origen,
            int filaOrigen,
            int colOrigen) {

        int filaBase = celdaBase.getRowIndex();
        int col2025 = celdaBase.getColumnIndex() + 5;

        Row fila = destino.getRow(filaBase);
        if (fila == null)
            fila = destino.createRow(filaBase);

        Cell celda2025 = fila.getCell(col2025);
        if (celda2025 == null)
            celda2025 = fila.createCell(col2025);

        celda2025.setCellValue(2025);

        for (int i = 1; i <= 6; i++) {

            Row filaDestino = destino.getRow(filaBase + i);
            if (filaDestino == null)
                filaDestino = destino.createRow(filaBase + i);

            Cell celdaDestino = filaDestino.getCell(col2025);
            if (celdaDestino == null)
                celdaDestino = filaDestino.createCell(col2025);

            Row filaOrigenExcel = origen.getRow(filaOrigen + i - 1);
            if (filaOrigenExcel == null) continue;

            Cell celdaOrigen = filaOrigenExcel.getCell(colOrigen);
            if (celdaOrigen != null) {
                copiarValorCelda(celdaOrigen, celdaDestino);
            }
        }
    }

    // ================= MAIN =================
    public static void main(String[] args) {

        LocalDate fecha = LocalDate.now().minusYears(7);
        int año = fecha.getYear();
        String ca = String.valueOf(año);

        File archivoDestinoOriginal =
                new File("src/fichero/EvolucionEncuestaSocio.xlsx");
        File archivoOrigen =
                new File("src/fichero/EncuestaSocio.xlsx");

        File archivoNuevo =
                new File("src/fichero/RESULTADO.xlsx");

        try (
            Workbook wbDestino =
                    WorkbookFactory.create(new FileInputStream(archivoDestinoOriginal));
            Workbook wbOrigen =
                    WorkbookFactory.create(new FileInputStream(archivoOrigen))
        ) {

            Sheet sheetDestino = wbDestino.getSheetAt(0);
            Sheet sheetOrigen = wbOrigen.getSheetAt(0);

            int filaOrigenBase = 1;
            int colOrigenBase = 1;

            Integer filaDestinoAnterior = null;

            int contadorFila53 = 0;
            int colBaseFila53 = 1;

            for (int i = 0; i <= sheetDestino.getLastRowNum(); i++) {

                Row row = sheetDestino.getRow(i);
                if (row == null) continue;

                for (Cell cell : row) {

                    boolean esAño = false;

                    if (cell.getCellType() == CellType.STRING &&
                            cell.getStringCellValue().trim().equals(ca)) {
                        esAño = true;
                    }

                    if (cell.getCellType() == CellType.NUMERIC &&
                            (int) cell.getNumericCellValue() == año) {
                        esAño = true;
                    }

                    if (!esAño) continue;

                    int filaActual = cell.getRowIndex();

                    if (filaDestinoAnterior != null) {

                        if (filaActual == 52) {
                            filaOrigenBase = 41;
                            colOrigenBase = colBaseFila53 + (contadorFila53 * 3);
                            contadorFila53++;
                        }
                        else if (filaActual >= 114) {
                            filaOrigenBase += 8;
                            colOrigenBase = 1;
                        }
                        else if (filaActual == filaDestinoAnterior) {
                            colOrigenBase++;
                        }
                        else {
                            filaOrigenBase += 8;
                            colOrigenBase = 1;
                        }
                    }

                    borrarColumna(sheetDestino, cell);
                    moverBloqueIzquierda(sheetDestino, cell);

                    rellenar2025YDatos(
                            sheetDestino,
                            cell,
                            sheetOrigen,
                            filaOrigenBase,
                            colOrigenBase
                    );

                    filaDestinoAnterior = filaActual;
                }
            }

            try (FileOutputStream fos = new FileOutputStream(archivoNuevo)) {
                wbDestino.write(fos);
            }

            System.out.println("Excel nuevo generado correctamente");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
