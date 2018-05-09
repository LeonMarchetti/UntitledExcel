import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

    public static void main(String[] args) {
        load("workbooks/libro1.xls");
        saveH("workbooks/libro2.xls");
        saveX("workbooks/libro3.xlsx");
    }

    static void load(String archivo) {

        try {
            Workbook libro = WorkbookFactory.create(new File(archivo));

            for (Sheet hoja : libro) {
                int numFila = 0;
                for (Row fila : hoja) {
                    System.out.printf("Fila %d\n", numFila);
                    int numCelda = 0;
                    for (Cell celda : fila) {
                        System.out.printf("\tCelda %d: ", numCelda);
                        switch (celda.getCellTypeEnum()) {
                            case BLANK:
                                System.out.printf("<vacío>\n");
                                break;
                            case BOOLEAN:
                                boolean valor = celda.getBooleanCellValue();
                                if (valor) {
                                    System.out.printf("<True>\n");
                                } else {
                                    System.out.printf("<False>\n");
                                }
                                break;
                            case ERROR:
                                Byte error = celda.getErrorCellValue();
                                System.out.printf("Error: %d\n", error);
                                break;
                            case FORMULA:
                                String formula = celda.getCellFormula();
                                System.out.printf("%s\n", formula);
                                break;
                            case NUMERIC:
                                if (DateUtil.isCellDateFormatted(celda)) {
                                    Date fecha = celda.getDateCellValue();
                                    System.out.printf("%s\n", fecha.toString());
                                } else {
                                    double numero = celda.getNumericCellValue();
                                    System.out.printf("%f\n", numero);
                                }
                                break;
                            case STRING:
                                String texto = celda.getStringCellValue();
                                System.out.printf("%s\n", texto);
                                break;
                            case _NONE:
                                System.out.printf("<none>\n");
                                break;
                            default:
                                System.out.printf("<default>\n");
                                break;
                        }
                        numCelda++;
                    }
                    numFila++;
                }
            }
            libro.close();

        } catch (EncryptedDocumentException |
                 InvalidFormatException |
                 IOException e) {
            e.printStackTrace();
        } finally {
            System.out.printf("load:\tTerminado\n\n");
        }

    }

    static void save(Workbook libro, String archivo) {
        Sheet hoja = libro.createSheet("Hoja 1");

        // Estilos de celdas:
        Font fuente1 = libro.createFont();
        fuente1.setBold(true);
        fuente1.setFontName("Liberation Mono");

        CellStyle estilo1 = libro.createCellStyle();
        estilo1.setBorderBottom(BorderStyle.THIN);
        estilo1.setBorderLeft(BorderStyle.THIN);
        estilo1.setBorderRight(BorderStyle.THIN);
        estilo1.setBorderTop(BorderStyle.THIN);
        estilo1.setFont(fuente1);


        Font fuente2 = libro.createFont();
        fuente2.setFontName("Liberation Mono");

        CellStyle estilo2 = libro.createCellStyle();
        estilo2.setAlignment(HorizontalAlignment.CENTER);
        estilo2.setBorderBottom(BorderStyle.THIN);
        estilo2.setBorderLeft(BorderStyle.THIN);
        estilo2.setBorderRight(BorderStyle.THIN);
        estilo2.setBorderTop(BorderStyle.THIN);
        estilo2.setFont(fuente2);

        // Creo las celdas:
        for (int i = 1; i <= 10; i++) {
            Row fila = hoja.createRow(i);
            Cell celdaTitulo = fila.createCell(0);
            celdaTitulo.setCellValue(String.format("Fila %d", i));
            celdaTitulo.setCellStyle(estilo1);

            for (int j = 1; j <= 10; j++) {
                Cell celda = fila.createCell(j);
                celda.setCellValue(i * j);
                celda.setCellStyle(estilo2);
            }
        }

        // Ajusto el ancho de las columnas:
        for (int i = 0; i <= 10; i++) {
            hoja.autoSizeColumn(i);
        }

        try {
            libro.write(new FileOutputStream(archivo));
            libro.close();
        } catch (IOException e) {
            System.out.printf("%s\n", e.getMessage());
        }
    }

    static void saveH(String archivo) {
        save(new HSSFWorkbook(), archivo);
        System.out.printf("saveH:\tTerminado\n\n");
    }

    static void saveX(String archivo) {
        save(new XSSFWorkbook(), archivo);
        System.out.printf("saveX:\tTerminado\n\n");
    }

}
