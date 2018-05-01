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

public class Main {

    public static void main(String[] args) {
        // load("workbooks/Untitled1.xls");
        save("workbooks/Workbook.xls");
    }

    public static void load(String archivo) {
        try {
            Workbook planilla = WorkbookFactory.create(new File(archivo));

            for (Sheet hoja : planilla) {
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
            planilla.close();

        } catch (IOException |
                 EncryptedDocumentException |
                 InvalidFormatException e) {
            e.printStackTrace();
        }

    }

    public static void save(String archivo) {
        Workbook planilla = new HSSFWorkbook();
        Sheet hoja = planilla.createSheet("Hola mundo");

        Font fuente = planilla.createFont();
        fuente.setBold(true);

        CellStyle estilo1 = planilla.createCellStyle();
        estilo1.setBorderBottom(BorderStyle.THIN);
        estilo1.setBorderLeft(BorderStyle.THIN);
        estilo1.setBorderRight(BorderStyle.THIN);
        estilo1.setBorderTop(BorderStyle.THIN);
        estilo1.setFont(fuente);

        CellStyle estilo2 = planilla.createCellStyle();
        estilo2.setBorderBottom(BorderStyle.THIN);
        estilo2.setBorderLeft(BorderStyle.THIN);
        estilo2.setBorderRight(BorderStyle.THIN);
        estilo2.setBorderTop(BorderStyle.THIN);
        estilo2.setAlignment(HorizontalAlignment.CENTER);

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

        try {
            planilla.write(new FileOutputStream(archivo));
            planilla.close();
        } catch (IOException e) {
            System.out.printf("%s\n", e.getMessage());
        }

        System.out.printf("Terminado\n");

    }
}
