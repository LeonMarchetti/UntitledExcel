import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Main {

    private static List<List<Cell>> cells = new ArrayList<List<Cell>>();

    public static void main(String[] args) {
        load("workbooks/Untitled1.xls");
        save("workbooks/Workbook1.xls");
    }

    public static void load(String fileLocation) {
        try {
            Workbook workbook = WorkbookFactory.create(new File(fileLocation));

            for (Sheet sheet : workbook) {
                int rowNumber = 0;
                for (Row row : sheet) {
                    System.out.printf("Fila %d\n", rowNumber);
                    List<Cell> thisRow = new ArrayList<Cell>();
                    int cellNumber = 0;
                    for (Cell cell : row) {
                        thisRow.add(cell);
                        System.out.printf("\tCelda %d: ", cellNumber);
                        switch (cell.getCellTypeEnum()) {
                            case BLANK:
                                System.out.printf("<vacío>\n");
                                break;
                            case BOOLEAN:
                                boolean valor = cell.getBooleanCellValue();
                                if (valor) {
                                    System.out.printf("<True>\n");
                                } else {
                                    System.out.printf("<False>\n");
                                }
                                break;
                            case ERROR:
                                Byte error = cell.getErrorCellValue();
                                System.out.printf("Error: %d\n", error);
                                break;
                            case FORMULA:
                                String formula = cell.getCellFormula();
                                System.out.printf("%s\n", formula);
                                break;
                            case NUMERIC:
                                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                                    Date fecha = cell.getDateCellValue();
                                    System.out.printf("%s\n", fecha.toString());
                                } else {
                                    double numero = cell.getNumericCellValue();
                                    System.out.printf("%f\n", numero);
                                }
                                break;
                            case STRING:
                                String string = cell.getStringCellValue();
                                System.out.printf("%s\n", string);
                                break;
                            case _NONE:
                                System.out.printf("<none>\n");
                                break;
                            default:
                                System.out.printf("<default>\n");
                                break;
                        }
                        cellNumber++;
                    }
                    rowNumber++;
                    cells.add(thisRow);
                }
            }
            workbook.close();

        } catch (IOException e) {
            e.printStackTrace();
        } catch (EncryptedDocumentException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

    }

    public static void save(String fileLocation) {

    }
}
