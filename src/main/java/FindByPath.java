import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;

/**
 * Created by IBM on 7/15/2018.
 */
public class FindByPath {
    String filePath= "CSV1.5.3.15.xlsx";
    ArrayList<String> sheets=new ArrayList<String>();
    Workbook workbook;
    String toFind="Claim/ClaimDetails/ClaimContractInfo/ContractCode";

    void openFile(){
        ClassLoader classLoader = getClass().getClassLoader();
        File file = new File(classLoader.getResource(filePath).getFile());

        //FileInputStream fileStream = new FileInputStream(file);
        try {
            workbook = new XSSFWorkbook(file);
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    ArrayList<Sheet> getFiltredSheets(){
        for (int i=0; i<workbook.getNumberOfSheets(); i++){
            System.out.println(workbook.getSheetName(i));
        }

        System.out.println(workbook.getNumberOfSheets());
        return null;
    }

    void findExcelPosition(){
        DataFormatter formatter = new DataFormatter();
        for (int i=0; i<workbook.getNumberOfSheets(); i++){
            Sheet sheet = workbook.getSheetAt(i);
            for (Row row : sheet) {
                for (Cell cell : row) {
                    CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());

                    // get the text that appears in the cell by getting the cell value and applying any data formats (Date, 0.00, 1.23e9, $1.23, etc)
                    String text = formatter.formatCellValue(cell);

                    // is it an exact match?
                    if (toFind.equals(text)) {
                        System.out.println("sheet: "+ workbook.getSheetName(i) + " equals: " + cellRef.formatAsString());
                    }
                    // is it a partial match?
                    else if (text.contains(toFind)) {
                        Row rowFound = sheet.getRow(row.getRowNum());

                        System.out.println( "sheet: "+ workbook.getSheetName(i) +
                                            " contains: " + cellRef.formatAsString() +
                                            " FirtCell="+row.getCell(1));
                    }
                }
            }
        }
    }
}
