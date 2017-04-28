package ReadExcelData;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.io.File;
import java.io.FileInputStream;

/**
 * @ author Jay Vaghani on 19/04/2017.
 */
public class ReadNriExcel {
    public static void main(String[] args) throws Exception {
        File scr = new File("C:\\UniqueTesting\\Home work given\\NRI.xls");

        FileInputStream fis = new FileInputStream(scr);

        HSSFWorkbook wb = new HSSFWorkbook(fis);

        HSSFSheet sheet1 = wb.getSheetAt(0);

        int rowCount = sheet1.getLastRowNum();

        System.out.println("Total rows is " + rowCount);
        try {

            for (int i = 0; i < rowCount; i++) {
                Row row = sheet1.getRow(i);
                //Create a loop to print cell values in a row
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    //Print Excel data in console

                    if (row.getCell(j).getCellType() == XSSFCell.CELL_TYPE_STRING) {
                        System.out.format("%-22s", "|\t" + row.getCell(j).getStringCellValue());

                    } else if ((row.getCell(j).getCellType() == XSSFCell.CELL_TYPE_NUMERIC)) {
                        System.out.format("%-22s", ":\t" + row.getCell(j).getNumericCellValue());

                    } else if (row.getCell(j).getCellType() == XSSFCell.CELL_TYPE_FORMULA) {
                        System.out.format("%-22s", ":\t" + row.getCell(j).getNumericCellValue());

                    }
                }
                System.out.println();
            }
        } catch (Exception e) {

        }
    }
}



