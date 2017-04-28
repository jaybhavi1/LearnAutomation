package ReadExcelData;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;

/**
 * @ author Jay Vaghani on 19/04/2017.
 */
public class ReadExcel1 {
    public static void main(String[] args) throws Exception {
        File scr = new File("C:\\UniqueTesting\\Home work given\\TestData.xlsx");

        FileInputStream fis = new FileInputStream(scr);

        XSSFWorkbook wb = new XSSFWorkbook(fis);

        XSSFSheet sheet1 = wb.getSheetAt(0);

        int rowCount = sheet1.getLastRowNum();

        System.out.println("Total rows is "+rowCount);

        for (int i = 0; i <rowCount ; i++)
        {
          String data0 =  sheet1.getRow(i).getCell(0).getStringCellValue();
            System.out.println("Data from Excel is "+i+ "  is " +data0);
            
        }

    }
}