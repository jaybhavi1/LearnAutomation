package ReadExcelData;

//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;
//
//import java.io.File;
//import java.io.FileInputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

/**
 * @ author Jay Vaghani on 19/04/2017.
 */
public class ReadExcel
{
    public static void main(String[] args) throws Exception {
        File scr = new File("C:\\UniqueTesting\\Home work given\\TestData.xlsx");

        FileInputStream fis = new FileInputStream(scr);

        HSSFWorkbook wb1 = new HSSFWorkbook(); // for .xls file

        XSSFWorkbook wb = new XSSFWorkbook(fis); // .xlsx file

        XSSFSheet sheet1 = wb.getSheetAt(0);

        int a = sheet1.getLastRowNum();
        System.out.println(a);

        for (int i = 0; i <= sheet1.getLastRowNum(); i++) {


            for (int j = 0; j < sheet1.getRow(i).getLastCellNum(); j++) {
                if((sheet1.getRow(i).getCell(j).getCellType()== XSSFCell.CELL_TYPE_NUMERIC)){
                System.out.print(sheet1.getRow(i).getCell(j).getNumericCellValue()+" ");
                }else {
                    System.out.print(sheet1.getRow(i).getCell(j).getStringCellValue() + " ");
                }
            }
            System.out.println("");
        }

//        String data0 = sheet1.getRow(3).getCell(1).getStringCellValue();
//
//        System.out.println("Data from Excel is "+data0);
//
//        String data1 = sheet1.getRow(0).getCell(1).getStringCellValue();
//
//        System.out.println("Data from Excel is "+data1);

        wb.close();



    }
}
