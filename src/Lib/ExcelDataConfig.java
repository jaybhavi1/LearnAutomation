package Lib;

//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;

/**
 * @ author Jay Vaghani on 19/04/2017.
 */
public class ExcelDataConfig
{
    HSSFWorkbook wb;
    HSSFSheet sheet1;
    public ExcelDataConfig(String excelPath)
    {
        try {
            File scr = new File(excelPath);

            FileInputStream fis = new FileInputStream(scr);

            wb = new HSSFWorkbook(fis);

        }catch (Exception e){
            System.out.println(e.getMessage());
        }
    }

    public String getData(int sheetNumber,int row,int column)
    {
        sheet1 = wb.getSheetAt(sheetNumber);
        String data = sheet1.getRow(row).getCell(column).getStringCellValue();
        return data;

    }











}
