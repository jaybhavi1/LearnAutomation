package ReadExcelData;

import Lib.ExcelDataConfig;

/**
 * @ author Jay Vaghani on 19/04/2017.
 */
public class ReadExcelData
{
    public static void main(String[] args) {

        ExcelDataConfig excel = new ExcelDataConfig("C:\\UniqueTesting\\Home work given\\TestData.xlsx");
        System.out.println(excel.getData(1,0,1));
    }
}
