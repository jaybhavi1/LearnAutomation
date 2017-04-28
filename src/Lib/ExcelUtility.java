package Lib;

import org.apache.poi.hssf.usermodel.HSSFCell;

/**
 * @ author Jay Vaghani on 19/04/2017.
 */

import java.io.FileInputStream;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

    public class ExcelUtility
    {
        private static HSSFWorkbook ExcelWBook;
        private static HSSFSheet ExcelWSheet;

        /*
         * Set the File path,open Excel file
         * @Jay Vaghani-Excel Path and Sheet Name
         */
        public static void setExcelFile(String path,String sheetName) throws Exception{
            try{
                //Open the Excel file
                FileInputStream ExcelFile = new FileInputStream(path);

                // Access the excel data sheet
                ExcelWBook = new HSSFWorkbook(ExcelFile);
                ExcelWSheet = ExcelWBook.getSheet(sheetName);
            }catch (Exception e){
                throw (e);
            }
        }

        public static String[][] getTestData(String tableName){
            String[][] testData = null;

            try{
                HSSFCell[] boundaryCells = findCells(tableName);
                HSSFCell startCell = boundaryCells[0];

                HSSFCell endCell = boundaryCells[1];

                int startRow = startCell.getRowIndex()+1;
                int endRow = endCell.getRowIndex()-1;
                int startCol = startCell.getColumnIndex()+1;
                int endCol = endCell.getColumnIndex()-1;

                testData = new String[endRow - startRow + 1][endCol - startCol + 1];

                for(int i=startRow; i<endRow+1; i++){
                    for(int j=startCol; j<endCol+1; j++){
                        testData[i-startRow][j-startCol] = ExcelWSheet.getRow(i).getCell(j).getStringCellValue();
                    }
                }
            }catch (Exception e){
                e.printStackTrace();

            }
            return testData;
        }

        public static HSSFCell[] findCells(String tableName){
            String pos = "begin";
            HSSFCell[] cells = new HSSFCell[2];

            for(Row row : ExcelWSheet){
                for(Cell cell : row){
                    if(tableName.equals(cell.getStringCellValue())){
                        if(pos.equalsIgnoreCase("begin")){
                            cells[0] = (HSSFCell) cell;
                            pos = "end";
                        }else {
                            cells[1] = (HSSFCell) cell;
                        }
                    }
                }

            }
            return cells;
        }
    }


