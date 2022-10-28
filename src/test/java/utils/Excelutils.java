package utils;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;


public class Excelutils {
    static XSSFWorkbook workbook;
    static XSSFSheet sheet;
//    static HSSFWorkbook workbook;
//    static HSSFSheet sheet;

    public Excelutils(String excelPath, String sheetName) {
        try {
            InputStream file = new FileInputStream(excelPath);
            workbook = new XSSFWorkbook(excelPath);
            sheet = workbook.getSheet(sheetName);
        } catch (Exception e) {
            System.out.println(e.getCause());
            System.out.println(e.getStackTrace());
            System.out.println(e.getMessage());
        }


    }

   public static void main(String [] args) throws IOException {
      getRowCount();
      getCellData(int );

   }

    public static void getCellData(int rowNum, int colNum) {
//            String excelPath = "./data/TestData.xlsx";
//            XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
//            XSSFSheet sheet = workbook.getSheet("Sheet1");
        DataFormatter dataFormatter = new DataFormatter();
        Object value = dataFormatter.formatCellValue(sheet.getRow(rowNum).getCell(colNum));
//            String value= sheet.getRow(1).getCell(2).getStringCellValue();
        List<Object> rowdata = new ArrayList<>();
        rowdata.add(value);
//            System.out.println(value);
//        System.out.println("List data:" +rowdata);
        System.out.print( rowdata);
//        System.out.println("");
    }

    public static int getRowCount() throws IOException {
//            String excelPath = "./data/TestData.xlsx";
//            XSSFWorkbook workbook = new XSSFWorkbook(excelPath);
//            XSSFSheet sheet = workbook.getSheet("Sheet1");
        int rowCount = sheet.getPhysicalNumberOfRows();
        System.out.println("No of rows :" + rowCount);
        return rowCount;

    }
}
