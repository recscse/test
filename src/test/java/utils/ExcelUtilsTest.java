package utils;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

public class ExcelUtilsTest {
    public static void main(String[] args) throws IOException {
        long start = System.nanoTime();
        String excelPath = "./data/Spiriva_March2021_Excel_TEST.xlsx";
        String sheetName = "Export";
        Excelutils excel = new Excelutils(excelPath, sheetName);

        int row = excel.getRowCount();
        for (int i = 1; i < row; i++) {
            System.out.println(i + ": ");
            for (int j = 0; j <= row; j++) {
                excel.getCellData(i, j);
            }
            System.out.println("");

        }
        long end = System.nanoTime();
        long exec = end - start;
        double inSeconds = (double)exec / 1_000_000_000.0;
        double min = inSeconds/60;
        System.out.println("Time Taken in Min. :" +TimeUnit.MINUTES.convert(exec, TimeUnit.NANOSECONDS));
        System.out.println("The program takes "+exec+" nanoseconds that is "+inSeconds+" seconds to execute." + "min :"+ min );

       excel.getCellData(1,0);
       excel.getCellData(1,1);
//        excel.getCellData(1,2);
//        excel.getCellData(1,3);
//        excel.getCellData(2,0);
//        excel.getCellData(2,1);
//        excel.getCellData(2,2);
//        excel.getCellData(2,3);
    }
}
