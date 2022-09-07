import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Date;

public class Main {

    public static void main(String[] args) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("sheetOne");

        Row[] rows = new Row[66];

        for (int i = 0; i < rows.length; i++) {
            rows[i] = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                rows[i].createCell(j);
            }
        }

        merge(sheet);

        workbook.write(new FileOutputStream("kekis.xlsx"));
        workbook.close();
    }

    private static void merge(Sheet sheet) {
        sheet.addMergedRegion(new CellRangeAddress(1,1,8,9));
        sheet.addMergedRegion(new CellRangeAddress(2,2,8,9));
        sheet.addMergedRegion(new CellRangeAddress(3,3,8,9));
        sheet.addMergedRegion(new CellRangeAddress(4,4,8,9));

        sheet.addMergedRegion(new CellRangeAddress(9,9,0,9));
        sheet.addMergedRegion(new CellRangeAddress(10,10,0,9));
        sheet.addMergedRegion(new CellRangeAddress(11,11,0,9));
        sheet.addMergedRegion(new CellRangeAddress(12,12,0,9));

        sheet.addMergedRegion(new CellRangeAddress(14,15,0,0));
        sheet.addMergedRegion(new CellRangeAddress(14,15,1,1));
        sheet.addMergedRegion(new CellRangeAddress(14,15,2,2));
        sheet.addMergedRegion(new CellRangeAddress(14,15,3,3));
        sheet.addMergedRegion(new CellRangeAddress(14,14,4,5));
        sheet.addMergedRegion(new CellRangeAddress(14,15,6,6));
        sheet.addMergedRegion(new CellRangeAddress(14,15,7,7));
        sheet.addMergedRegion(new CellRangeAddress(14,15,8,8));
        sheet.addMergedRegion(new CellRangeAddress(14,15,9,9));

        sheet.addMergedRegion(new CellRangeAddress(62,62,4,5));
        sheet.addMergedRegion(new CellRangeAddress(63,63,4,5));
        sheet.addMergedRegion(new CellRangeAddress(64,64,4,5));
        sheet.addMergedRegion(new CellRangeAddress(65,65,4,5));
        sheet.addMergedRegion(new CellRangeAddress(62,62,7,8));
        sheet.addMergedRegion(new CellRangeAddress(63,63,7,8));
        sheet.addMergedRegion(new CellRangeAddress(64,64,7,8));
        sheet.addMergedRegion(new CellRangeAddress(65,65,7,8));
    }
}