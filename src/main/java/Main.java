import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Random;

public class Main {

    private static final String[] name = {"Андрей", "Илья", "Никита", "Артем", "Евгений", "Петр", "Вадим"};
    private static final String[] surname = {"Иванов", "Смирнов", "Петров", "Попов", "Васильев", "Медведев", "Сидоров", "Антонов"};
    private static final String[] classLetter = {"A", "B", "C", "D"};
    private static final int daysPlan = 22;
    private static final int dayCost = 50;

    private static Workbook workbook;
    private static Sheet sheet;
    private static Row[] rows;
    private static Random random;

    public static void main(String[] args) throws IOException {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Начальная школа");
        rows = new Row[66];
        random = new Random();

        for (int i = 0; i < rows.length; i++) {
            rows[i] = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                rows[i].createCell(j);
            }
        }

        merge();
        fill();
        setLayout();

        workbook.write(new FileOutputStream("result.xlsx"));
        workbook.close();
    }

    private static void merge() {
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

    private static void fill() { //a0 b1 c2 d3 e4 f5 g6 h7 i8 j9
        String[] randomData = getRandomData();
        String[] split;
        rows[0].getCell(7).setCellValue("УТВЕРЖДАЮ:");

        rows[1].getCell(7).setCellValue("Директор");

        rows[2].getCell(8).setCellValue("(сокращенное наименование образовательного учреждения)");

        rows[3].getCell(7).setCellValue("_____________");
        rows[3].getCell(8).setCellValue("___________________________");

        rows[4].getCell(7).setCellValue("(подпись)");
        rows[4].getCell(8).setCellValue("(расшифровка подписи)");

        rows[6].getCell(7).setCellValue("14.05.2022");

        rows[7].getCell(7).setCellValue("М.П.");

        rows[9].getCell(0).setCellValue("Отчёт о фактическом предоставленном бесплатном питании");

        rows[10].getCell(0).setCellValue("за период с 01.05.2022 по 31.05.2022");

        rows[11].getCell(0).setCellValue("_________________________________________________________________________________________________________________");

        rows[12].getCell(0).setCellValue("(сокращенное наименование образовательного учреждения)");

        rows[14].getCell(0).setCellValue("№ п/п");
        rows[14].getCell(1).setCellValue("№ счета");
        rows[14].getCell(2).setCellValue("Класс");
        rows[14].getCell(3).setCellValue("Ф.И. ребенка");
        rows[14].getCell(4).setCellValue("Дни посещения");
        rows[14].getCell(6).setCellValue("Остаток на начало месяца, руб.");
        rows[14].getCell(7).setCellValue("Поступило в текущем месяце на питание, руб.");
        rows[14].getCell(8).setCellValue("Израсходовано в текущем месяце на питание, руб.");
        rows[14].getCell(9).setCellValue("Остаток на конец месяца, руб.");

        rows[15].getCell(4).setCellValue("плановые");
        rows[15].getCell(5).setCellValue("Фактические");

        for (int i = 16; i < 56; i++) {
            split = randomData[i - 16].split("\\s");
            for (int j = 0; j < 2; j++) {
                rows[i].getCell(j).setCellValue(Double.parseDouble(split[j]));
            }
            rows[i].getCell(2).setCellValue(split[2]);
            rows[i].getCell(3).setCellValue(split[3] + " " + split[4]);
            for (int j = 4; j < 10; j++) {
                rows[i].getCell(j).setCellValue(Double.parseDouble(split[j + 1]));
            }
        }

        rows[56].getCell(3).setCellValue("Итого:");

        int value;
        for (int j = 4; j < 10; j++) {
            value = 0;
            for (int i = 16; i < 56; i++) {
                value += rows[i].getCell(j).getNumericCellValue();
            }
            rows[56].getCell(j).setCellValue(value);
        }

        rows[58].getCell(1).setCellValue("Отчет составлен в двух экземплярах.");

        rows[60].getCell(1).setCellValue("Подписи сторон:");

        rows[62].getCell(1).setCellValue("Лицо, ответственное за организацию");
        rows[62].getCell(4).setCellValue("___________________");
        rows[62].getCell(7).setCellValue("_________________________");

        rows[63].getCell(4).setCellValue("(подпись)");
        rows[63].getCell(7).setCellValue("(Ф.И.О.)");

        rows[64].getCell(1).setCellValue("Заведующий производством");
        rows[64].getCell(4).setCellValue("___________________");
        rows[64].getCell(7).setCellValue("_________________________");

        rows[65].getCell(4).setCellValue("(подпись)");
        rows[65].getCell(7).setCellValue("(Ф.И.О.)");
    }

    private static void setLayout() {
        CellStyle cellStyleBorder = workbook.createCellStyle();
        CellStyle cellStyleDefault = workbook.createCellStyle();
        CellStyle cellStyleCenter = workbook.createCellStyle();
        CellStyle cellStyleLeft = workbook.createCellStyle();
        CellStyle cellStyleWrap = workbook.createCellStyle();

        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 10);
        font.setFontName("Times new roman");

        cellStyleBorder.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        cellStyleBorder.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStyleBorder.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStyleBorder.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cellStyleBorder.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyleBorder.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyleBorder.setFont(font);

        cellStyleDefault.setFont(font);

        cellStyleCenter.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyleCenter.setFont(font);

        cellStyleLeft.setAlignment(CellStyle.ALIGN_LEFT);
        cellStyleLeft.setFont(font);
        cellStyleLeft.setBorderBottom(XSSFCellStyle.BORDER_THIN);

        cellStyleWrap.setWrapText(true);
        cellStyleWrap.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyleWrap.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        cellStyleWrap.setFont(font);

        rows[2].getCell(8).setCellStyle(cellStyleWrap);
        rows[2].getCell(9).setCellStyle(cellStyleWrap);
        cellStyleWrap.setBorderBottom(XSSFCellStyle.BORDER_THIN);
        cellStyleWrap.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStyleWrap.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStyleWrap.setBorderRight(XSSFCellStyle.BORDER_THIN);

        for (Row row : rows) {
            for (int j = 0; j < 10; j++) {
                row.getCell(j).setCellStyle(cellStyleDefault);
            }
        }

        for (int i = 14; i < 56; i++) {
            for (int j = 0; j < 10; j++) {
                rows[i].getCell(j).setCellStyle(cellStyleBorder);
            }
        }
        for (int i = 3; i < 10; i++) {
            rows[56].getCell(i).setCellStyle(cellStyleBorder);
        }

        for (int i = 0; i < 10; i++) {
            rows[14].getCell(i).setCellStyle(cellStyleWrap);
            rows[15].getCell(i).setCellStyle(cellStyleWrap);
        }

        rows[4].getCell(7).setCellStyle(cellStyleCenter);
        rows[4].getCell(8).setCellStyle(cellStyleCenter);

        rows[9].getCell(0).setCellStyle(cellStyleCenter);
        rows[10].getCell(0).setCellStyle(cellStyleCenter);
        rows[12].getCell(0).setCellStyle(cellStyleCenter);

        for (int i = 16; i < 56; i++) {
            rows[i].getCell(3).setCellStyle(cellStyleLeft);
        }

        rows[63].getCell(4).setCellStyle(cellStyleCenter);
        rows[63].getCell(7).setCellStyle(cellStyleCenter);

        rows[65].getCell(4).setCellStyle(cellStyleCenter);
        rows[65].getCell(7).setCellStyle(cellStyleCenter);

        sheet.setColumnWidth(0, 7 * 256);
        sheet.setColumnWidth(1, 12 * 256);
        sheet.setColumnWidth(2, 7 * 256);
        sheet.setColumnWidth(3, 25 * 256);
        sheet.setColumnWidth(4, 12 * 256);
        sheet.setColumnWidth(5, 13 * 256);
        sheet.setColumnWidth(6, 15 * 256);
        sheet.setColumnWidth(7, 15 * 256);
        sheet.setColumnWidth(8, 15 * 256);
        sheet.setColumnWidth(9, 15 * 256);

        rows[2].setHeightInPoints(32);
        rows[3].setHeightInPoints(32);
        rows[14].setHeightInPoints(54);
        rows[15].setHeightInPoints(32);
        rows[62].setHeightInPoints(32);
        rows[64].setHeightInPoints(32);
    }

    private static String[] getRandomData() {
        String[] randomData = new String[40];
        int daysFact;

        for (int i = 1; i <= randomData.length; i++) {
            daysFact = random.nextInt(daysPlan);
            randomData[i - 1] =
                (i + " ") + 66101 + (((i / 10) + "") + ((i % 10) + "")) + " 4" + classLetter[i % 4] + " " +
                surname[random.nextInt(surname.length)] + " " + name[random.nextInt(name.length)] +
                " " + daysPlan + " " + daysFact + " " + 0 + " " + daysPlan * dayCost + " " +
                daysFact * dayCost + " " + (daysPlan - daysFact) * 50;
        }
        return randomData;
    }
}