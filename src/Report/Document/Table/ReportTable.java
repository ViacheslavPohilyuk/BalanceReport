package Report.Document.Table;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.*;

import java.util.Date;
import java.util.Map;
import java.util.Random;

/**
 * Created by mac on 14.01.17.
 */
public class ReportTable {
    private static Row row;

    private int balance;
    private int monthDayCount;
    private int indexMonth;
    private HSSFSheet sheet;
    private Map<String, CellStyle> allStyles;

    private int totalSum[] = new int[2];  // В цьому масиві буде зберігатися кінцева сума грошей яких
                                          // було отримано в totalSum[0] і виплачено в totalSum[1]

    public ReportTable(HSSFSheet sheet, Map<String, CellStyle> allStyles, int balance, int monthDayCount, int indexMonth)
    {
        this.monthDayCount = monthDayCount;
        this.indexMonth = indexMonth;
        this.balance = balance;
        this.sheet = sheet;
        this.allStyles = allStyles;
    }

    public void table(int monthDayCount, int indexMonth) {
        TotalBalance TB = new TotalBalance(sheet, allStyles, balance,monthDayCount);

        firstRowTable();     // Перший рядок балансового звіту
        tableCellsValues();  // Створення і заповненя всіх рядків таблиці звіту
        lastRowTable();      // Cтиль для останього рядка звіту
        TB.totalBalance(totalSum); // Позначення виплачених і отриманих сум за місяц.
                                   // А також остаточний баланс
    }

    private void firstRowTable() {
        CellStyle style;
        String[] namesRowCells = {"Дата", "Операція", "Отримано", "Виплачено", "Баланс"};  // текст комірок першого рядка звіту
        Integer borderWidthRowCells[] = {512 * 6, 2048 * 6, 768 * 6, 768 * 6, 768 * 6}; // ширина кожного з стовбців таблиці звіту

        int column = 1;
        Row rowReport = sheet.createRow(11);
        style = allStyles.get("Перший рядок таблиці");
        for (int i = 0; i < 5; i++) {
            Integer currentCell = column + i;
            rowReport.createCell(currentCell).setCellValue(namesRowCells[i]);
            sheet.setColumnWidth(currentCell, borderWidthRowCells[i]);
            rowReport.getCell(currentCell).setCellStyle(style);
        }
    }

    private void lastRowTable() {
        CellStyle style;
        CellStyle styleDate;
        int startRow = 12;

        style = allStyles.get("Межі таблиці звіту");
        styleDate = allStyles.get("Перший стовбець таблиці");;
        style.setBorderBottom(CellStyle.BORDER_DOTTED);
        styleDate.setBorderBottom(CellStyle.BORDER_DOTTED);
        row = sheet.getRow(monthDayCount + startRow - 1);

        for (int j = 1; j < 6; j++)
            row.getCell(j).setCellStyle(style);
        row.getCell(1).setCellStyle(styleDate);
    }

    private void tableCellsValues() {
        CellStyle style;
        CellStyle styleDate;

        style = allStyles.get("Межі таблиці звіту");
        styleDate = allStyles.get("Перший стовбець таблиці");;

        String formula = "E8"; // комірка excel, в якій знаходится розмір балансу
        int startRow = 12;     // Рядок на якому знаходится таблиця звіту
        for (int i = startRow; i < (monthDayCount + startRow); i++) {
            row = sheet.createRow(i);

            Random randEvent = new Random();

            // випадкове число типу Int в діапазоні від 0 до 19
            int randInt = randEvent.nextInt(19);
            for (int j = 1; j < 6; j++) {
                switch (j) {
                    case 1: {
                        row.createCell(j).setCellStyle(styleDate);
                        row.getCell(j).setCellValue(new Date(110, indexMonth, (i - startRow + 1)));
                        break;
                    }
                    case 2: {
                        String event;  // рядок, що зберігає випадкову подію
                        event = randEvent(randInt);
                        row.createCell(j).setCellStyle(style);
                        row.getCell(j).setCellValue(event);
                        break;
                    }
                    case 3: {
                        row.createCell(j).setCellStyle(style);
                        if (randInt < 10)
                            randCell(randInt, j);
                        break;
                    }
                    case 4: {
                        row.createCell(j).setCellStyle(style);
                        if (randInt >= 10)
                            randCell(randInt, j);
                        break;
                    }
                    case 5: {
                        Cell sum = row.createCell(j);
                        String summand = Integer.toString(i + 1);

                        // Формула, яка зменшить або збільшіть розмір балансу в залежності від характеру події
                        formula += (randInt < 10) ? ("+D" + summand) : ("-E" + summand);
                        sum.setCellFormula(formula);
                        sum.setCellStyle(style);
                        formula = "F" + summand;
                        break;
                    }
                }
            }
        }
    }

    private void randCell(int randInt, int index) {
        int oneToTenVal = ((randInt % 10) + 1);
        int randSum = (oneToTenVal * 50);
        totalSum[randInt / 10] += randSum;
        row.getCell(index).setCellValue(oneToTenVal * 50); // тут у комірку записується випадкове значення від 50 до 500
    }

    private String randEvent(int randomStr) {
        String events[] = {
                "Позитивна подія 1", "Позитивна подія 2",
                "Позитивна подія 3", "Позитивна подія 4",
                "Позитивна подія 5", "Позитивна подія 6",
                "Позитивна подія 7", "Позитивна подія 8",
                "Позитивна подія 9", "Позитивна подія 10",
                "Негативна подія 1", "Негативна подія 2",
                "Негативна подія 3", "Негативна подія 4",
                "Негативна подія 5", "Негативна подія 6",
                "Негативна подія 7", "Негативна подія 8",
                "Негативна подія 9", "Негативна подія 10"
        };
        return events[randomStr];
    }

    public int[] getTotalSum() {
        int result[] = new int[3];
        result[0] = totalSum[0];
        result[1] = totalSum[1];
        result[2] = (balance + totalSum[0] - totalSum[1]);
        return result;
    }

}