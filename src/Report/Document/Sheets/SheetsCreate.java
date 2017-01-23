package Report.Document.Sheets;

import Report.Document.Sheets.YearReport.YearReportTable;
import Report.Document.Table.OtherTableElements;
import Report.Document.Table.ReportTable;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.Map;

/**
 * Created by mac on 16.01.17.
 */
public class SheetsCreate {
    int yearMoneyGet = 0;
    int yearMoneyGive = 0;
    int yearMoneyBalance;
    private final String[] months = {"Січень", "Лютий", "Березень", "Квітень", "Травень", "Червень",
                                     "Липень", "Серпень", "Вересень", "Жовтень", "Листопад", "Грудень"};
    private final int[] monthsDaysCount = {31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31};

    public SheetsCreate(HSSFWorkbook workbook) {
        Styles objStyles = new Styles(workbook);
        Map<String,CellStyle> allStyles = objStyles.styles();

        int yearBalance[] = new int[3];

        yearBalance[2] = 5000; // Початковий баланс

        for(int indexMonth = 0; indexMonth < 12; indexMonth++) {
            //Створення сторінок на кожний місяць
            yearBalance = sheetMonthCreate(workbook, months[indexMonth], monthsDaysCount[indexMonth], indexMonth, yearBalance[2]);
            yearMoneyGet += yearBalance[0];  // в цю змінну додається кількість отриманних грошей за 12 місяців
            yearMoneyGive += yearBalance[1]; // а в цю виплачених
        }
        yearMoneyBalance = yearBalance[2]; // Значення балансу після всіх 12 місяців

        YearReportTable YRT = new YearReportTable(workbook, allStyles, yearMoneyGet, yearMoneyGive, yearMoneyBalance);
        YRT.yearReportTableCreate(); // Створення сторінки з річним звітом
    }

    private int[] sheetMonthCreate(HSSFWorkbook workbook, String month, int monthDayCount, int indexMonth, int balance) {
        Styles mapstyles = new Styles(workbook);
        HSSFSheet sheet = workbook.createSheet(month);         // Створення сторінки для відповідного місяця
        Map<String, CellStyle> allStyles = mapstyles.styles(); // Усі стилі проекту
        OtherTableElements OTE = new OtherTableElements(sheet, allStyles, balance);
        ReportTable RT = new ReportTable(sheet, allStyles, balance, monthDayCount, indexMonth);

        OTE.otherElements(month);              // Створення інших елементів документу, які не є частинами таблиці звіту
        RT.table(monthDayCount, indexMonth);   // Створення безпосередньо таблиці звіту

        return RT.getTotalSum();
    }
}
