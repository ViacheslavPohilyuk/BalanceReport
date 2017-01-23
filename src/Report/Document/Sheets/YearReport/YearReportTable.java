package Report.Document.Sheets.YearReport;

import Report.Document.Table3x2;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.util.Map;

/**
 * Created by mac on 20.01.17.
 */
public class YearReportTable {
    private HSSFWorkbook workbook;
    Map<String, CellStyle> allStyles;
    private int yearMoneyGet;
    private int yearMoneyGive;
    private int yearMoneyBalance;

    Row row;
    private int startRow = 2;
    private int startColumn = 1;
    private final String stylesName = "Річний звіт";
    private String notations[] = {"Отримано", "Виплачено", "Поточний баланс"};

    public YearReportTable(HSSFWorkbook workbook, Map<String, CellStyle> allStyles, int yearMoneyGet, int yearMoneyGive, int yearMoneyBalance) {
        this.workbook = workbook;
        this.allStyles = allStyles;
        this.yearMoneyGet = yearMoneyGet;
        this.yearMoneyGive = yearMoneyGive;
        this.yearMoneyBalance = yearMoneyBalance;
    }

    public void yearReportTableCreate() {
        HSSFSheet sheet = workbook.createSheet("Річний звіт");
        Table3x2 T = new Table3x2(sheet, allStyles, 0, startRow, startColumn, stylesName);
        int values[] = new int[3];

        values[0] = yearMoneyGet;     // Отримано
        values[1] = yearMoneyGive;    // Виплачено
        values[2] = yearMoneyBalance; // Поточний баланс

        row = sheet.createRow(1);
        row.createCell(1).setCellStyle(allStyles.get("Річний звіт"));
        row.getCell(1).setCellValue("Річний звіт");

        // Встановлення значень комірок для річного звіту
        T.putCellsValues(notations, values);

        // Встановлення ширини стовбців
        for(int i = 1; i < 3; i++)
            sheet.setColumnWidth(i, 1280 * 6);
    }
}
