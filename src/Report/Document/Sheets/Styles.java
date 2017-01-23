package Report.Document.Sheets;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.util.HashMap;
import java.util.Map;

/**
 * Created by mac on 14.01.17.
 */
public class Styles {
    private HSSFWorkbook workbook;

    Styles (HSSFWorkbook workbook) {
        this.workbook = workbook;
    }
    public Map<String, CellStyle> styles () {
        Map<String, CellStyle> styles = new HashMap<>();
        CellStyle style = workbook.createCellStyle();

        // Перший рядок таблиці звіту
        Font fontFirstTable = workbook.createFont();
        fontFirstTable.setFontHeightInPoints((short)11);
        fontFirstTable.setBoldweight(Font.BOLDWEIGHT_BOLD);
        fontFirstTable.setFontName("Times New Roman");
        style.setFont(fontFirstTable);

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderBottom(CellStyle.BORDER_DOTTED);
        style.setBorderLeft(CellStyle.BORDER_DOTTED);
        style.setBorderRight(CellStyle.BORDER_DOTTED);
        style.setBorderTop(CellStyle.BORDER_DOTTED);
        styles.put("Перший рядок таблиці", style);

        // Напис "Балансовий звіт"
        style = workbook.createCellStyle();
        Font fontBalance = workbook.createFont();
        fontBalance.setFontHeightInPoints((short)22);
        fontBalance.setBoldweight(Font.BOLDWEIGHT_BOLD);
        fontBalance.setItalic(true);
        fontBalance.setFontName("Times New Roman");
        style.setFont(fontBalance);

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderLeft(CellStyle. BORDER_MEDIUM);
        style.setBorderTop(CellStyle. BORDER_MEDIUM);
        styles.put("Балансовий звіт", style);

        // Напис місяця і року
        style = workbook.createCellStyle();
        Font fontMonth = workbook.createFont();
        fontMonth.setFontHeightInPoints((short)16);
        fontMonth.setBoldweight(Font.BOLDWEIGHT_BOLD);
        fontMonth.setFontName("Times New Roman");
        style.setFont(fontMonth);

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderBottom(CellStyle. BORDER_MEDIUM);
        styles.put("Місяць і рік", style);

        // Середня межа B2
        style = workbook.createCellStyle();
        style.setBorderBottom(CellStyle. BORDER_MEDIUM);
        styles.put("Середня межа B2", style);

        // Опис
        style = workbook.createCellStyle();
        Font fontDescript = workbook.createFont();
        fontDescript.setFontHeightInPoints((short)11);
        fontDescript.setBoldweight(Font.BOLDWEIGHT_BOLD);
        fontDescript.setFontName("Times New Roman");

        style.setFont(fontDescript);
        style.setBorderBottom(CellStyle. BORDER_MEDIUM);
        style.setBorderRight(CellStyle. BORDER_MEDIUM);
        styles.put("Опис", style);

        // Ліва комірка підпису "Балансовий звіт"
        style = workbook.createCellStyle();
        style.setBorderLeft(CellStyle. BORDER_MEDIUM);
        style.setBorderTop(CellStyle. BORDER_MEDIUM);
        styles.put("Ліва комірка", style);

        // Права комірка підпису "Балансовий звіт"
        style = workbook.createCellStyle();
        style.setBorderTop(CellStyle. BORDER_MEDIUM);
        styles.put("Права комірка", style);

        // Сердня межа C7
        style = workbook.createCellStyle();
        Font fontLeftBal = workbook.createFont();
        fontLeftBal.setFontHeightInPoints((short)11);
        fontLeftBal.setBoldweight(Font.BOLDWEIGHT_BOLD);
        fontLeftBal.setFontName("Times New Roman");
        style.setFont(fontLeftBal);

        style.setBorderTop(CellStyle. BORDER_MEDIUM);
        styles.put("Сердня межа C7", style);

        // Створення підпису "Початковий"
        style = workbook.createCellStyle();
        Font fontBegin = workbook.createFont();
        fontBegin.setFontHeightInPoints((short)11);
        fontBegin.setBoldweight(Font.BOLDWEIGHT_BOLD);
        fontBegin.setFontName("Times New Roman");
        style.setFont(fontBegin);
        styles.put("Початковий", style);

        // Створення підпису "баланс"
        style = workbook.createCellStyle();
        Font fontBal = workbook.createFont();
        fontBal.setFontHeightInPoints((short)11);
        fontBal.setBoldweight(Font.BOLDWEIGHT_BOLD);
        fontBal.setFontName("Times New Roman");
        style.setFont(fontBal);

        style.setBorderBottom(CellStyle. BORDER_MEDIUM);
        style.setBorderRight(CellStyle. BORDER_MEDIUM);
        styles.put("баланс", style);

        // Стиль комірок першого стовбця таблиці звіту
        style = workbook.createCellStyle();
        Font fontDate = workbook.createFont();
        fontDate.setFontHeightInPoints((short) 11);
        fontDate.setFontName("Times New Roman");
        style.setFont(fontDate);

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderLeft(CellStyle.BORDER_DOTTED);

        // Створення формату дати
        DataFormat format = workbook.createDataFormat();
        style.setDataFormat(format.getFormat("dd.mm.yyyy"));
        styles.put("Перший стовбець таблиці", style);

        // Комірки таблиці звіту
        style = workbook.createCellStyle();
        Font fontRowsTable = workbook.createFont();
        fontRowsTable.setFontHeightInPoints((short)11);
        fontRowsTable.setFontName("Times New Roman");
        style.setFont(fontRowsTable);

        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderLeft(CellStyle.BORDER_DOTTED);
        style.setBorderRight(CellStyle.BORDER_DOTTED);
        styles.put("Межі таблиці звіту", style);

        // Стилі для комірок загального балансу (class TotalBalance)
        CellStyle styleTotalBalance = workbook.createCellStyle();
        Font fontTotalBalance = workbook.createFont();
        fontTotalBalance.setFontHeightInPoints((short)11);
        fontTotalBalance.setBoldweight(Font.BOLDWEIGHT_BOLD);
        fontTotalBalance.setFontName("Times New Roman");
        styleTotalBalance.setFont(fontTotalBalance);

        style = workbook.createCellStyle();
        style.cloneStyleFrom(styleTotalBalance);
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setBorderTop(CellStyle.BORDER_DOTTED);
        style.setBorderLeft(CellStyle.BORDER_DOTTED);
        styles.put("Загальний баланс (0;0)", style);

        style = workbook.createCellStyle();
        style.cloneStyleFrom(styleTotalBalance);
        style.setBorderTop(CellStyle.BORDER_DOTTED);
        style.setBorderRight(CellStyle.BORDER_DOTTED);
        style.setAlignment(HorizontalAlignment.CENTER);
        styles.put("Загальний баланс (0;1)", style);

        style = workbook.createCellStyle();
        style.cloneStyleFrom(styleTotalBalance);
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setBorderLeft(CellStyle.BORDER_DOTTED);
        styles.put("Загальний баланс (1;0)", style);

        style = workbook.createCellStyle();
        style.cloneStyleFrom(styleTotalBalance);
        style.setBorderRight(CellStyle.BORDER_DOTTED);
        style.setAlignment(HorizontalAlignment.CENTER);
        styles.put("Загальний баланс (1;1)", style);

        style = workbook.createCellStyle();
        style.cloneStyleFrom(styleTotalBalance);
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setBorderBottom(CellStyle.BORDER_DOTTED);
        style.setBorderLeft(CellStyle.BORDER_DOTTED);
        styles.put("Загальний баланс (2;0)", style);

        style = workbook.createCellStyle();
        style.cloneStyleFrom(styleTotalBalance);
        style.setBorderBottom(CellStyle.BORDER_DOTTED);
        style.setBorderRight(CellStyle.BORDER_DOTTED);
        style.setAlignment(HorizontalAlignment.CENTER);
        styles.put("Загальний баланс (2;1)", style);

        // Стилі для комірок річного звіту (class YearReportTable)
        Font fontYearReportV1 = workbook.createFont();
        fontYearReportV1.setFontHeightInPoints((short)36);
        fontYearReportV1.setBoldweight(Font.BOLDWEIGHT_BOLD);
        fontYearReportV1.setItalic(true);
        fontYearReportV1.setFontName("Times New Roman");

        Font fontYearReportV2 = workbook.createFont();
        fontYearReportV2.setFontHeightInPoints((short)22);
        fontYearReportV2.setFontName("Times New Roman");

        style = workbook.createCellStyle();
        style.setFont(fontYearReportV1);
        styles.put("Річний звіт", style);

        style = workbook.createCellStyle();
        Font font = workbook.createFont();
        style.setFont(fontYearReportV2);
        style.setBorderTop(CellStyle.BORDER_THICK);
        style.setBorderLeft(CellStyle.BORDER_THICK);
        styles.put("Річний звіт (0;0)", style);

        style = workbook.createCellStyle();
        style.setFont(fontYearReportV2);
        style.setBorderTop(CellStyle.BORDER_THICK);
        style.setBorderRight(CellStyle.BORDER_THICK);
        styles.put("Річний звіт (0;1)", style);

        style = workbook.createCellStyle();
        style.setFont(fontYearReportV2);
        style.setBorderLeft(CellStyle.BORDER_THICK);
        styles.put("Річний звіт (1;0)", style);

        style = workbook.createCellStyle();
        style.setFont(fontYearReportV2);
        style.setBorderRight(CellStyle.BORDER_THICK);
        styles.put("Річний звіт (1;1)", style);

        style = workbook.createCellStyle();
        style.setFont(fontYearReportV2);
        style.setBorderBottom(CellStyle.BORDER_THICK);
        style.setBorderLeft(CellStyle.BORDER_THICK);
        styles.put("Річний звіт (2;0)", style);

        style = workbook.createCellStyle();
        style.setFont(fontYearReportV2);
        style.setBorderBottom(CellStyle.BORDER_THICK);
        style.setBorderRight(CellStyle.BORDER_THICK);
        styles.put("Річний звіт (2;1)", style);

        return styles;
    }
}
