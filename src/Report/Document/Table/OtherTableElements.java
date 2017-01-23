package Report.Document.Table;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import java.util.Map;

/**
 * Created by mac on 14.01.17.
 */
public class OtherTableElements {
    private int balance;
    private HSSFSheet sheet;
    private Map<String, CellStyle> allStyles;

    public OtherTableElements (HSSFSheet sheet, Map<String, CellStyle> allStyles, int balance) {
        this.balance = balance;
        this.sheet = sheet;
        this.allStyles = allStyles;
    }
    public void otherElements(String month) {
        Row row;

        for(int i = 2; i <= 3; i++) {
            row = sheet.createRow(i);
            row.setHeight((short)10);
        }

        CellStyle style = allStyles.get("Середня межа B2");
        row = sheet.createRow(1);
        row.createCell(1).setCellStyle(style);

        //Створення підпису про місяць і рік
        style = allStyles.get("Місяць і рік");
        String MonthYear = month + " 2010 року";
        row.createCell(2).setCellStyle(style);
        row.getCell(2).setCellValue(MonthYear);

        //Створення підпису "Балансовий звіт"
        style = allStyles.get("Балансовий звіт");
        //row = sheet.createRow(1);
        row.createCell(4).setCellValue("Балансовий звіт");
        sheet.setColumnWidth(6, 1380*6);
        row.getCell(4).setCellStyle(style);

        // Ліва комірка підпису "Балансовий звіт"
        style = allStyles.get("Ліва комірка");
        row.createCell(3).setCellStyle(style);

        // Права комірка підпису "Балансовий звіт"
        style = allStyles.get("Права комірка");
        row.createCell(5).setCellStyle(style);

        // Створення підпису "Опис"
        row = sheet.createRow(6);
        style = allStyles.get("Опис");
        row.createCell(1).setCellValue("Опис");
        row.getCell(1).setCellStyle(style);

        // Права верхня межа від "Опис"'у
        style = allStyles.get("Сердня межа C7");
        row.createCell(2).setCellStyle(style);

        // Створення підпису "Початковий"
        style = allStyles.get("Початковий");
        row.createCell(3).setCellValue("Початковий");
        row.getCell(3).setCellStyle(style);

        // Створення підпису "баланс"
        row = sheet.createRow(7);
        style = allStyles.get("баланс");
        row.createCell(3).setCellValue("баланс");
        row.getCell(3).setCellStyle(style);

        // Межі праворуч від "баланс"'у
        style = allStyles.get("Сердня межа C7");
        row.createCell(4).setCellStyle(style);
        row.createCell(5).setCellStyle(style);

        // Встановлюємо розмір балансу
        row.getCell(4).setCellValue(balance);
    }
}
