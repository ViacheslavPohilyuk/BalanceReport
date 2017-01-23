package Report;

import Report.Document.Sheets.SheetsCreate;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;

/**
 * Created by mac on 14.01.17.
 */

public class Main {

    public static void main(String[] args) throws ParseException {
        HSSFWorkbook workbook = new HSSFWorkbook(); // створення excel документу

        SheetsCreate SC = new SheetsCreate(workbook); // створення усіх сторінок

        // Записуємо сторінку у створений ексель-документ
        try (FileOutputStream out = new FileOutputStream(new File("resources/Балансовий звіт.xls"))) {
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Excel файл успішно створений!");
    }
}
