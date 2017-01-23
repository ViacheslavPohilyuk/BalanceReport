package Report.Document.Table;

import Report.Document.Table3x2;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.*;

import java.util.Map;

/**
 * Created by mac on 20.01.17.
 */
public class TotalBalance {
    private int balance;
    private int monthDayCount;
    private HSSFSheet sheet;
    private Map<String, CellStyle> allStyles;

    private int startRow = (monthDayCount + 12) + 2;
    private int startColumn = 4;
    private final String stylesName = "Загальний баланс";

    TotalBalance(HSSFSheet sheet, Map<String, CellStyle> allStyles, int balance, int monthDayCount) {
        this.balance = balance;
        this.monthDayCount = monthDayCount;
        this.sheet = sheet;
        this.allStyles = allStyles;
    }

    public void totalBalance(int totalSum[]) {
        Table3x2 T = new Table3x2(sheet, allStyles, monthDayCount, startRow, startColumn, stylesName);
        int valuse[] = new int[3];

        String notations[] = {"Отримано", "Виплачено", "Поточний баланс"}; 

        valuse[0] = totalSum[0]; // Отримано
        valuse[1] = totalSum[1]; // Виплачено
        valuse[2] = (balance + totalSum[0] - totalSum[1]); // Поточний баланс

        // Встановлення значень і стилів комірок, в яких буде розміщуватися загальний баланс
        T.putCellsValues(notations,valuse);

    }
}
