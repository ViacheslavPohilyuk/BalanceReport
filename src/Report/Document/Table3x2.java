package Report.Document;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;

import java.util.Map;

/**
 * Created by mac on 23.01.17.
 */
public class Table3x2 {

    private HSSFSheet sheet;
    private Map<String, CellStyle> allStyles;

    private Row row;
    private int startRow;
    private int startColumn;
    private String stylesName;

    public Table3x2(HSSFSheet sheet, Map<String, CellStyle> allStyles, int monthDayCount, int startRow, int startColumn, String stylesName) {
        this.sheet = sheet;
        this.allStyles = allStyles;
        this.startRow = startRow + monthDayCount;
        this.startColumn = startColumn;
        this.stylesName = stylesName;
    }

    public void putCellsValues(String[] notations, int[] values) {
        for (int i = 0; i < 3; i++) {
            row = sheet.createRow(startRow + i);
            for (int j = 0; j < 2; j++)
                if (j == 0) {
                    row.createCell(startColumn + j).setCellValue(notations[i]);
                    row.getCell(startColumn + j).setCellStyle(allStyles.get(getStyleString(stylesName,i,j)));
                } else {
                    row.createCell(startColumn + j).setCellValue(values[i]);
                    row.getCell(startColumn + j).setCellStyle(allStyles.get(getStyleString(stylesName,i,j)));
                }
        }
    }

    private String getStyleString(String stylesName, int i, int j) {
        return stylesName + " ("+ Integer.toString(i) + ";" + Integer.toString(j) + ")";
    }
}
