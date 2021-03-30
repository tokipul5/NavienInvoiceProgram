import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import org.apache.poi.ss.usermodel.Row;

import java.io.File;
import java.io.IOException;

import static org.junit.jupiter.api.Assertions.*;

class CreateInvoiceTest {
    private static String[] titleName = {"Billing Doc.", "Sold-to", "Ship-to", "PO NO.", "Payment terms", "Sales Rep" +
            "(Doc)", "Tracking No", "Carrier Name", "Billing Date"};
    private static int[] titleIndex = new int[9];
    private static String[] detailName = {"Billing Qty", "Material", "Material", "Unit Price", "Total amount", "Currency"};
    private static int[] detailIndex = new int[6];

    @org.junit.jupiter.api.Test
    void findIndexOfColumns() throws IOException, BiffException {
        WorkbookSettings workbookSettings = new WorkbookSettings();
        workbookSettings.setEncoding("Cp1252"); //Recognize special character


        Workbook data = null;
        data = Workbook.getWorkbook(new File("5-6-billing-status.xls"), workbookSettings);

        Sheet sheet = data.getSheet(0);
        Cell[] firstRow = sheet.getRow(0);
        for (int index = 0; index < firstRow.length; index++) {
            Cell currentCell = firstRow[index];
            String columnName = currentCell.getContents();
            for (int i = 0; i < titleName.length; i++){
                String title = titleName[i];
                if (title.equals(columnName)) {
                    titleIndex[i] = index;
                    break;
                }
            }
            for (int i = 0; i < detailName.length; i++) {
                String detail = detailName[i];
                if (detail.equals(columnName)) {
                    detailIndex[i] = index;
                    break;
                }
            }
        }
    }

    @org.junit.jupiter.api.Test
    void storeDataInHashMap() {
    }
}