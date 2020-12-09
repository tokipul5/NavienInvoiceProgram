import org.apache.poi.ss.usermodel.Cell;
import jxl.read.biff.BiffException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;

public class TrackInvoice {
    private String pathTrack;
    public void recordInvoice(String pathSave, ArrayList<String> po, ArrayList<String> email, ArrayList<String> pathFile, ArrayList<String> attachmentName) throws IOException, BiffException {
        pathTrack = pathSave + "\\" + "trackMail.xlsx";
        //load Excel
        InputStream input = new FileInputStream(pathTrack);
        XSSFWorkbook workbook = new XSSFWorkbook(input);
        XSSFSheet sheet = workbook.getSheetAt(0);

        for (int i = 0; i < po.size(); i++) {
            Row r = sheet.createRow(sheet.getPhysicalNumberOfRows());
            Cell c0 = r.createCell(0);
            c0.setCellValue(po.get(i));
            Cell c1 = r.createCell(1);
            c1.setCellValue("keeyoukim@gmail.com"); //For testing, change it to email.get(i) later
            Cell c2 = r.createCell(2);
            c2.setCellValue(pathFile.get(i));
            Cell c3 = r.createCell(3);
            c3.setCellValue(attachmentName.get(i));
        }

        FileOutputStream fileOutputStream = new FileOutputStream(pathTrack);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();
    }
}
