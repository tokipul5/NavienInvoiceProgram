import org.apache.poi.ss.usermodel.Cell;
import jxl.read.biff.BiffException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
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
    public void checkStatus(String pathSave, JTextArea consoleOutput) throws IOException {
        String pathCountInvoices = pathSave + "\\" + "countInvoices.txt";
        String pathCountSentEmails = pathSave + "\\" + "countSentEmails.txt";
        pathTrack = pathSave + "\\" + "trackMail.xlsx";
        FileReader reader = null;
        int countInvoices = 0;
        int countEmails = 0;

        try {
            reader = new FileReader(pathCountInvoices);
        } catch (Exception e){
            consoleOutput.append("Please choose the correct directory.\n");
        }

        BufferedReader brInvoices = new BufferedReader(reader);
        try {
            String line = brInvoices.readLine();
            countInvoices = Integer.valueOf(line);
        } finally {
            brInvoices.close();
        }

        BufferedReader brEmails = new BufferedReader(new FileReader(pathCountSentEmails));
        try {
            String line = brEmails.readLine();
            countEmails = Integer.valueOf(line);
        } finally {
            brEmails.close();
        }

        int missingEmails = countInvoices - countEmails;

        if (missingEmails > 0) {
            //load Excel
            InputStream input = new FileInputStream(pathTrack);
            XSSFWorkbook workbook = new XSSFWorkbook(input);
            XSSFSheet sheet = workbook.getSheetAt(0);

            for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
                Row r = sheet.getRow(i);
                String filePath = r.getCell(2).getStringCellValue();
                if (r.getCell(1) == null) {
                    consoleOutput.append(filePath + " has no email and not yet sent.\n");
                } else if (r.getCell(4) == null) {
                    consoleOutput.append(filePath + " has not yet sent.\n");
                }
            }
            workbook.close();
        }
        consoleOutput.append("There are " + countInvoices + " invoices created in total.\n");
        consoleOutput.append("There are " + countEmails + " invoices sent in total.\n");
        consoleOutput.append("Missing " + missingEmails + " emails.\n");
        consoleOutput.append("If you want to send emails, click the \"Send emails to buyers\" button or check the existence of their emails.\n");
    }
}
