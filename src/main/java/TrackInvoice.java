import jxl.Sheet;
import jxl.Workbook;
import org.apache.poi.ss.usermodel.Cell;
import jxl.read.biff.BiffException;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.*;
import java.util.ArrayList;

public class TrackInvoice {
    private String pathTrack;
    private String folderName = "Track Mails";

    public void createMailFile(String monthYear) throws IOException {
        String pathFile = pathTrack + "\\" + folderName + "\\" + monthYear + ".xlsx";
        File f = new File(pathFile);
        if (f.exists()) {
            return;
        }
        XSSFWorkbook workbook = new XSSFWorkbook();
        /* CreationHelper helps us create instances of various things like DataFormat,
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        XSSFSheet sheet = workbook.createSheet();
        Row headerRow = sheet.createRow(0);

        Cell c0 = headerRow.createCell(0);
        c0.setCellValue("PO");
        Cell c1 = headerRow.createCell(1);
        c1.setCellValue("EMAIL"); //For testing, change it to email.get(i) later
        Cell c2 = headerRow.createCell(2);
        c2.setCellValue("PATH FILE");
        Cell c3 = headerRow.createCell(3);
        c3.setCellValue("FILE NAME");
        Cell c4 = headerRow.createCell(4);
        c4.setCellValue("SENT");
        Cell c5 = headerRow.createCell(5);
        c5.setCellValue("NAME OF CUSTOMER");
        Cell c6 = headerRow.createCell(6);
        c6.setCellValue("INVOICE NUMBER");
        Cell c7 = headerRow.createCell(7);
        c7.setCellValue("NAME");
        Cell c8 = headerRow.createCell(8);
        c8.setCellValue("MONTHYEAR");

        FileOutputStream fileOutputStream = new FileOutputStream(pathFile);
        workbook.write(fileOutputStream);
        fileOutputStream.close();
        workbook.close();
    }

    public void recordInvoice(String pathSave, ArrayList<String> po, ArrayList<String> email,
                              ArrayList<String> pathFile, ArrayList<String> attachmentName,
                              ArrayList<String> customerList, ArrayList<String> invoiceList,
                              ArrayList<String> nameList,
                              ArrayList<String> dateList) throws IOException,
            BiffException {
        pathTrack = pathSave;

        File f = new File(pathTrack + "\\" + folderName);
        if (!f.exists()) {
            f.mkdirs();
        }
        String pathAllMail = pathSave + "\\" + "trackMail.xlsx";
        //load Excel
        InputStream inputAllMail = new FileInputStream(pathAllMail);
        XSSFWorkbook workbookAllMail = new XSSFWorkbook(inputAllMail);
        XSSFSheet sheetAllMail = workbookAllMail.getSheetAt(0);

        for (int i = 0; i < po.size(); i++) {
            String monthYear = dateList.get(i);
            createMailFile(monthYear);
            String locationFile = pathTrack + "\\" + folderName + "\\" + monthYear + ".xlsx";
            //load Excel according to monthYear
            InputStream input = new FileInputStream(locationFile);
            XSSFWorkbook workbook = new XSSFWorkbook(input);
            XSSFSheet sheet = workbook.getSheetAt(0);

            Row r = sheet.createRow(sheet.getPhysicalNumberOfRows());
            Cell c0 = r.createCell(0);
            c0.setCellValue(po.get(i));
            Cell c1 = r.createCell(1);
            c1.setCellValue("keeyoukim@gmail.com"); //For testing, change it to email.get(i) later
            Cell c2 = r.createCell(2);
            c2.setCellValue(pathFile.get(i));
            Cell c3 = r.createCell(3);
            c3.setCellValue(attachmentName.get(i));
            Cell c5 = r.createCell(5);
            c5.setCellValue(customerList.get(i));
            Cell c6 = r.createCell(6);
            c6.setCellValue(invoiceList.get(i));
            Cell c7 = r.createCell(7);
            c7.setCellValue(nameList.get(i));
            Cell c8 = r.createCell(8);
            c8.setCellValue(dateList.get(i));

            //Load Excel to all mails
            Row rowAllMail = sheetAllMail.createRow(sheetAllMail.getPhysicalNumberOfRows());
            Cell cAll0 = rowAllMail.createCell(0);
            cAll0.setCellValue(po.get(i));
            Cell cAll1 = rowAllMail.createCell(1);
            cAll1.setCellValue("keeyoukim@gmail.com"); //For testing, change it to email.get(i) later
            Cell cAll2 = rowAllMail.createCell(2);
            cAll2.setCellValue(pathFile.get(i));
            Cell cAll3 = rowAllMail.createCell(3);
            cAll3.setCellValue(attachmentName.get(i));
            Cell cAll5 = rowAllMail.createCell(5);
            cAll5.setCellValue(customerList.get(i));
            Cell cAll6 = rowAllMail.createCell(6);
            cAll6.setCellValue(invoiceList.get(i));
            Cell cAll7 = rowAllMail.createCell(7);
            cAll7.setCellValue(nameList.get(i));
            Cell cAll8 = rowAllMail.createCell(8);
            cAll8.setCellValue(dateList.get(i));

            FileOutputStream fileOutputStream = new FileOutputStream(locationFile);
            workbook.write(fileOutputStream);
            fileOutputStream.close();
            workbook.close();
        }
        FileOutputStream fileOutputStream = new FileOutputStream(pathAllMail);
        workbookAllMail.write(fileOutputStream);
        fileOutputStream.close();
        workbookAllMail.close();
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
