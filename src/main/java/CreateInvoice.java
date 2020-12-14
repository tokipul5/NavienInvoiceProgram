import java.io.*;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.*;

import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import com.documents4j.api.DocumentType;
import com.sun.prism.impl.ps.CachingEllipseRep;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
//import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;

import javax.swing.*;

public class CreateInvoice {
    private String pathData;
    private String pathSave;
    private String pathTemplate = "Invoice-Template.docx";
    private int count = 0;
    private static HashMap<String, ArrayList<Integer>> poAndRows;
    private static HashMap<String, String> poAndDate;
    private static String[] colName = {"C", "E", "G", "H", "M", "Amount", "Q"};
    private static int[] colIndex = {4, 0, 16, 2, 6, 7, 12};
    //4: Qty, 0: Order type Desc. (But it should be ITEM NO which does not exist in excel file),
    //16: C/M Desc (old item no), 2: Material Desc, 6: order price, 7: order amount, 12: cur
    private OutlookEmail email;

    public int getCount() {
        return count;
    }

    public void updateCount() {
        File fileToBeModified = new File(pathSave + "\\" + "countInvoices.txt");
        String oldContent = "";
        BufferedReader reader = null;
        FileWriter writer = null;
        int newCount = 0;
        try
        {
            reader = new BufferedReader(new FileReader(fileToBeModified));
            //Reading all the lines of input text file into oldContent
            oldContent = reader.readLine();
            System.out.println(oldContent);
            int oldCount = Integer.valueOf(oldContent);
            newCount = oldCount + count;
            //Replacing oldString with newString in the oldContent
            String newContent = Integer.toString(newCount);
            //Rewriting the input text file with newContent
            writer = new FileWriter(fileToBeModified);
            writer.write(newContent);
            count = 0;
        }
        catch (IOException e) {
            e.printStackTrace();
        } finally {
            try
            {
                //Closing the resources
                reader.close();
                writer.close();
            }
            catch (IOException e)
            {
                e.printStackTrace();
            }
        }
    }

    public CreateInvoice(String pathData, String pathSave, String id, String pw) {
        this.pathData = pathData;
        this.pathSave = pathSave;
        poAndRows = new HashMap<>();
        poAndDate = new HashMap<>();
        email = new OutlookEmail(id,pw, null);
    }

    public static void createDateFolder(String dateInvoice) {
        File f = new File(dateInvoice);
        if (!f.exists()) {
            f.mkdirs();
        }
    }

    public void generateInvoice(JTextArea textArea) throws IOException, BiffException, XmlException {
        //Store po number, email, file path, file name in ArrayList
        ArrayList<String> poList = new ArrayList<>();
        ArrayList<String> emailList = new ArrayList<>();
        ArrayList<String> pathFileList = new ArrayList<>();
        ArrayList<String> fileNameList = new ArrayList<>();

        WorkbookSettings workbookSettings = new WorkbookSettings();
        workbookSettings.setEncoding("Cp1252"); //Recognize special character

        Workbook data = null;
        try {
            data = Workbook.getWorkbook(new File(this.pathData), workbookSettings);
        } catch (Exception e) {
            textArea.append("Please select the data file.\n");
        }
        storeDataInHashMap(data);

        for (Map.Entry<String, ArrayList<Integer>> entry : poAndRows.entrySet()) {
            if (entry.getKey().equals("") || entry.getKey().equals("Order PO Detail"))
                continue;
            String po = entry.getKey();
            Sheet sheet = data.getSheet(0);
            String date = poAndDate.get(po);
            if (date.equals("")) {
                date = "No date";
            }
            createDateFolder(pathSave + "\\" + date);
            String pathNewDoc = pathSave + "\\" + date + "\\" + po + ".docx";
            File checkFileExist = new File(pathNewDoc);
            if (checkFileExist.exists()) {
                continue;
            }
            CopyWord copyWord = new CopyWord(pathTemplate, pathNewDoc);
            copyWord.copy(); //copy template to new word file

            //Edit newDoc file
            XWPFDocument newDoc = new XWPFDocument(new FileInputStream(pathNewDoc));
            XWPFTable table = newDoc.getTableArray(3);

            //returns  array of row indices
            ArrayList<Integer> intWithSamePo = entry.getValue();

            //Take the first row to find information
            Cell[] firstRow = sheet.getRow(intWithSamePo.get(0) - 1);
            String invoiceNum = "";
            String companyName = firstRow[1].getContents();
            String companyAddress = "";
            String contactPerson = firstRow[18].getContents();

            String poNumber = firstRow[8].getContents();
            String paymentTerm = firstRow[13].getContents();
            String salesRep = firstRow[11].getContents();
            String trackingNumber = "";
            String shipVia = firstRow[15].getContents();
            String shippingDate = "";
            String dueDate = "";

            String cur = firstRow[12].getContents();

            String email = "";
            String pathPDF = pathSave + "\\" + date + "\\" + po + ".pdf";

            //Console output
            textArea.append("Created " + pathPDF + "\n");
            poList.add(poNumber);
            emailList.add(email);
            pathFileList.add(pathPDF);
            fileNameList.add(po + ".pdf");


            int countTable = 0; //Find 4 in order to add detail in table.
            double totalAmount = 0;
            for (XWPFTable tableDoc : newDoc.getTables()) {
                if (countTable == 4) {
                    XWPFTableRow oldRow = tableDoc.getRow(1); //The first empty row to copy
                    for (int i = 0; i < intWithSamePo.size(); i++) {
                        int rowNum = intWithSamePo.get(i);
                        Cell[] rowData = sheet.getRow(rowNum - 1);
                        String[] items = new String[7];
                        for (int j = 0; j < colIndex.length; j++) {
                            int index = colIndex[j];
                            if (index == 6 || index == 7) {
                                items[j] = rowData[index].getContents().replaceAll(",", "");
                                if (index == 7) {
                                    totalAmount += Double.parseDouble(items[j]);
                                }
                            }  else {
                                items[j] = rowData[index].getContents();
                            }
                        }
                        CTRow row = CTRow.Factory.parse(oldRow.getCtRow().newInputStream());
                        XWPFTableRow newRow = new XWPFTableRow(row, tableDoc);
                        int count = 0;
                        for (XWPFTableCell cell : newRow.getTableCells()) {
                            cell.setText(items[count]);
                            count++;
                        }
                        tableDoc.addRow(newRow, i + 2);
                    }
                    tableDoc.removeRow(1);
                }
                for (XWPFTableRow rowDoc : tableDoc.getRows()) {
                    for (XWPFTableCell cellDoc : rowDoc.getTableCells()) {
                        for (XWPFParagraph paraDoc : cellDoc.getParagraphs()) {
                            for (XWPFRun run : paraDoc.getRuns()) {
                                String str = run.getText(0);
                                if (str != null && str.equals("InvoiceNum")) {
                                    run.setText("invoice number", 0);
                                } else if (str != null && str.equals("Name")) {
                                    run.setText(companyName, 0);
                                }  else if (str != null && str.equals("ContactPerson")) {
                                    run.setText(contactPerson, 0);
                                }else if (str != null && str.equals("PoNumber")) {
                                    run.setText(poNumber, 0);
                                } else if (str != null && str.equals("PaymentTerm")) {
                                    run.setText(paymentTerm, 0);
                                } else if (str != null && str.equals("SalesRep")) {
                                    run.setText(salesRep, 0);
                                } else if (str != null && str.equals("TrackingNumber")) {
                                    run.setText(trackingNumber, 0);
                                } else if (str != null && str.equals("ShipVia")) {
                                    run.setText(shipVia, 0);
                                } else if (str != null && str.equals("ShippingDate")) {
                                    run.setText(shippingDate, 0);
                                } else if (str != null && str.equals("DueDate")) {
                                    run.setText(dueDate, 0);
                                } else if (str != null && str.equals("Cur")) {
                                    run.setText(cur, 0);
                                } else if (str != null && str.equals("Sum")) {
                                    String total = String.format("%.2f", totalAmount);
                                    run.setText(total, 0);
                                }
                            }
                        }
                    }
                }
                countTable++;
            }
            File output = new File(pathNewDoc);
            newDoc.write(new FileOutputStream(output));
            //Convert word to pdf
            File inputWord = new File(pathNewDoc);
            File outputFile = new File(pathPDF);
            count++;
            try  {
                InputStream docxInputStream = new FileInputStream(inputWord);
                OutputStream outputStream = new FileOutputStream(outputFile);
                IConverter converter = LocalConverter.builder().build();
                converter.convert(docxInputStream).as(DocumentType.DOCX).to(outputStream).as(DocumentType.PDF).execute();
                outputStream.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        TrackInvoice trackInvoice = new TrackInvoice();
        trackInvoice.recordInvoice(pathSave, poList, emailList, pathFileList, fileNameList);

        //System.exit(1);
    }

    /*
    Iterate each row from data excel file and store index of rows according to po numbers
    into hashmap<String, ArrayList<Integer>>
    Key: po number, Value: String of row numbers
    indexOfRows: integer in ArrayList

    Put PO number and date of PO into hashmap<String, String>
    Key: PO number, Value: PO date
    */
    public static void storeDataInHashMap(Workbook wb) {
        Sheet sheet = wb.getSheet(0);
        for (int row = 1; row < sheet.getRows(); row++) {
            String po = sheet.getCell("I" + row).getContents(); //Order PO Detail on column I
            if (poAndRows.containsKey(po)) {
                ArrayList<Integer> temp = poAndRows.get(po);
                temp.add(row);
            } else {
                ArrayList<Integer> temp = new ArrayList();
                temp.add(row);
                poAndRows.put(po, temp);
            }
            String date = sheet.getCell("J" + row).getContents(); //PO Date on column J
            if (!poAndDate.containsKey(po)) {
                poAndDate.put(po, date);
            }
        }
    }
}
