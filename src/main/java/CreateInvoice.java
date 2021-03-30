import java.io.*;
import java.nio.file.Paths;
import java.text.ParseException;
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
    private String name;
    private int count = 0;
    private static HashMap<String, ArrayList<Integer>> poAndRows;
    private static HashMap<String, String> poAndDate;
    private static HashMap<String, String> poAndCustomer;
    private static String[] titleName = {"Billing Doc.", "Sold-to", "Ship-to", "PO NO.", "Payment terms", "Sales Rep" +
            "(Doc)", "Tracking No", "Carrier Name", "Billing Date", "Per Name"}; // 0-9
    private static int[] titleIndex;
    private static String[] detailName = {"Billing Qty", "Material", "Tax code", "Material", "Unit Price", "Total " +
            "amount", "Currency"};
    private static int[] detailIndex;
    private static HashMap<String, ArrayList<String>> customerAndPo;

    /*
    Constructor
     */
    public CreateInvoice(String pathData, String pathSave, String name) {
        this.pathData = pathData;
        this.pathSave = pathSave;
        this.name = name;
        poAndRows = new HashMap<>();
        poAndDate = new HashMap<>();
        poAndCustomer = new HashMap<>();
        titleIndex = new int[10];
        detailIndex = new int[7];
    }

    public HashMap<String, ArrayList<String>> getCustomerAndPo() {
        return customerAndPo;
    }

    public int getCount() {
        return count;
    }

    /*
    Fill in the titleIndex and detailIndex arrays to know the indices of each column to use.
     */
    public static void findIndexOfColumns(Workbook wb) {
        Sheet sheet = wb.getSheet(0);
        Cell[] firstRow = sheet.getRow(0);
        boolean paymentTerm = true;
        boolean salesRep = true;
        boolean soldTo = true;
        boolean shipTo = true;
        boolean material = true;
        for (int index = 0; index < firstRow.length; index++) {
            Cell currentCell = firstRow[index];
            String columnName = currentCell.getContents();
            for (int i = 0; i < titleName.length; i++) {
                String title = titleName[i];
                if (title.equals(columnName)) {
                    if (title.equals("Payment terms")) {
                        if (paymentTerm) {
                            paymentTerm = false;
                            titleIndex[i] = index;
                            break;
                        }
                    } else if (title.equals("Sales Rep(Doc)")) {
                        if (salesRep) {
                            salesRep = false;
                            titleIndex[i] = index;
                            break;
                        }
                    } else if (title.equals("Ship-to")) {
                        if (shipTo) {
                            shipTo = false;
                            titleIndex[i] = index;
                            break;
                        }
                    } else if (title.equals("Sold-to")) {
                        if (soldTo) {
                            soldTo = false;
                            titleIndex[i] = index;
                            break;
                        }
                    } else {
                        titleIndex[i] = index;
                        break;
                    }
                }
            }
            for (int i = 0; i < detailName.length; i++) {
                String detail = detailName[i];
                if (detail.equals(columnName)) {
                    if (detail.equals("Material") && material) {
                        material = false;
                        detailIndex[i] = index;
                        break;
                    } else if (detail.equals("Material")) {
                        detailIndex[i+2] = index;
                        break;
                    }
                    detailIndex[i] = index;
                    break;
                }
            }
        }
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
            String po = sheet.getRow(row)[titleIndex[3]].getContents(); //Order PO Detail is the index of 3 in the
            // titleName
            // array
            if (poAndRows.containsKey(po)) {
                ArrayList<Integer> temp = poAndRows.get(po);
                temp.add(row);
            } else {
                ArrayList<Integer> temp = new ArrayList();
                temp.add(row);
                poAndRows.put(po, temp);
            }
            String date = sheet.getRow(row)[titleIndex[8]].getContents(); //Billing Date is the index of 8 in the
            // titleName
            date = date.replaceAll("/", ".");
            if (!poAndDate.containsKey(po)) {
                poAndDate.put(po, date);
            }

            String customer = sheet.getRow(row)[titleIndex[1]].getContents(); //Sold-to is the index of 1 in the
            // titleName
            if (!poAndCustomer.containsKey(po)) {
                poAndCustomer.put(po, customer);
            }
        }
    }

    /*
    Update count of txt files
     */
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

    public static void createDateFolder(String dateInvoice) {
        File f = new File(dateInvoice);
        if (!f.exists()) {
            f.mkdirs();
        }
    }

    public void generateInvoice(JTextArea textArea) throws IOException, BiffException, XmlException, ParseException {
        //Store po number, email, file path, file name in ArrayList
        ArrayList<String> poList = new ArrayList<>();
        ArrayList<String> emailList = new ArrayList<>();
        ArrayList<String> pathFileList = new ArrayList<>();
        ArrayList<String> fileNameList = new ArrayList<>();
        ArrayList<String> customerList = new ArrayList<>();
        ArrayList<String> invoiceList = new ArrayList<>();
        ArrayList<String> nameList = new ArrayList<>();
        ArrayList<String> dateList = new ArrayList<>();

        WorkbookSettings workbookSettings = new WorkbookSettings();
        workbookSettings.setEncoding("Cp1252"); //Recognize special character

        Workbook data = null;
        try {
            data = Workbook.getWorkbook(new File(this.pathData), workbookSettings);
        } catch (Exception e) {
            textArea.append("Please select the data file.\n");
        }
        findIndexOfColumns(data);
        storeDataInHashMap(data);

        for (Map.Entry<String, ArrayList<Integer>> entry : poAndRows.entrySet()) {
            if (entry.getKey().equals("") || entry.getKey().equals("Order PO Detail"))
                continue;
            String po = entry.getKey();
            Sheet sheet = data.getSheet(0);
//            String date = poAndDate.get(po);
//            if (date.equals("")) {
//                date = "No date";
//            }

            //returns  array of row indices
            ArrayList<Integer> intWithSamePo = entry.getValue();



            //Take the first row to find information
            Cell[] firstRow = sheet.getRow(intWithSamePo.get(0));

            //Check name if not equal pass
            String employeeName = firstRow[titleIndex[9]].getContents();
            if (!employeeName.equals(name)) {
                continue;
            }

            String invoiceNum = firstRow[titleIndex[0]].getContents();
            String companyName = firstRow[titleIndex[1]].getContents();
            String companyAddress = "Not Found";
            String contactPerson = firstRow[titleIndex[2]].getContents();

            String poNumber = firstRow[titleIndex[3]].getContents();
            String paymentTerm = firstRow[titleIndex[4]].getContents();
            String salesRep = firstRow[titleIndex[5]].getContents();
            String trackingNumber = firstRow[titleIndex[6]].getContents();
            String shipVia = firstRow[titleIndex[7]].getContents();
            String shippingDate = firstRow[titleIndex[8]].getContents().replace("/", "_");

            //Name the word file
            String customer = poAndCustomer.get(po);
            createDateFolder(pathSave + "\\" + customer);
            String fileName = invoiceNum + "-" + poNumber + "-" + shippingDate;
            String pathNewDoc = pathSave + "\\" + customer + "\\" + fileName +
                    ".docx";
            File checkFileExist = new File(pathNewDoc);
            if (checkFileExist.exists()) {
                continue;
            }
            CopyWord copyWord = new CopyWord(pathTemplate, pathNewDoc);
            copyWord.copy(); //copy template to new word file

            //Edit newDoc file
            XWPFDocument newDoc = new XWPFDocument(new FileInputStream(pathNewDoc));
            XWPFTable table = newDoc.getTableArray(3);

            String[] termArr = paymentTerm.split(" ");
            int lenArr = termArr.length;
//            System.out.println(lenArr);
//            System.out.println(termArr[lenArr-2]);
            int days = Integer.parseInt(termArr[lenArr-2]);
            Date due = new SimpleDateFormat("MM_dd_yyyy").parse(shippingDate);
            SimpleDateFormat format = new SimpleDateFormat("MM_dd_yyyy");
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(due);
            calendar.add(Calendar.DATE, days);
            String dueDate = format.format(calendar.getTime()); //take number of days

            String cur = firstRow[detailIndex[6]].getContents();

            String email = "";
            String pathPDF = pathSave + "\\" + customer + "\\" + fileName +
                    ".pdf";

            //Console output
            textArea.append("Created " + pathPDF + "\n");
            poList.add(poNumber);
            emailList.add(email);
            pathFileList.add(pathPDF);
            fileNameList.add(fileName + ".pdf");
            customerList.add(customer);
            invoiceList.add(invoiceNum);
            nameList.add(employeeName);

            //Cleaning date into format of MM_YYYY
            String monthYear = shippingDate.split("_")[0] + '_' + shippingDate.split("_")[2];
            dateList.add(monthYear);


            int countTable = 0; //Find 4 in order to add item information in table.
            double totalAmount = 0;
            for (XWPFTable tableDoc : newDoc.getTables()) {
                if (countTable == 4) {
                    XWPFTableRow oldRow = tableDoc.getRow(1); //The first empty row to copy
                    for (int i = 0; i < intWithSamePo.size(); i++) {
                        int rowNum = intWithSamePo.get(i);
                        Cell[] rowData = sheet.getRow(rowNum);
                        String[] items = new String[7];
                        for (int j = 0; j < detailIndex.length; j++) {
                            int index = detailIndex[j];
                            if (j == 4 || j == 5) {
                                items[j] = rowData[index].getContents().replaceAll(",", "");
                                if (j == 5) {
                                    totalAmount += Double.parseDouble(items[j]);
                                }
                            } else if (j == 0) { //Remove decimal for quantity
                                items[j] = rowData[index].getContents().replaceAll("\\.0*$", "");
                            } else {
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
                                    run.setText(invoiceNum, 0);
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
        trackInvoice.recordInvoice(pathSave, poList, emailList, pathFileList, fileNameList, customerList, invoiceList
                , nameList, dateList);

        //System.exit(1);
    }

}
