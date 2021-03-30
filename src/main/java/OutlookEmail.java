import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;
import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.*;
import javax.swing.*;

import java.awt.Desktop;
import java.io.IOException;
import java.net.URI;
import java.net.URISyntaxException;

public class OutlookEmail {
    private String id;
    private String pw;
    private Session session;
    private String subject;
    private String bodyMessage;
    private String pathTrack;
    private String pathSave;
    private String termPath;
    private JTextArea textArea;
    private int count = 0;
    public OutlookEmail(String id, String pw, JTextArea textArea) {
        this.id = id;
        this.pw = pw;
        this.textArea = textArea;
        final String username = id;
        final String password = pw;
        Properties props = new Properties();
        props.put("mail.smtp.auth", "true");
        props.put("mail.smtp.starttls.enable", "true");
        props.put("mail.smtp.host", "outlook.office365.com");
        props.put("mail.smtp.port", "587");

        //Checking the password
        session = Session.getInstance(props,
                new javax.mail.Authenticator() {
                    protected PasswordAuthentication getPasswordAuthentication() {
                        return new PasswordAuthentication(username, password);
                    }
                });
    }

    public void setSubject(String subject) {
        this.subject = subject;
    }

    public void setBodyMessage(String bodyMessage) {
        this.bodyMessage = bodyMessage;
    }

    public int getCount() {
        return count;
    }

    public void updateCount() {
        File fileToBeModified = new File(pathSave + "\\" + "countSentEmails.txt");
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

    public void sendEmail(String pathSave) throws IOException, BiffException {
        this.pathSave = pathSave;
        this.pathTrack = pathSave + "\\" + "trackMail.xlsx";
        termPath = pathSave + "\\" + "Navien Invoice Terms.pdf";
        //obtaining input bytes from a file
        FileInputStream input = new FileInputStream(new File(pathTrack));
        //creating workbook instance that refers to .xls file
        XSSFWorkbook workbook = new XSSFWorkbook(input);
        //creating a Sheet object to retrieve the object
        XSSFSheet sheet = workbook.getSheetAt(0);

        //Find indices of same customer
        HashMap<String, ArrayList<Integer>> customerAndIndex = new HashMap<>();
        for(int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
            Row row = sheet.getRow(i);
            String po = row.getCell(0).getStringCellValue();
            String fileName = row.getCell(3).getStringCellValue();
            String customer = row.getCell(5).getStringCellValue();
            if (row.getCell(0) == null) {
                break;
            } else if (row.getCell(4) != null) { // 4: If email is sent
                continue;
            } else if (row.getCell(1) == null) { // 1: If there is no email
                textArea.append("\n" + customer + " with " + fileName + " has no email.\n");
                continue;
            }
            if (customerAndIndex.containsKey(customer)) {
                ArrayList<Integer> temp = customerAndIndex.get(customer);
                temp.add(i);
            } else {
                ArrayList<Integer> temp = new ArrayList<>();
                temp.add(i);
                customerAndIndex.put(customer, temp);
            }
        }

        for (Map.Entry<String, ArrayList<Integer>> entry : customerAndIndex.entrySet()) {
            String nameCustomer = entry.getKey();
            ArrayList<Integer> indicesExcel = entry.getValue();
            send(indicesExcel);
        }

//        for(int i = 1; i < sheet.getPhysicalNumberOfRows(); i++) {
//            Row row = sheet.getRow(i);
//            if (row.getCell(0) == null) {
//                break;
//            } else if (row.getCell(4) != null || row.getCell(1) == null) { //4: If email is sent, 1: If there is no
//                // email
//                continue;
//            }
//            String po = row.getCell(0).getStringCellValue();
//            String email = row.getCell(1).getStringCellValue();
//            String filePath = row.getCell(2).getStringCellValue();
//            String fileName = row.getCell(3).getStringCellValue();
//            String customer = row.getCell(5).getStringCellValue();
//            try {
//                if (email.equals(" ")) {
//                    throw new Exception();
//                }
//
//                for ()
//
//                //Sending an email to corresponding information
//                sendToBuyer(email, filePath, fileName);
//                Cell cellSent = row.createCell(4);
//                cellSent.setCellValue("Yes");
//                //System.out.println(po + " " + email + " " + filePath + " " + fileName);
//                count++;
//            } catch (Exception e) {
//                textArea.append(fileName + " " + "has no email of the buyer.\n");
//            }
//        }
//        try {
//            FileOutputStream fileOutputStream = new FileOutputStream(pathTrack);
//            workbook.write(fileOutputStream);
//            fileOutputStream.close();
//            workbook.close();
//        } catch (Exception e) {
//            textArea.append("Please close the trackMail Excel file.\n");
//        }
    }

    public void send(ArrayList<Integer> indice) throws IOException {
        //obtaining input bytes from a file
        FileInputStream input = new FileInputStream(new File(pathTrack));
        //creating workbook instance that refers to .xls file
        XSSFWorkbook workbook = new XSSFWorkbook(input);
        //creating a Sheet object to retrieve the object
        XSSFSheet sheet = workbook.getSheetAt(0);
        try {
            int firstOne = indice.get(0);
            Row row = sheet.getRow(firstOne);
            String email = row.getCell(1).getStringCellValue();
            String customer = row.getCell(5).getStringCellValue();

            //Create message envelope.
            MimeMessage msg = new MimeMessage(session);
            msg.addFrom(InternetAddress.parse(id));
            msg.setRecipients(Message.RecipientType.TO,
                    InternetAddress.parse(email));
            msg.setRecipients(Message.RecipientType.CC,
                    InternetAddress.parse("tokipul5@berkeley.edu"));
            //msg.setHeader("X-Unsent", "1");

            MimeMultipart mmp = new MimeMultipart();
            MimeBodyPart body = new MimeBodyPart();
            body.setDisposition(MimePart.INLINE);
            body.setContent(bodyMessage, "text/plain");
            mmp.addBodyPart(body);

            //Terms
            MimeBodyPart attTerm = new MimeBodyPart();
            attTerm.attachFile(termPath);
            mmp.addBodyPart(attTerm);

            String subject = "Invoice ";
            for (int index : indice) {
                row = sheet.getRow(index);
                String filePath = row.getCell(2).getStringCellValue();
                String invoiceNum = row.getCell(6).getStringCellValue();
                subject += invoiceNum + "/";
                attTerm = new MimeBodyPart();
                attTerm.attachFile(filePath);
                mmp.addBodyPart(attTerm);
            }
            subject = subject.substring(0, subject.length()-1);
            msg.setSubject(subject);

//            MimeBodyPart att = new MimeBodyPart();
//            att.attachFile(pathFile);
//            mmp.addBodyPart(att);

            msg.setContent(mmp);
            msg.saveChanges();


            File resultEmail = File.createTempFile("test", ".eml");
            try (FileOutputStream fs = new FileOutputStream(resultEmail)) {
                msg.writeTo(fs);
                fs.flush();
                fs.getFD().sync();
            }

            System.out.println(resultEmail.getCanonicalPath());

            ProcessBuilder pb = new ProcessBuilder();
            pb.command("cmd.exe", "/C", "start", "outlook.exe",
                    "/eml", resultEmail.getCanonicalPath());
            Process p = pb.start();


            for (int index : indice) {
                Row temp = sheet.getRow(index);
                Cell cellSent = temp.createCell(4);
                cellSent.setCellValue("Yes");
                //System.out.println(po + " " + email + " " + filePath + " " + fileName);
                count++;
            }

            try {
                p.waitFor();
            } finally {
                p.getErrorStream().close();
                p.getInputStream().close();
                p.getErrorStream().close();
                p.destroy();
            }

            try {
                FileOutputStream fileOutputStream = new FileOutputStream(pathTrack);
                workbook.write(fileOutputStream);
                fileOutputStream.close();
                workbook.close();
            } catch (Exception e) {
                textArea.append("\nPlease close the trackMail Excel file.\n");
            }

            textArea.append("Sent " + indice.size() + " attachments to " + customer + "\n");

        } catch (MessagingException | FileNotFoundException e) {
            throw new RuntimeException("Failed to send an email");
        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
        }
    }

    public void login() {
        try {
            Message message = new MimeMessage(session);
            message.setFrom(new InternetAddress(id));
            message.setRecipients(Message.RecipientType.TO,
                    InternetAddress.parse(id));
            message.setSubject("Signed in successfully in Navien Invoice Program");
            //message.setText("HI");

            // Create the message part
            BodyPart messageBodyPart = new MimeBodyPart();

            // Now set the actual message
            messageBodyPart.setText("Please change your password if you did not use the program.");

            // Create a multipart message
            Multipart multipart = new MimeMultipart();

            // Set text message part
            multipart.addBodyPart(messageBodyPart);

            // Send the complete message parts
            message.setContent(multipart);

            Transport.send(message);
        } catch (MessagingException e) {
            textArea.append("\n***Please check your username and password.***\n");
            throw new RuntimeException("Wrong username or password");
        }
        textArea.append("\nLogin in successfully!\n");
    }

    public void sendToBuyer(String buyerEmail, String pathFile, String pdfName) {
        try {
            //Create message envelope.
            MimeMessage msg = new MimeMessage(session);
            msg.addFrom(InternetAddress.parse(id));
            msg.setRecipients(Message.RecipientType.TO,
                    InternetAddress.parse(buyerEmail));
            msg.setRecipients(Message.RecipientType.CC,
                    InternetAddress.parse("tokipul5@berkeley.edu"));
            msg.setSubject(subject);
            //msg.setHeader("X-Unsent", "1");

            MimeMultipart mmp = new MimeMultipart();
            MimeBodyPart body = new MimeBodyPart();
            body.setDisposition(MimePart.INLINE);
            body.setContent(bodyMessage, "text/plain");
            mmp.addBodyPart(body);

            //Terms
            MimeBodyPart attTerm = new MimeBodyPart();
            attTerm.attachFile(termPath);
            mmp.addBodyPart(attTerm);
            attTerm = new MimeBodyPart();
            attTerm.attachFile(pathFile);
            mmp.addBodyPart(attTerm);

//            MimeBodyPart att = new MimeBodyPart();
//            att.attachFile(pathFile);
//            mmp.addBodyPart(att);

            msg.setContent(mmp);
            msg.saveChanges();


            File resultEmail = File.createTempFile("test", ".eml");
            try (FileOutputStream fs = new FileOutputStream(resultEmail)) {
                msg.writeTo(fs);
                fs.flush();
                fs.getFD().sync();
            }

            System.out.println(resultEmail.getCanonicalPath());

            ProcessBuilder pb = new ProcessBuilder();
            pb.command("cmd.exe", "/C", "start", "outlook.exe",
                    "/eml", resultEmail.getCanonicalPath());
            Process p = pb.start();
            try {
                p.waitFor();
            } finally {
                p.getErrorStream().close();
                p.getInputStream().close();
                p.getErrorStream().close();
                p.destroy();
            }

            textArea.append("Sent " + pdfName + "\n");

        } catch (MessagingException | FileNotFoundException e) {
            throw new RuntimeException("Failed to send an email");
        } catch (IOException | InterruptedException e) {
            e.printStackTrace();
        }
    }
}

//            Message message = new MimeMessage(session);
//            message.setFrom(new InternetAddress(id));
//            message.setRecipients(Message.RecipientType.TO,
//                    InternetAddress.parse(buyerEmail));
//            message.setSubject(subject);
//            //message.setText("HI");
//
//            // Create the message part
//            BodyPart messageBodyPart = new MimeBodyPart();
//
//            // Now set the actual message
//            messageBodyPart.setText(bodyMessage);
//
//            // Create a multipart message
//            Multipart multipart = new MimeMultipart();
//
//            // Set text message part
//            multipart.addBodyPart(messageBodyPart);
//
//            // Part two is attachment
//            messageBodyPart = new MimeBodyPart();
//            String filename = pathFile;
//            DataSource source = new FileDataSource(filename);
//            messageBodyPart.setDataHandler(new DataHandler(source));
//            messageBodyPart.setFileName(pdfName);
//            multipart.addBodyPart(messageBodyPart);
//
//            messageBodyPart = new MimeBodyPart();
//            filename = "C:\\Users\\Keeyou\\Downloads\\sample\\Navien Invoice Terms.pdf";
//            source = new FileDataSource(filename);
//            messageBodyPart.setDataHandler(new DataHandler(source));
//            messageBodyPart.setFileName("Navien Invoice Terms.pdf");
//            multipart.addBodyPart(messageBodyPart);
//
//            // Send the complete message parts
//            message.setContent(multipart);
//
//            //https://stackoverflow.com/questions/28471326/how-to-open-mail-in-draft-and-attach-file-to-mail-using-java
//            message.saveChanges();
//
//            FileOutputStream emailFile = new FileOutputStream("C:\\Users\\Keeyou\\Downloads\\sample\\email.eml");
//            message.writeTo(emailFile);
//            emailFile.flush();
//            emailFile.getFD().sync();
//
//            File resultEmail = File.createTempFile("test", ".eml");
//            //File resultEmail = new File("C:\\Users\\Keeyou\\Downloads\\sample" + filename + ".eml");
//            try (FileOutputStream fs = new FileOutputStream(resultEmail)) {
//                message.writeTo(fs);
//                fs.flush();
//                fs.getFD().sync();
//            }
//
//            System.out.println(resultEmail.getCanonicalPath());
//
//            ProcessBuilder pb = new ProcessBuilder();
//            pb.command("cmd.exe", "/C", "start", "outlook.exe",
//                    "/eml", resultEmail.getCanonicalPath());
//            Process p = pb.start();
//            try {
//                p.waitFor();
//            } catch (InterruptedException e) {
//                e.printStackTrace();
//            } finally {
//                p.getErrorStream().close();
//                p.getInputStream().close();
//                p.getErrorStream().close();
//                p.destroy();
//            }
//
//            File email = new File("C:\\Users\\Keeyou\\Downloads\\sample\\email.eml");
//            //Desktop.getDesktop().open(email);
//            //Desktop.getDesktop().edit(email);
//            System.out.println(email.getCanonicalPath());
//
//            ProcessBuilder processBuilder = new ProcessBuilder();
//            processBuilder.command("C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.exe",
//                    "/m", "keeyoukim@gmail.com",
//                    //"subject=Invoice&body=Test Body",
//                    //"/a", pathFile,
//                    "/a", resultEmail.getCanonicalPath());
//                    //"/eml", email.getCanonicalPath()
//                    //"/a", pathFile
//                    //"/c", "imp.note",
//                    //"/m", buyerEmail + "?subject=" + subject + "&body=" + bodyMessage,
//            processBuilder.start();


//Transport.send(message);